"""Create and Ed25519-sign the package manifest uploaded with a GitHub release."""

import base64
import hashlib
import json
import os
import platform
import sys
from pathlib import Path

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey


def platform_id():
    system = platform.system().lower()
    machine = platform.machine().lower()
    machine = {"amd64": "x64", "x86_64": "x64", "aarch64": "arm64"}.get(
        machine, machine)
    prefix = {"windows": "windows", "darwin": "macos", "linux": "linux"}.get(system)
    if not prefix or machine not in ("x64", "arm64"):
        raise RuntimeError(f"Unsupported release platform: {system}-{machine}")
    return f"{prefix}-{machine}"


def main():
    if len(sys.argv) != 3:
        raise SystemExit("usage: create_update_manifest.py <version> <artifact-directory>")
    version = sys.argv[1].removeprefix("v")
    artifact_dir = Path(sys.argv[2]).resolve()
    repository = os.environ.get("GITHUB_REPOSITORY", "bkoustoov-oss/ModelAI")
    signing_key = os.environ.get("LAYA_UPDATE_SIGNING_KEY_B64", "")
    if not signing_key:
        raise RuntimeError("Set the LAYA_UPDATE_SIGNING_KEY_B64 GitHub Actions secret.")
    private_bytes = base64.b64decode(signing_key, validate=True)
    if len(private_bytes) != 32:
        raise RuntimeError("The Ed25519 signing secret must encode exactly 32 private-key bytes.")
    private_key = Ed25519PrivateKey.from_private_bytes(private_bytes)

    platform_packages = {
        "windows-x64": "LAYA-Windows-x64.zip",
        "macos-x64": "LAYA-macOS-x64.zip",
        "macos-arm64": "LAYA-macOS-arm64.zip",
        "linux-x64": "LAYA-Linux-x64.tar.gz",
    }
    packages = {}
    for platform_name, filename in platform_packages.items():
        path = artifact_dir / filename
        if not path.is_file():
            continue
        digest_builder = hashlib.sha256()
        with open(path, "rb") as package_file:
            for block in iter(lambda: package_file.read(1024 * 1024), b""):
                digest_builder.update(block)
        packages[platform_name] = {
            "name": filename,
            "sha256": digest_builder.hexdigest(),
            "size": path.stat().st_size,
        }
    if not packages:
        raise RuntimeError(f"No update packages found in {artifact_dir}")
    if platform_id() not in packages:
        raise RuntimeError(f"No update bundle produced for this runner ({platform_id()}).")

    manifest = {
        "app": "laya-agentic-studio",
        "repository": repository,
        "version": version,
        "packages": packages,
    }
    raw = (json.dumps(manifest, sort_keys=True, separators=(",", ":")) + "\n").encode()
    signature = private_key.sign(raw)
    (artifact_dir / "update-manifest.json").write_bytes(raw)
    (artifact_dir / "update-manifest.sig").write_text(
        base64.b64encode(signature).decode("ascii") + "\n", encoding="ascii")
    print(f"Signed update manifest for {version}: {', '.join(sorted(packages))}")


if __name__ == "__main__":
    main()

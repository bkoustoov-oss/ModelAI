"""Signature-verified GitHub Release updates for packaged Agentic Studio builds."""

import base64
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
from packaging.version import InvalidVersion, Version
from platformdirs import user_data_dir

REPOSITORY = "bkoustoov-oss/ModelAI"
APP_SLUG = "laya-agentic-studio"
UPDATE_PUBLIC_KEY_B64 = "/5lNwKma58jLEJppYXYCy1Qq5TBzJGP8pePR2cgfF4M="
USER_AGENT = "LAYA-Agentic-Studio-Updater"


class UpdateError(RuntimeError):
    pass


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    platform: str
    asset_name: str
    asset_url: str
    sha256: str
    size: int
    release_notes: str


def platform_id(system=None, machine=None):
    system = (system or platform.system()).lower()
    machine = (machine or platform.machine()).lower()
    machine = {"amd64": "x64", "x86_64": "x64", "aarch64": "arm64"}.get(
        machine, machine)
    system_name = {"windows": "windows", "darwin": "macos", "linux": "linux"}.get(system)
    if system_name is None or machine not in ("x64", "arm64"):
        raise UpdateError(f"Unsupported update platform: {system}-{machine}")
    return f"{system_name}-{machine}"


def _get_bytes(url, accept="application/octet-stream", timeout=20):
    request = urllib.request.Request(
        url, headers={"User-Agent": USER_AGENT, "Accept": accept})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def _verify_manifest(raw_manifest, raw_signature):
    try:
        public_key = Ed25519PublicKey.from_public_bytes(
            base64.b64decode(UPDATE_PUBLIC_KEY_B64, validate=True))
        signature = base64.b64decode(raw_signature.strip(), validate=True)
        public_key.verify(signature, raw_manifest)
        manifest = json.loads(raw_manifest.decode("utf-8"))
    except Exception as exc:
        raise UpdateError(f"Release manifest signature is invalid: {exc}") from exc
    if manifest.get("app") != APP_SLUG:
        raise UpdateError("Release manifest belongs to a different application.")
    return manifest


def check_for_update(current_version, repository=REPOSITORY):
    release_url = f"https://api.github.com/repos/{repository}/releases/latest"
    try:
        release = json.loads(_get_bytes(
            release_url, accept="application/vnd.github+json").decode("utf-8"))
    except Exception as exc:
        raise UpdateError(f"Could not check GitHub Releases: {exc}") from exc

    assets = {item.get("name"): item.get("browser_download_url")
              for item in release.get("assets", [])}
    manifest_url = assets.get("update-manifest.json")
    signature_url = assets.get("update-manifest.sig")
    if not manifest_url or not signature_url:
        raise UpdateError("The latest release does not contain a signed update manifest.")

    manifest = _verify_manifest(_get_bytes(manifest_url), _get_bytes(signature_url))
    latest = str(manifest.get("version", ""))
    try:
        if Version(latest) <= Version(current_version):
            return None
    except InvalidVersion as exc:
        raise UpdateError(f"Release version is invalid: {latest}") from exc

    target_platform = platform_id()
    package = manifest.get("packages", {}).get(target_platform)
    if not package:
        raise UpdateError(f"Release {latest} has no package for {target_platform}.")
    asset_name = package.get("name", "")
    asset_url = assets.get(asset_name) or package.get("url")
    sha256 = str(package.get("sha256", "")).lower()
    if not isinstance(asset_url, str) or not asset_url.startswith("https://"):
        raise UpdateError("Update package URL must use HTTPS.")
    if len(sha256) != 64 or any(char not in "0123456789abcdef" for char in sha256):
        raise UpdateError("Update manifest contains an invalid SHA-256 digest.")
    return UpdateInfo(
        version=latest,
        platform=target_platform,
        asset_name=asset_name,
        asset_url=asset_url,
        sha256=sha256,
        size=int(package.get("size", 0)),
        release_notes=str(release.get("body", "")))


def download_package(update, progress_callback=None):
    updates_dir = Path(user_data_dir("LAYA Agentic Studio", "Bkoustoov")) / "updates"
    updates_dir.mkdir(parents=True, exist_ok=True)
    destination = updates_dir / update.asset_name
    request = urllib.request.Request(
        update.asset_url, headers={"User-Agent": USER_AGENT,
                               "Accept": "application/octet-stream"})
    digest = hashlib.sha256()
    total = 0
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            expected_size = int(response.headers.get("Content-Length") or update.size or 0)
            with open(destination, "wb") as output:
                while True:
                    block = response.read(1024 * 1024)
                    if not block:
                        break
                    output.write(block)
                    digest.update(block)
                    total += len(block)
                    if progress_callback:
                        progress_callback(total, expected_size)
    except Exception as exc:
        destination.unlink(missing_ok=True)
        raise UpdateError(f"Update download failed: {exc}") from exc
    if digest.hexdigest() != update.sha256:
        destination.unlink(missing_ok=True)
        raise UpdateError("Downloaded update failed SHA-256 verification.")
    if update.size and total != update.size:
        destination.unlink(missing_ok=True)
        raise UpdateError("Downloaded update size does not match the signed manifest.")
    return destination


def _application_install_root():
    if not getattr(sys, "frozen", False):
        raise UpdateError("OTA installation is available only in an installed build.")
    executable = Path(sys.executable).resolve()
    if sys.platform == "darwin":
        for parent in executable.parents:
            if parent.suffix.lower() == ".app":
                return parent
    return executable.parent


def launch_installer_update(update, archive_path, parent_pid):
    install_root = _application_install_root()
    executable = Path(sys.executable).resolve()
    helper_name = "update_helper.exe" if os.name == "nt" else "update_helper"
    candidates = [executable.parent / helper_name]
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        candidates.append(Path(bundle_root) / helper_name)
    helper_source = next((path for path in candidates if path.is_file()), None)
    if helper_source is None:
        raise UpdateError("Update helper is missing from this installation.")

    helper_dir = Path(user_data_dir("LAYA Agentic Studio", "Bkoustoov")) / "updates"
    helper_dir.mkdir(parents=True, exist_ok=True)
    helper_copy = helper_dir / helper_name
    shutil.copy2(helper_source, helper_copy)
    if os.name != "nt":
        helper_copy.chmod(0o700)
    command = [str(helper_copy), "--pid", str(parent_pid),
               "--archive", str(archive_path),
               "--install-root", str(install_root),
               "--executable", str(executable),
               "--version", update.version]
    options = {"close_fds": True}
    if os.name == "nt":
        options["creationflags"] = (getattr(subprocess, "DETACHED_PROCESS", 0x00000008) |
                                    getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200))
    else:
        options["start_new_session"] = True
    subprocess.Popen(command, **options)
    return install_root

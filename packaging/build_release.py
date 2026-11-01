"""Build a platform-specific LAYA application bundle and update archive."""

import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import zipfile
import argparse
import ast
import inspect
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
BUILD = ROOT / "build"
APP_NAME = "LAYAAgenticStudio"
HELPER_NAME = "update_helper.exe" if os.name == "nt" else "update_helper"
COLLECT_ALL = [
    "customtkinter", "laya", "ollama", "openai", "keyring", "PIL",
    "pytesseract", "speech_recognition", "paramiko", "pdfplumber",
    "pandas", "matplotlib", "pyttsx3", "reportlab", "platformdirs",
    "cryptography", "packaging", "torch", "transformers",
    "safetensors", "huggingface_hub", "numpy",
]
HIDDEN_IMPORTS = ["speech_recognition"]
if os.name == "nt":
    HIDDEN_IMPORTS.extend(["win32com.client", "pyttsx3.drivers.sapi5"])
elif sys.platform == "darwin":
    HIDDEN_IMPORTS.append("pyttsx3.drivers.nsss")
else:
    HIDDEN_IMPORTS.append("pyttsx3.drivers.espeak")


def run_pyinstaller(script, name, dist_path, windowed, onefile=False,
                    collect_app_dependencies=True):
    import PyInstaller.__main__

    args = [
        str(script), "--noconfirm", "--clean",
        "--name", name, "--distpath", str(dist_path),
        "--workpath", str(BUILD / f"pyinstaller-{name}"),
        "--specpath", str(BUILD),
    ]
    args.append("--onefile" if onefile else "--onedir")
    if windowed:
        args.append("--windowed")
    if collect_app_dependencies:
        for module in COLLECT_ALL:
            args.extend(["--collect-all", module])
        for module in HIDDEN_IMPORTS:
            args.extend(["--hidden-import", module])
    PyInstaller.__main__.run(args)


def set_build_version(version):
    version = version.removeprefix("v")
    if not re.fullmatch(r"\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?", version):
        raise ValueError(f"Invalid release version: {version}")
    (ROOT / "app_version.py").write_text(
        f'APP_VERSION = "{version}"\n', encoding="utf-8")
    return version


def validate_scrollable_frame_api():
    """Fail the build if the app passes options unsupported by this CTk version."""
    import customtkinter as ctk

    source_path = ROOT / "laya_agentic_studio_v15.py"
    syntax = ast.parse(source_path.read_text(encoding="utf-8"), filename=str(source_path))
    parameters = inspect.signature(
        ctk.CTkScrollableFrame.__init__).parameters
    supported = {
        name for name, parameter in parameters.items()
        if name not in ("self", "master")
        and parameter.kind not in (inspect.Parameter.VAR_POSITIONAL,
                                   inspect.Parameter.VAR_KEYWORD)
    }
    unsupported = []
    for node in ast.walk(syntax):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Attribute):
            continue
        if node.func.attr != "CTkScrollableFrame":
            continue
        for keyword in node.keywords:
            if keyword.arg is not None and keyword.arg not in supported:
                unsupported.append((node.lineno, keyword.arg))
    if unsupported:
        details = ", ".join(
            f"line {line}: {name}" for line, name in unsupported)
        raise RuntimeError(
            "The application uses CTkScrollableFrame options unsupported "
            f"by CustomTkinter {ctk.__version__}: {details}")
    call_count = sum(
        1 for node in ast.walk(syntax)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "CTkScrollableFrame")
    print(
        f"CustomTkinter {ctk.__version__} scroll-frame API check passed "
        f"({call_count} call sites).")


def clean_platform_outputs(system):
    """Remove only this builder's outputs so packages cannot contain stale files."""
    for path in (
            BUILD / "app-build",
            BUILD / "helper-dist",
            DIST / APP_NAME):
        shutil.rmtree(path, ignore_errors=True)

    generated_files = {
        "windows": ("LAYA-Windows-x64.zip", "LAYA-Windows-Setup.exe"),
        "darwin": tuple(
            f"LAYA-macOS-{architecture}.{extension}"
            for architecture in ("x64", "arm64")
            for extension in ("zip", "dmg")),
        "linux": ("LAYA-Linux-x64.tar.gz",),
    }.get(system, ())
    for filename in generated_files:
        (DIST / filename).unlink(missing_ok=True)


def package_windows(version, app_dir, portable_only=False):
    archive = DIST / "LAYA-Windows-x64.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as output:
        for path in app_dir.rglob("*"):
            if path.is_file():
                output.write(path, path.relative_to(DIST))
    installer_script = ROOT / "packaging" / "windows.iss"
    iscc = shutil.which("ISCC.exe")
    if iscc is None:
        candidates = [Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")) /
                      "Inno Setup 6" / "ISCC.exe",
                      Path(os.environ.get("PROGRAMFILES", r"C:\Program Files")) /
                      "Inno Setup 6" / "ISCC.exe"]
        iscc = next((str(path) for path in candidates if path.is_file()), None)
    if iscc is None and portable_only:
        print("Inno Setup not found; built the portable ZIP only.")
        return
    if iscc is None:
        raise RuntimeError("Inno Setup compiler (ISCC.exe) was not found.")
    subprocess.run([iscc, f"/DAppVersion={version}", str(installer_script)],
                   cwd=ROOT, check=True)
    shutil.rmtree(app_dir, ignore_errors=True)


def package_macos(version, app_bundle):
    architecture = "arm64" if platform.machine().lower() in ("arm64", "aarch64") else "x64"
    archive = DIST / f"LAYA-macOS-{architecture}.zip"
    subprocess.run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent",
                    str(app_bundle), str(archive)], check=True)
    staging = BUILD / "dmg-staging"
    shutil.rmtree(staging, ignore_errors=True)
    staging.mkdir(parents=True)
    shutil.copytree(app_bundle, staging / app_bundle.name)
    install_command = ROOT / "packaging" / "install_macos.command"
    shutil.copy2(install_command, staging / install_command.name)
    (staging / install_command.name).chmod(0o755)
    (staging / "Applications").symlink_to("/Applications", target_is_directory=True)
    subprocess.run([
        "hdiutil", "create", "-volname", "LAYA Agentic Studio",
        "-srcfolder", str(staging), "-ov", "-format", "UDZO",
        str(DIST / f"LAYA-macOS-{architecture}.dmg")], check=True)


def package_linux(version, app_dir):
    archive = DIST / "LAYA-Linux-x64.tar.gz"
    installer = ROOT / "packaging" / "install_linux.sh"
    with tarfile.open(archive, "w:gz") as output:
        for path in app_dir.iterdir():
            output.add(path, arcname=path.name)
        output.add(installer, arcname="install_linux.sh")
    shutil.rmtree(app_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("version", nargs="?",
                        default=os.environ.get("GITHUB_REF_NAME", "v0.1.0"))
    parser.add_argument("--portable-only", action="store_true",
                        help="Skip the Windows installer compiler for local smoke builds.")
    args = parser.parse_args()
    validate_scrollable_frame_api()
    version = set_build_version(args.version)
    system = platform.system().lower()
    if system not in ("windows", "darwin", "linux"):
        raise RuntimeError(f"Unsupported build system: {system}")
    DIST.mkdir(parents=True, exist_ok=True)
    BUILD.mkdir(parents=True, exist_ok=True)
    clean_platform_outputs(system)

    helper_dist = BUILD / "helper-dist"
    run_pyinstaller(ROOT / "update_helper.py", "update_helper",
                    helper_dist, windowed=False, onefile=True,
                    collect_app_dependencies=False)
    helper_binary = helper_dist / HELPER_NAME
    app_dist = BUILD / "app-build"
    run_pyinstaller(ROOT / "laya_agentic_studio_v15.py", APP_NAME,
                    app_dist, windowed=True)

    if system == "darwin":
        app_bundle = app_dist / f"{APP_NAME}.app"
        helper_target = app_bundle / "Contents" / "MacOS" / HELPER_NAME
        shutil.copy2(helper_binary, helper_target)
        package_macos(version, app_bundle)
    elif system == "windows":
        app_dir = app_dist / APP_NAME
        shutil.copy2(helper_binary, app_dir / HELPER_NAME)
        shutil.copytree(app_dir, DIST / APP_NAME, dirs_exist_ok=True)
        package_windows(version, DIST / APP_NAME,
                portable_only=args.portable_only)
    elif system == "linux":
        app_dir = app_dist / APP_NAME
        shutil.copy2(helper_binary, app_dir / HELPER_NAME)
        shutil.copytree(app_dir, DIST / APP_NAME, dirs_exist_ok=True)
        package_linux(version, DIST / APP_NAME)
    else:
        raise RuntimeError(f"Unsupported build system: {system}")

    print(f"Built LAYA Agentic Studio {version} for {system} in {DIST}")


if __name__ == "__main__":
    main()

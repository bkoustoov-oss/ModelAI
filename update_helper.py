"""Install a verified Agentic Studio update after the running app exits."""

import argparse
import os
import shutil
import subprocess
import sys
import tarfile
import time
import zipfile
from pathlib import Path


def wait_for_process(pid):
    if os.name == "nt":
        import ctypes
        from ctypes import wintypes

        synchronize = 0x00100000
        handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, pid)
        if not handle:
            return
        try:
            ctypes.windll.kernel32.WaitForSingleObject(handle, wintypes.DWORD(0xFFFFFFFF))
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)
        return

    while True:
        try:
            os.kill(pid, 0)
        except ProcessLookupError:
            return
        except PermissionError:
            return
        time.sleep(0.25)


def safe_destination(root, member_name):
    destination = (root / member_name).resolve()
    if not destination.is_relative_to(root.resolve()):
        raise RuntimeError(f"Unsafe archive path: {member_name}")
    return destination


def extract_archive(archive_path, destination):
    archive_path = Path(archive_path)
    if zipfile.is_zipfile(archive_path):
        with zipfile.ZipFile(archive_path) as package:
            for member in package.infolist():
                safe_destination(destination, member.filename)
            package.extractall(destination)
        return

    if tarfile.is_tarfile(archive_path):
        with tarfile.open(archive_path, "r:*") as package:
            for member in package.getmembers():
                if member.issym() or member.islnk() or member.isdev():
                    raise RuntimeError(f"Unsupported link/device in archive: {member.name}")
                safe_destination(destination, member.name)
            package.extractall(destination)
        return

    raise RuntimeError("Update archive must be a ZIP or TAR file.")


def wait_for_install_root(staging):
    entries = list(staging.iterdir())
    if len(entries) == 1 and entries[0].is_dir():
        return entries[0]
    return staging


def install_update(args):
    target = Path(args.install_root).resolve()
    if not target.is_dir():
        raise RuntimeError(f"Application install folder does not exist: {target}")
    if not os.access(target.parent, os.W_OK):
        raise RuntimeError(f"No write permission for application folder: {target.parent}")

    wait_for_process(args.pid)
    staging = target.parent / f".{target.name}.update-staging"
    backup = target.parent / f".{target.name}.update-backup"
    shutil.rmtree(staging, ignore_errors=True)
    shutil.rmtree(backup, ignore_errors=True)
    staging.mkdir(parents=True)
    extract_archive(args.archive, staging)
    payload = wait_for_install_root(staging)

    old_executable = Path(args.executable).resolve()
    try:
        executable_relative = old_executable.relative_to(target)
    except ValueError as exc:
        raise RuntimeError("Application executable is outside its install directory.") from exc
    replacement_executable = payload / executable_relative
    if not replacement_executable.is_file():
        raise RuntimeError("Update package does not contain the application executable.")

    try:
        target.rename(backup)
        payload.rename(target)
    except Exception:
        if not target.exists() and backup.exists():
            backup.rename(target)
        raise

    updated_executable = target / executable_relative
    subprocess.Popen([str(updated_executable)], cwd=str(target), close_fds=True)
    shutil.rmtree(backup, ignore_errors=True)
    shutil.rmtree(staging, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pid", required=True, type=int)
    parser.add_argument("--archive", required=True)
    parser.add_argument("--install-root", required=True)
    parser.add_argument("--executable", required=True)
    parser.add_argument("--version", required=True)
    args = parser.parse_args()
    try:
        install_update(args)
    except Exception as exc:
        error_path = Path(args.archive).with_suffix(".update-error.txt")
        error_path.write_text(f"Update to {args.version} failed:\n{exc}\n", encoding="utf-8")
        raise


if __name__ == "__main__":
    main()

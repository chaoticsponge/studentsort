"""Public GitHub release checks. No spreadsheet data is sent or read here."""

import ctypes
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
from urllib.error import HTTPError
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener


VERSION = "1.0.2"
REPOSITORY = "chaoticsponge/studentsort"
RELEASES_URL = f"https://github.com/{REPOSITORY}/releases/latest"
ASSET_NAME = "ExamAccessLogGenerator.exe"
MAX_DOWNLOAD = 150 * 1024 * 1024


def version_tuple(value):
    if not isinstance(value, str) or not re.fullmatch(r"v?\d+\.\d+\.\d+", value):
        raise ValueError("The release has an unsupported version number.")
    return tuple(int(part) for part in value.lstrip("v").split("."))


class GitHubRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        url = urlsplit(newurl)
        if url.scheme != "https" or url.hostname not in {
            "github.com", "api.github.com", "release-assets.githubusercontent.com",
            "objects.githubusercontent.com",
        }:
            raise ValueError("The update redirected outside GitHub.")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def open_url(url):
    request = Request(url, headers={"User-Agent": f"StudentSort/{VERSION}"})
    return build_opener(GitHubRedirects()).open(request, timeout=30)


def release_update(release):
    if release.get("draft") or release.get("prerelease"):
        return None
    tag = release.get("tag_name")
    if version_tuple(tag) <= version_tuple(VERSION):
        return None
    for asset in release.get("assets", []):
        if asset.get("name") != ASSET_NAME:
            continue
        expected_url = f"https://github.com/{REPOSITORY}/releases/download/{tag}/{ASSET_NAME}"
        if asset.get("browser_download_url") != expected_url:
            raise ValueError("The release download address is unexpected.")
        digest = asset.get("digest", "")
        if not isinstance(digest, str) or not re.fullmatch(r"sha256:[a-fA-F0-9]{64}", digest):
            raise ValueError("This release has no verified download checksum yet. Try again later.")
        size = asset.get("size", 0)
        if not isinstance(size, int) or not 0 < size <= MAX_DOWNLOAD:
            raise ValueError("The update has an unexpected size.")
        return {"version": tag, "url": expected_url, "sha256": digest[7:].lower(), "size": size}
    raise ValueError("The new Windows release is not ready to download yet.")


def check_for_update():
    try:
        with open_url(f"https://api.github.com/repos/{REPOSITORY}/releases/latest") as response:
            data = response.read(1024 * 1024 + 1)
    except HTTPError as exc:
        if exc.code == 404:
            return None
        if exc.code in (403, 429):
            raise ValueError("GitHub is limiting update checks. Please try again later.") from exc
        raise
    if len(data) > 1024 * 1024:
        raise ValueError("GitHub returned an unexpectedly large response.")
    return release_update(json.loads(data))


def download_update(update):
    folder = Path(tempfile.mkdtemp(prefix="studentsort-update-"))
    target = folder / ASSET_NAME
    try:
        checksum = hashlib.sha256()
        total = 0
        with open_url(update["url"]) as response, target.open("wb") as output:
            while chunk := response.read(1024 * 1024):
                total += len(chunk)
                if total > update["size"]:
                    raise ValueError("The download is larger than expected.")
                checksum.update(chunk)
                output.write(chunk)
        if total != update["size"] or checksum.hexdigest() != update["sha256"]:
            raise ValueError("The download could not be verified. Your app has not been changed.")
        with target.open("rb") as source:
            if source.read(2) != b"MZ":
                raise ValueError("The downloaded file is not a Windows executable.")
        return target
    except Exception:
        shutil.rmtree(folder)
        raise


def start_install(download):
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        raise ValueError("In-app installation is available in the Windows EXE only.")
    target = Path(sys.executable).resolve()
    # Check write access before closing the teacher's app.
    with tempfile.TemporaryFile(dir=target.parent):
        pass
    environment = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")
    subprocess.Popen(
        [str(download), "--apply-update", str(target), str(os.getpid())],
        env=environment,
        creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
    )


def apply_update(target, parent_pid):
    """Run inside the verified NEW EXE after the original instance closes."""
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        raise ValueError("Updates can only be installed by the Windows application.")
    target = Path(target).resolve()
    if target.suffix.lower() != ".exe" or target == Path(sys.executable).resolve():
        raise ValueError("Invalid update destination.")
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
    kernel.OpenProcess.restype = ctypes.c_void_p
    kernel.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    process = kernel.OpenProcess(0x00100000, False, int(parent_pid))
    if process:
        try:
            if kernel.WaitForSingleObject(process, 60000) != 0:
                raise ValueError("The app did not close. Close it and try updating again.")
        finally:
            kernel.CloseHandle(process)
    # Stage on the same volume, then replace atomically. Retry while Windows
    # releases the old PyInstaller bootloader's executable handle.
    fd, staged = tempfile.mkstemp(suffix=".exe", dir=target.parent)
    os.close(fd)
    try:
        shutil.copyfile(sys.executable, staged)
        for attempt in range(30):
            try:
                os.replace(staged, target)
                break
            except PermissionError:
                if attempt == 29:
                    raise
                time.sleep(1)
        subprocess.Popen([str(target)], env=dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1"))
    finally:
        Path(staged).unlink(missing_ok=True)

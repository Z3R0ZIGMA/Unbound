"""Check for updates: asks GitHub for Unbound's latest release, and on request
downloads Unbound.exe from it, checks it against the release's SHA256SUMS.txt
and puts it in place of the running Unbound.exe.

Only happens when the user clicks Check for updates. Nothing about the PC or
its mail is sent: just a request for the release list, with Unbound's version
as the user agent.

Windows lets a running program be renamed but not overwritten, so the new
EXE is downloaded next to it as "Unbound.exe.new", and swapped in when
Unbound closes or restarts: the running one becomes "Unbound.exe.old", which
the next start deletes. Swapping only then means a running Unbound never
starts its file reader (which is Unbound.exe started again) from the new
version.
"""

import hashlib
import json
import os
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass

from . import __version__

REPO = "Z3R0ZIGMA/Unbound"
LATEST_RELEASE_API = f"https://api.github.com/repos/{REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{REPO}/releases/latest"
DOWNLOAD_PREFIX = f"https://github.com/{REPO}/releases/download/"
EXE_NAME = "Unbound.exe"
SUMS_NAME = "SHA256SUMS.txt"


class UpdateError(Exception):
    """An update that couldn't be checked or used; the message is for the user."""


@dataclass(frozen=True)
class Release:
    version: tuple
    tag: str
    page_url: str
    exe_url: str
    sums_url: str

    @property
    def name(self):
        return ".".join(map(str, self.version))


def parse_version(text):
    """(1, 3, 5) from "v1.3.5" or "1.3.5"; None when it isn't a plain version."""
    m = re.fullmatch(r"[vV]?(\d+)\.(\d+)(?:\.(\d+))?", (text or "").strip())
    return tuple(int(part or 0) for part in m.groups()) if m else None


def is_newer(latest, current):
    return latest > current


def parse_latest(text):
    """The release in GitHub's "latest release" JSON, or None when it has no
    usable version, is a draft or pre-release, or lacks the EXE or its
    checksum file (a release must have both to be installed)."""
    data = json.loads(text)
    if not isinstance(data, dict) or data.get("draft") or data.get("prerelease"):
        return None
    tag = data.get("tag_name") or ""
    version = parse_version(tag)
    if version is None:
        return None
    urls = {}
    for asset in data.get("assets") or []:
        name, url = asset.get("name"), asset.get("browser_download_url") or ""
        # Only files from this project's own releases are ever downloaded
        if isinstance(name, str) and url.startswith(DOWNLOAD_PREFIX):
            urls[name.lower()] = url
    exe, sums = urls.get(EXE_NAME.lower()), urls.get(SUMS_NAME.lower())
    if not exe or not sums:
        return None
    return Release(version, tag, data.get("html_url") or RELEASES_PAGE, exe, sums)


def expected_hash(sums, name=EXE_NAME):
    """The SHA-256 listed for `name` in a SHA256SUMS.txt ("hash  name" per line)."""
    for line in sums.splitlines():
        parts = line.strip().split(None, 1)
        if (len(parts) == 2 and re.fullmatch(r"[0-9a-fA-F]{64}", parts[0])
                and parts[1].strip().lstrip("*").lower() == name.lower()):
            return parts[0].lower()
    return None


def _request(url):
    return urllib.request.Request(url, headers={
        "User-Agent": f"Unbound/{__version__}",
        "Accept": "application/vnd.github+json",
    })


def _open(url, timeout):
    try:
        return urllib.request.urlopen(_request(url), timeout=timeout)
    except urllib.error.HTTPError as e:
        if e.code in (403, 429):
            raise UpdateError("GitHub is limiting how often it can be asked. "
                              "Try again in an hour.") from e
        raise UpdateError(f"GitHub answered {e.code} {e.reason}.") from e
    except (urllib.error.URLError, OSError) as e:
        reason = getattr(e, "reason", e)
        raise UpdateError(f"Couldn't reach GitHub ({reason}). Check the internet connection.") from e


def get_latest(timeout=20):
    """Ask GitHub for the latest release. Raises UpdateError."""
    with _open(LATEST_RELEASE_API, timeout) as response:
        text = response.read().decode("utf-8")
    try:
        return parse_latest(text)
    except ValueError as e:
        raise UpdateError("GitHub's answer couldn't be read.") from e


def download(release, target, progress=None, cancelled=lambda: False, timeout=60):
    """Download the release's EXE to `target` and check it against the
    release's SHA256SUMS.txt. The file is deleted and UpdateError raised when
    it doesn't match, or when the download is cancelled or fails part-way."""
    with _open(release.sums_url, timeout) as response:
        sums = response.read().decode("utf-8", "replace")
    expected = expected_hash(sums)
    if not expected:
        raise UpdateError(f"{EXE_NAME} isn't listed in the release's {SUMS_NAME}.")
    digest = hashlib.sha256()
    try:
        with _open(release.exe_url, timeout) as response, open(target, "wb") as fh:
            total = int(response.headers.get("Content-Length") or 0)
            done = 0
            while True:
                if cancelled():
                    raise UpdateError("The download was cancelled.")
                chunk = response.read(256 * 1024)
                if not chunk:
                    break
                fh.write(chunk)
                digest.update(chunk)
                done += len(chunk)
                if progress and total:
                    progress(min(100, done * 100 // total))
        if digest.hexdigest() != expected:
            raise UpdateError("The download doesn't match the release's checksum, so it wasn't used.")
    except BaseException:
        _remove(target)
        raise


def running_exe():
    """The path of the running Unbound.exe; None when run from source."""
    return os.path.abspath(sys.executable) if getattr(sys, "frozen", False) else None


def staged_path(exe):
    return exe + ".new"


def install(new_exe, exe):
    """Put `new_exe` in place of `exe`, which may be running: the running one
    is renamed to "*.old" (Windows allows that) and the new one moved to its
    name. Undone if the move fails."""
    old = exe + ".old"
    _remove(old)
    os.replace(exe, old)
    try:
        os.replace(new_exe, exe)
    except OSError:
        os.replace(old, exe)
        raise


def start_new_instance(exe):
    """Start `exe` as a separate program, not as a part of this one. A
    one-file PyInstaller EXE passes settings to programs it starts so that
    copies of itself reuse its unpacked files; the new version must unpack
    its own (PYINSTALLER_RESET_ENVIRONMENT, PyInstaller 6.9 and later)."""
    import subprocess
    env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([exe], env=env, creationflags=flags, close_fds=True,
                     cwd=os.path.dirname(exe))


def clean_up(exe):
    """Remove what an earlier update left: the previous version (it was still
    running then) and a download that was never swapped in. Best effort."""
    if exe:
        for leftover in (exe + ".old", staged_path(exe)):
            _remove(leftover)


def _remove(path):
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
    except OSError:
        pass  # still in use: the next start tries again

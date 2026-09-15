"""
updater.py
----------
Lightweight self-updater for the packaged .exe, backed by GitHub Releases.
No extra pip dependency — uses only the Python standard library (urllib),
so it works inside a PyInstaller --onefile build without extra --collect-all.

HOW IT WORKS (like a mobile app update):
  1. check_for_update()  -> asks GitHub "what's your latest release?"
                             compares it to the version baked into this build.
  2. download_update()   -> downloads the new .exe (published as a Release
                             asset) to a temp file, reporting progress.
  3. apply_update_and_restart()
                          -> writes a tiny .bat "relauncher" that waits for
                             this exe to fully exit (Windows locks a running
                             .exe so it can't overwrite itself), deletes the
                             old .exe, moves the new one into its place,
                             restarts it, then deletes itself. Then this
                             process exits immediately so the lock releases.

CONFIG: set GITHUB_REPO below to "yourusername/yourrepo" once the repo
exists (see README.md's "Auto-Update Setup" section for the full walkthrough).
"""

import os
import sys
import json
import time
import subprocess
import urllib.request
import urllib.error

# --------------------------------------------------------------------------
# CONFIG — update this to your actual GitHub repo, e.g. "swiftprosys/tdm-billing-tool"
# --------------------------------------------------------------------------
GITHUB_REPO = "YOUR_GITHUB_USERNAME/YOUR_REPO_NAME"
GITHUB_API_LATEST = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
REQUEST_TIMEOUT = 10  # seconds


class UpdateError(Exception):
    pass


def _version_tuple(v: str):
    """'v1.2.3' or '1.2.3' -> (1, 2, 3), for safe numeric comparison."""
    v = v.strip().lstrip("vV")
    parts = []
    for p in v.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def is_frozen() -> bool:
    """True only when running as the PyInstaller-built .exe (not `python app.py`)."""
    return bool(getattr(sys, "frozen", False))


def check_for_update(current_version: str):
    """
    Returns a dict:
        {"available": True/False, "latest_version": "1.2.0",
         "download_url": "...", "notes": "..."}
    Raises UpdateError on any network/parse failure (caller should show a
    friendly message rather than a raw traceback).
    """
    if "YOUR_GITHUB_USERNAME" in GITHUB_REPO:
        raise UpdateError(
            "Updater isn't configured yet — set GITHUB_REPO in updater.py "
            "to your real GitHub repo (see README.md)."
        )

    req = urllib.request.Request(
        GITHUB_API_LATEST,
        headers={"Accept": "application/vnd.github+json", "User-Agent": "SPS-TDM-Updater"},
    )
    try:
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise UpdateError("No releases found yet on GitHub for this repo.")
        raise UpdateError(f"GitHub API error: HTTP {e.code}")
    except Exception as e:
        raise UpdateError(f"Could not reach GitHub: {e}")

    latest_tag = data.get("tag_name", "0.0.0")
    notes = data.get("body", "") or ""

    # find the .exe asset attached to the release
    download_url = None
    for asset in data.get("assets", []):
        if asset.get("name", "").lower().endswith(".exe"):
            download_url = asset.get("browser_download_url")
            break

    available = _version_tuple(latest_tag) > _version_tuple(current_version)
    return {
        "available": available and download_url is not None,
        "latest_version": latest_tag,
        "download_url": download_url,
        "notes": notes.strip(),
    }


def download_update(download_url: str, dest_path: str, progress_callback=None):
    """Streams the new .exe to dest_path. progress_callback(pct, downloaded, total)."""
    req = urllib.request.Request(download_url, headers={"User-Agent": "SPS-TDM-Updater"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            total = int(resp.headers.get("Content-Length", 0)) or None
            downloaded = 0
            chunk_size = 65536
            with open(dest_path, "wb") as out:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    out.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        pct = (downloaded / total * 100) if total else 0
                        progress_callback(pct, downloaded, total)
    except Exception as e:
        raise UpdateError(f"Download failed: {e}")
    return dest_path


def apply_update_and_restart(new_exe_path: str):
    """
    Replaces the currently-running .exe with new_exe_path and relaunches it.
    Only works when is_frozen() is True (i.e. actually running as the built
    .exe) — sys.executable then points at the real .exe file on disk.
    This function does not return: it exits the current process.
    """
    if not is_frozen():
        raise UpdateError("Self-update only works in the built .exe, not when running from source.")

    current_exe = sys.executable  # the real .exe path when frozen
    bat_path = os.path.join(os.path.dirname(new_exe_path), "_sps_updater.bat")

    bat_contents = f'''@echo off
setlocal
set OLD_EXE={current_exe}
set NEW_EXE={new_exe_path}

:wait_loop
del "%OLD_EXE%" >nul 2>&1
if exist "%OLD_EXE%" (
    timeout /t 1 /nobreak >nul
    goto wait_loop
)

move /y "%NEW_EXE%" "%OLD_EXE%" >nul
start "" "%OLD_EXE%"

del "%~f0" & exit
'''
    with open(bat_path, "w") as f:
        f.write(bat_contents)

    # DETACHED_PROCESS + CREATE_NEW_PROCESS_GROUP so the relauncher survives
    # after this exe closes.
    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )

    # Exit immediately so Windows releases the lock on the .exe file for
    # the relauncher's "del" step above.
    os._exit(0)

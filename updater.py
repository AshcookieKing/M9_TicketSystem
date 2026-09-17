"""Проверка GitHub Releases и установка новой сборки M9 Gate."""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import threading
import urllib.error
import urllib.request
from pathlib import Path

from app_meta import APP_VERSION, GITHUB_REPO, UPDATE_ASSET

TIMEOUT = 8
API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def _parse_version(text: str) -> tuple[int, ...]:
    digits = []
    for part in str(text or "").lstrip("vV").replace("-", ".").split("."):
        num = "".join(ch for ch in part if ch.isdigit())
        if num:
            digits.append(int(num))
    return tuple(digits or [0])


def is_newer(remote: str, local: str = APP_VERSION) -> bool:
    return _parse_version(remote) > _parse_version(local)


def install_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def fetch_latest_release():
    req = urllib.request.Request(
        API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"M9-Gate/{APP_VERSION}",
        },
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    tag = str(data.get("tag_name") or "").strip()
    assets = data.get("assets") or []
    asset = next((item for item in assets if item.get("name") == UPDATE_ASSET), None)
    if asset is None and assets:
        asset = next((item for item in assets if str(item.get("name") or "").endswith(".zip")), None)
    return {
        "tag": tag,
        "name": data.get("name") or tag,
        "url": (asset or {}).get("browser_download_url") or "",
        "size": (asset or {}).get("size") or 0,
        "notes": data.get("body") or "",
    }


def check_for_update():
    try:
        release = fetch_latest_release()
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError):
        return None
    if not release.get("tag") or not release.get("url"):
        return None
    if not is_newer(release["tag"], APP_VERSION):
        return None
    return release


def _download(url: str, dest: Path):
    req = urllib.request.Request(url, headers={"User-Agent": f"M9-Gate/{APP_VERSION}"})
    with urllib.request.urlopen(req, timeout=60) as resp, dest.open("wb") as out:
        shutil.copyfileobj(resp, out)


def apply_update(release: dict) -> Path:
    if not getattr(sys, "frozen", False):
        raise RuntimeError("Автообновление работает в установленной сборке (EXE).")
    url = release.get("url")
    if not url:
        raise RuntimeError("В релизе нет файла M9_Gate.zip")
    work = Path(tempfile.mkdtemp(prefix="m9gate_upd_"))
    archive = work / UPDATE_ASSET
    _download(url, archive)
    extract = work / "unpack"
    extract.mkdir()
    shutil.unpack_archive(str(archive), str(extract))
    payload = extract
    nested = extract / "M9_Gate"
    if nested.is_dir() and (nested / "M9_Gate.exe").exists():
        payload = nested
    elif not (extract / "M9_Gate.exe").exists():
        found = next(extract.rglob("M9_Gate.exe"), None)
        if found:
            payload = found.parent
        else:
            raise RuntimeError("В архиве нет M9_Gate.exe")
    app_dir = install_dir()
    bat = work / "install_update.bat"
    bat.write_text(
        "\n".join([
            "@echo off",
            "setlocal",
            "timeout /t 2 /nobreak >nul",
            f'xcopy /E /Y /Q "{payload}\\*" "{app_dir}\\" >nul',
            f'start "" "{app_dir / "M9_Gate.exe"}"',
            "endlocal",
        ]),
        encoding="utf-8",
    )
    os.startfile(str(bat))
    return bat


def start_background_check(on_found):
    def worker():
        release = check_for_update()
        if release:
            on_found(release)

    thread = threading.Thread(target=worker, daemon=True, name="m9-update")
    thread.start()
    return thread

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

from app_meta import APP_VERSION, GITHUB_REPO, UPDATE_ASSET, UPDATE_EXE

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


NOTICE_NAME = "update_notice.json"


def notice_path() -> Path:
    path = install_dir() / "data"
    path.mkdir(parents=True, exist_ok=True)
    return path / NOTICE_NAME


def consume_update_notice():
    path = notice_path()
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        data = {"to": APP_VERSION}
    try:
        path.unlink()
    except OSError:
        pass
    return data


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
    zip_asset = next((item for item in assets if item.get("name") == UPDATE_ASSET), None)
    exe_asset = next((item for item in assets if item.get("name") == UPDATE_EXE), None)
    if zip_asset is None:
        zip_asset = next((item for item in assets if str(item.get("name") or "").endswith(".zip")), None)
    return {
        "tag": tag,
        "name": data.get("name") or tag,
        "url": (zip_asset or {}).get("browser_download_url") or "",
        "exe_url": (exe_asset or {}).get("browser_download_url") or "",
        "size": (zip_asset or exe_asset or {}).get("size") or 0,
        "notes": data.get("body") or "",
    }


def check_for_update():
    try:
        release = fetch_latest_release()
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError):
        return None
    if not release.get("tag"):
        return None
    if not release.get("url") and not release.get("exe_url"):
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
    app_dir = install_dir()
    exe_path = Path(sys.executable).resolve()
    onedir = (app_dir / "_internal").exists()
    work = Path(tempfile.mkdtemp(prefix="m9gate_upd_"))
    notice = work / NOTICE_NAME
    notice.write_text(
        json.dumps(
            {
                "from": APP_VERSION,
                "to": release.get("tag") or "",
                "name": release.get("name") or "",
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    data_dir = app_dir / "data"
    if onedir:
        url = release.get("url")
        if not url:
            raise RuntimeError("В релизе нет файла M9_Gate.zip")
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
        install_lines = [
            f'xcopy /E /Y /Q "{payload}\\*" "{app_dir}\\" >nul',
            f'start "" "{app_dir / "M9_Gate.exe"}"',
        ]
    else:
        url = release.get("exe_url") or release.get("url")
        if not url:
            raise RuntimeError("В релизе нет файла M9_Gate.exe")
        new_exe = work / UPDATE_EXE
        if url.lower().endswith(".zip"):
            raise RuntimeError("Для одиночного EXE нужен файл M9_Gate.exe в релизе")
        _download(url, new_exe)
        install_lines = [
            f'copy /Y "{new_exe}" "{exe_path}" >nul',
            f'start "" "{exe_path}"',
        ]
    bat = work / "install_update.bat"
    bat.write_text(
        "\n".join([
            "@echo off",
            "setlocal",
            "timeout /t 2 /nobreak >nul",
            *install_lines[:-1],
            f'if not exist "{data_dir}" mkdir "{data_dir}"',
            f'copy /Y "{notice}" "{data_dir / NOTICE_NAME}" >nul',
            install_lines[-1],
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

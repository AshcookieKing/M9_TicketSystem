"""
Сборка M9 Gate.exe — система заявок ММТС-9
"""

from pathlib import Path
import os
import sys
import zipfile
import PyInstaller.__main__

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "Пример пропуска ТехРу (актуальный).xlsx"
ICON = ROOT / "assets" / "app.ico"


def _fix_build_tcl():
    prefix = Path(sys.base_prefix)
    tcl = prefix / "tcl" / "tcl8.6"
    tk = prefix / "tcl" / "tk8.6"
    if tcl.exists():
        os.environ["TCL_LIBRARY"] = str(tcl)
    if tk.exists():
        os.environ["TK_LIBRARY"] = str(tk)


def _common_args():
    return [
        str(ROOT / "pass_gui.py"),
        "--name=M9_Gate",
        "--windowed",
        "--noconsole",
        f"--icon={ICON}",
        f"--add-data={TEMPLATE};.",
        f"--add-data={ROOT / 'data' / 'stamps'};data/stamps",
        f"--add-data={ROOT / 'data' / 'catalog.json'};data",
        f"--add-data={ROOT / 'assets'};assets",
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=customtkinter",
        "--hidden-import=PIL",
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=reportlab",
        "--hidden-import=reportlab.pdfbase",
        "--hidden-import=reportlab.pdfbase.ttfonts",
        "--hidden-import=reportlab.graphics",
        "--hidden-import=line_core",
        "--hidden-import=updater",
        "--hidden-import=app_meta",
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
        "--exclude-module=pandas",
        "--exclude-module=scipy",
        "--exclude-module=pytest",
        "--collect-all=customtkinter",
        "--noconfirm",
    ]


def main():
    _fix_build_tcl()
    if not TEMPLATE.exists():
        raise SystemExit(f"Нет шаблона Excel: {TEMPLATE}")
    if not ICON.exists():
        raise SystemExit(f"Нет иконки: {ICON}")
    if not (ROOT / "data" / "catalog.json").exists():
        from pass_core import load_catalog
        load_catalog()

    common = _common_args()
    print("Сборка папки dist/M9_Gate ...")
    PyInstaller.__main__.run(common + ["--onedir", "--clean"])

    dist = ROOT / "dist" / "M9_Gate"
    zip_path = ROOT / "dist" / "M9_Gate.zip"
    if dist.exists():
        if zip_path.exists():
            zip_path.unlink()
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in dist.rglob("*"):
                if file.is_file():
                    zf.write(file, file.relative_to(dist))
        print(f"ZIP: {zip_path}")

    print("Сборка одиночного dist/M9_Gate.exe ...")
    PyInstaller.__main__.run(common + ["--onefile"])
    exe_path = ROOT / "dist" / "M9_Gate.exe"
    print(f"EXE: {exe_path}")
    print(f"Папка: {dist / 'M9_Gate.exe'}")


if __name__ == "__main__":
    main()

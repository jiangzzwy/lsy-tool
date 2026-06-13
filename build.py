"""Cross-platform build script for 案件线索移送函批量生成工具.

Usage:
    python build.py              # build for current platform
    python build.py --mac        # build macOS .app bundle
    python build.py --win        # build Windows .exe  (run on Windows)
    python build.py --clean      # clean build artifacts
"""

import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent


def clean():
    targets = ["build", "dist", "__pycache__"]
    for t in targets:
        p = BASE_DIR / t
        if p.exists():
            shutil.rmtree(p)
            print(f"  Removed: {p}")
    print("Clean done.")


def _add_data_dirs(cmd, sep):
    """Add --add-data entries for directories that exist on disk."""
    data_dirs = [
        (BASE_DIR / "demands", "demands"),
        (BASE_DIR / "web" / "templates", "web/templates"),
        (BASE_DIR / "web" / "static", "web/static"),
    ]
    for src, dest in data_dirs:
        if src.exists():
            cmd.extend(["--add-data", f"{src}{sep}{dest}"])
        else:
            print(f"  WARNING: Skipping missing data dir: {src}")


def _build_cmd():
    """Common PyInstaller args for the pywebview-based app."""
    sep = ";" if platform.system() == "Windows" else ":"
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "--noconfirm",
        "--windowed",
        "--name", "移送函批量生成工具",
        "--hidden-import", "flask",
        "--hidden-import", "webview",
        "--hidden-import", "openpyxl",
        "--hidden-import", "docx",
        "--hidden-import", "requests",
        "--hidden-import", "api_client",
        "--hidden-import", "config",
        "--hidden-import", "excel_parser",
        "--hidden-import", "word_generator",
        "--hidden-import", "ledger_generator",
        "--exclude-module", "tkinter",
        "--exclude-module", "dearpygui",
        "--exclude-module", "matplotlib",
        "--exclude-module", "numpy",
        "--exclude-module", "pandas",
        str(BASE_DIR / "web.py"),
    ]
    _add_data_dirs(cmd, sep)
    return cmd


def build_mac():
    print("Building macOS application...")
    subprocess.run(_build_cmd(), check=True)
    print(f"\nmacOS app built: {BASE_DIR / 'dist' / '移送函批量生成工具.app'}")


def build_win():
    if platform.system() != "Windows":
        print("WARNING: Windows .exe must be built on a Windows machine.")
        print("PyInstaller does not support cross-compilation.")
        print("Please run this script on Windows to generate the .exe\n")
        print("Quick steps on Windows:")
        print("  1. Copy the entire 'lsy' folder to Windows")
        print("  2. pip install -r requirements.txt")
        print("  3. python build.py --win")
        return
    print("Building Windows application...")
    subprocess.run(_build_cmd(), check=True)
    print(f"\nWindows exe built successfully in: {BASE_DIR / 'dist'}")


def main():
    parser = argparse.ArgumentParser(description="Build 案件线索移送函批量生成工具")
    parser.add_argument("--mac", action="store_true", help="Build macOS .app bundle")
    parser.add_argument("--win", action="store_true", help="Build Windows .exe")
    parser.add_argument("--clean", action="store_true", help="Clean build artifacts")
    args = parser.parse_args()

    if args.clean:
        clean()
        return

    if args.mac:
        build_mac()
    elif args.win:
        build_win()
    else:
        system = platform.system()
        if system == "Darwin":
            build_mac()
        elif system == "Windows":
            build_win()
        else:
            print(f"Platform: {system}, running generic build...")
            subprocess.run(_build_cmd(), check=True)


if __name__ == "__main__":
    main()

import shutil
import subprocess
import sys
import tomllib
from pathlib import Path

ROOT = Path(__file__).parent.resolve()

PYPROJECT = ROOT / "pyproject.toml"
VERSION_FILE = ROOT / "src" / "_version.py"

APP_DIR = ROOT / "application"
BUILD_DIR = APP_DIR / "build"

ICON = ROOT / "asset" / "icon.ico"


def get_version() -> str:
    with PYPROJECT.open("rb") as f:
        config = tomllib.load(f)

    return config["project"]["version"]


def generate_version_file(version: str) -> None:
    VERSION_FILE.write_text(
        f'__version__ = "{version}"\n',
        encoding="utf-8",
    )


def build() -> None:
    version = get_version()

    print(f"Building version {version}")
    print(f"Using icon: {ICON}")

    generate_version_file(version)

    if BUILD_DIR.exists():
        shutil.rmtree(BUILD_DIR)

    # Xóa exe cũ
    for exe in (
        APP_DIR / "gold-promo-tool.exe",
        APP_DIR / "gold-promo-updater.exe",
    ):
        if exe.exists():
            exe.unlink()

    common = [
        "--onefile",
        "--windowed",
        "--icon",
        str(ICON),
        "--workpath",
        str(APP_DIR / "build"),
        "--distpath",
        str(APP_DIR),
        "--specpath",
        str(APP_DIR),
        "--clean",
    ]

    for entrypoint, name in (
        ("src/desktop_app.py", "gold-promo-tool"),
        ("src/updater.py", "gold-promo-updater"),
    ):
        subprocess.run(
            [
                sys.executable,
                "-m",
                "PyInstaller",
                entrypoint,
                "--name",
                name,
                *common,
            ],
            cwd=ROOT,
            check=True,
        )


if __name__ == "__main__":
    build()
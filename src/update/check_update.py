from dataclasses import dataclass

import requests
from packaging.version import Version

GITHUB_OWNER = "khongtrii"
GITHUB_REPO = "gold-promo-tool"

LATEST_RELEASE_API = (
    f"https://api.github.com/repos/"
    f"{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)


@dataclass
class UpdateInfo:
    available: bool
    current_version: str
    latest_version: str
    download_url: str | None
    file_name: str | None


def check_update(current_version: str) -> UpdateInfo:
    """Check the latest release against the version supplied by desktop_app."""
    response = requests.get(
        LATEST_RELEASE_API,
        timeout=10,
    )

    response.raise_for_status()

    release = response.json()

    latest_version = release["tag_name"].lstrip("v")

    exe_asset = next(
        (
            asset
            for asset in release.get("assets", [])
            if asset["name"].lower().endswith(".exe")
            and "updater" not in asset["name"].lower()
        ),
        None,
    )

    if exe_asset is None:
        raise RuntimeError("Release không chứa file application .exe")

    available = (
        Version(latest_version)
        > Version(current_version)
    )

    return UpdateInfo(
        available=available,
        current_version=current_version,
        latest_version=latest_version,
        download_url=exe_asset["browser_download_url"],
        file_name=exe_asset["name"],
    )

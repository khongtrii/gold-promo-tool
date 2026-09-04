"""Resolve shared Gold Promo data files stored below the user's OneDrive."""

from __future__ import annotations

import os
from pathlib import Path


DATA_FILE_SYSTEM_FOLDERS = (
    Path("gold_promo") / "data_file_system",
    Path("Khong Minh Tri - gold_promo") / "data_file_system",
    Path("Khong Minh Tri's files - gold_promo") / "data_file_system",
)

GOLD_PROMO_SYSTEM_DIR = Path(
    r"\\al-dc01\MasterData\MasterDataMacro\goldPromoSystem"
)
DEFAULT_MASTER_DATA_PATH = GOLD_PROMO_SYSTEM_DIR / "master_data_metadata.xlsm"


def get_onedrive_root() -> Path | None:
    for variable in ("OneDriveRoot", "OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        value = os.environ.get(variable, "").strip()
        if value:
            return Path(value).expanduser()
    return None


def data_file_candidates(file_name: str) -> tuple[Path, ...]:
    root = get_onedrive_root()
    if root is None:
        return ()
    return tuple(root / folder / file_name for folder in DATA_FILE_SYSTEM_FOLDERS)


def find_existing_data_file(file_name: str) -> Path | None:
    return next((path for path in data_file_candidates(file_name) if path.is_file()), None)


def find_data_file_location(file_name: str) -> Path | None:
    """Find an existing file or a known existing data directory for its creation."""
    existing_file = find_existing_data_file(file_name)
    if existing_file is not None:
        return existing_file
    return next((path for path in data_file_candidates(file_name) if path.parent.is_dir()), None)


def default_master_data_path() -> str:
    return str(DEFAULT_MASTER_DATA_PATH)

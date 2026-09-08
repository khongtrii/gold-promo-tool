"""Shared JSON state for Site Group exclusion and edit coordination."""

from __future__ import annotations

import json
import os
import tempfile
import time
from copy import deepcopy
from pathlib import Path
from typing import Callable

from src.data_file_paths import GOLD_PROMO_SYSTEM_DIR

DEFAULT_EXCLUDED_SITEGROUPS = (
    "8000", "8200", "8201", "8202", "8203", "8300", "8710", "8711", "8712",
    "8210", "8220", "8230", "8310", "8320", "8330",
    "1100", "1200", "1300", "2100", "2200", "2300",
    "99990", "99991", "99992", "99993", "99994",
    "99995", "99996", "99997", "99998", "99999",
)
DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES = (
    "02043862",
    "02043863",
    "02047463",
    "02047464",
)
DEFAULT_SITEGROUP_STATE = {
    "EXC_SITE_GROUP": list(DEFAULT_EXCLUDED_SITEGROUPS),
    "EXCEPTION_DISCOUNT_GC": list(DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES),
    "ACTIVE": "no",
}


def get_sitegroup_state_path(catalogue: str) -> Path | None:
    catalogue = str(catalogue).strip()
    if not catalogue:
        return None
    return GOLD_PROMO_SYSTEM_DIR / f"{catalogue}_sitegroup_state.json"


def _required_state_path(path: Path | None) -> Path:
    resolved = Path(path) if path is not None else None
    if resolved is None:
        raise FileNotFoundError(
            "Cannot determine the catalogue-specific Site Group state file."
        )
    return resolved


def _normalize_state(state) -> tuple[dict, bool]:
    normalized = dict(state) if isinstance(state, dict) else {}
    changed = not isinstance(state, dict)

    raw_codes = normalized.get("EXC_SITE_GROUP", [])
    if not isinstance(raw_codes, list):
        raw_codes = []
        changed = True
    codes = []
    seen = set()
    # Built-in exceptions are always present while retaining user-added codes.
    for value in (*DEFAULT_EXCLUDED_SITEGROUPS, *raw_codes):
        code = str(value).strip()
        if code and code not in seen:
            codes.append(code)
            seen.add(code)
    if normalized.get("EXC_SITE_GROUP") != codes:
        changed = True
    normalized["EXC_SITE_GROUP"] = codes

    raw_discount_codes = normalized.get(
        "EXCEPTION_DISCOUNT_GC",
        list(DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES),
    )
    if not isinstance(raw_discount_codes, list):
        raw_discount_codes = list(DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES)
        changed = True
    discount_codes = []
    seen_discount_codes = set()
    for value in raw_discount_codes:
        code = str(value).strip()
        if code and code not in seen_discount_codes:
            discount_codes.append(code)
            seen_discount_codes.add(code)
    if normalized.get("EXCEPTION_DISCOUNT_GC") != discount_codes:
        changed = True
    normalized["EXCEPTION_DISCOUNT_GC"] = discount_codes

    if normalized.get("ACTIVE") not in {"yes", "no"}:
        normalized["ACTIVE"] = "no"
        changed = True
    return normalized, changed


def _atomic_write(state: dict, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".tmp",
            prefix=f".{path.name}.",
            dir=path.parent,
            delete=False,
        ) as temporary_file:
            json.dump(state, temporary_file, ensure_ascii=False, indent=4)
            temporary_file.write("\n")
            temporary_file.flush()
            os.fsync(temporary_file.fileno())
            temporary_path = Path(temporary_file.name)
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


class _StateLock:
    def __init__(self, lock_path: Path, timeout: float = 2.0) -> None:
        self.lock_path = lock_path
        self.timeout = timeout
        self.file_descriptor: int | None = None

    def __enter__(self):
        deadline = time.monotonic() + self.timeout
        while True:
            try:
                self.file_descriptor = os.open(
                    self.lock_path,
                    os.O_CREAT | os.O_EXCL | os.O_WRONLY,
                )
                os.write(self.file_descriptor, f"{os.getpid()}\n".encode("ascii"))
                return self
            except FileExistsError:
                try:
                    if time.time() - self.lock_path.stat().st_mtime > 30:
                        self.lock_path.unlink()
                        continue
                except FileNotFoundError:
                    continue
                if time.monotonic() >= deadline:
                    raise TimeoutError("Site Group state is busy. Please try again.")
                time.sleep(0.05)

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        if self.file_descriptor is not None:
            os.close(self.file_descriptor)
        try:
            self.lock_path.unlink()
        except FileNotFoundError:
            pass


def _lock_path_for(path: Path) -> Path:
    return path.with_suffix(path.suffix + ".lock")


def _read_state(path: Path) -> tuple[dict, bool]:
    if not path.exists():
        return deepcopy(DEFAULT_SITEGROUP_STATE), True
    try:
        with path.open("r", encoding="utf-8") as state_file:
            return _normalize_state(json.load(state_file))
    except json.JSONDecodeError as error:
        raise ValueError(f"Invalid Site Group state JSON: {path}") from error


def load_sitegroup_state(path: Path | None = None) -> dict:
    """Read state and initialize missing file/fields without dropping unknown fields."""
    path = _required_state_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    state, changed = _read_state(path)
    if changed:
        with _StateLock(_lock_path_for(path)):
            state, changed = _read_state(path)
            if changed:
                _atomic_write(state, path)
    return state


def save_sitegroup_state(state: dict, path: Path | None = None) -> dict:
    """Merge supplied fields into the latest state and save atomically."""
    path = _required_state_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with _StateLock(_lock_path_for(path)):
        latest, _ = _read_state(path)
        latest.update(state)
        latest, _ = _normalize_state(latest)
        _atomic_write(latest, path)
        return latest


def _update_state(update: Callable[[dict], None], path: Path | None = None) -> dict:
    path = _required_state_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with _StateLock(_lock_path_for(path)):
        state, _ = _read_state(path)
        update(state)
        state, _ = _normalize_state(state)
        _atomic_write(state, path)
        return state


def get_excluded_sitegroups(path: Path | None = None) -> list[str]:
    return list(load_sitegroup_state(path)["EXC_SITE_GROUP"])


def add_excluded_sitegroup(code: str, path: Path | None = None) -> list[str]:
    code = str(code).strip()
    if not code:
        return get_excluded_sitegroups(path)

    state = _update_state(
        lambda current: current["EXC_SITE_GROUP"].append(code)
        if code not in current["EXC_SITE_GROUP"]
        else None,
        path,
    )
    return list(state["EXC_SITE_GROUP"])


def remove_excluded_sitegroup(code: str, path: Path | None = None) -> list[str]:
    code = str(code).strip()

    def remove(current: dict) -> None:
        current["EXC_SITE_GROUP"] = [value for value in current["EXC_SITE_GROUP"] if value != code]

    return list(_update_state(remove, path)["EXC_SITE_GROUP"])


def get_exception_discount_gold_codes(path: Path | None = None) -> list[str]:
    return list(load_sitegroup_state(path)["EXCEPTION_DISCOUNT_GC"])


def add_exception_discount_gold_code(code: str, path: Path | None = None) -> list[str]:
    code = str(code).strip()
    if not code:
        return get_exception_discount_gold_codes(path)

    state = _update_state(
        lambda current: current["EXCEPTION_DISCOUNT_GC"].append(code)
        if code not in current["EXCEPTION_DISCOUNT_GC"]
        else None,
        path,
    )
    return list(state["EXCEPTION_DISCOUNT_GC"])


def remove_exception_discount_gold_code(code: str, path: Path | None = None) -> list[str]:
    code = str(code).strip()

    def remove(current: dict) -> None:
        current["EXCEPTION_DISCOUNT_GC"] = [
            value for value in current["EXCEPTION_DISCOUNT_GC"] if value != code
        ]

    return list(_update_state(remove, path)["EXCEPTION_DISCOUNT_GC"])


def get_active_status(path: Path | None = None) -> str:
    return load_sitegroup_state(path)["ACTIVE"]


def set_active_status(status: str, path: Path | None = None) -> None:
    if status not in {"yes", "no"}:
        raise ValueError('ACTIVE must be "yes" or "no".')
    _update_state(lambda state: state.__setitem__("ACTIVE", status), path)


def try_acquire_active_status(path: Path | None = None) -> bool:
    """Atomically change ACTIVE from no to yes and report whether it succeeded."""
    acquired = False

    def acquire(state: dict) -> None:
        nonlocal acquired
        if state["ACTIVE"] == "no":
            state["ACTIVE"] = "yes"
            acquired = True

    _update_state(acquire, path)
    return acquired

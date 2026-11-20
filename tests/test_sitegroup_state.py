import json
import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from src.sitegroup_state import (
    DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES,
    _atomic_write,
    _normalize_state,
    add_excluded_sitegroups,
    get_sitegroup_state_path,
    save_sitegroup_state,
)


class SiteGroupStateTest(unittest.TestCase):
    def test_state_file_is_next_to_selected_metadata(self):
        metadata = Path("D:/custom/metadata/master_data_metadata.xlsm")

        path = get_sitegroup_state_path("C01", metadata)

        self.assertEqual(path, metadata.parent / "C01_sitegroup_state.json")

    def test_initializes_default_discount_exceptions(self):
        state, changed = _normalize_state({})

        self.assertTrue(changed)
        self.assertEqual(
            state["EXCEPTION_DISCOUNT_GC"],
            list(DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES),
        )

    def test_removed_default_is_not_added_back(self):
        removed = DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES[0]
        saved_codes = list(DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES[1:])

        state, _ = _normalize_state({"EXCEPTION_DISCOUNT_GC": saved_codes})

        self.assertNotIn(removed, state["EXCEPTION_DISCOUNT_GC"])

    def test_bulk_add_writes_once_and_preserves_state(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            save_sitegroup_state(
                {"EXC_SITE_GROUP": ["60001"], "CUSTOM": "preserved"},
                path,
            )

            with patch(
                "src.sitegroup_state._atomic_write", wraps=_atomic_write
            ) as atomic_write:
                codes = add_excluded_sitegroups(
                    ["60001", "60002", " 60003 ", "60002", ""],
                    path,
                )

            self.assertEqual(atomic_write.call_count, 1)
            self.assertEqual(codes.count("60001"), 1)
            self.assertEqual(codes.count("60002"), 1)
            self.assertEqual(codes.count("60003"), 1)
            state = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(state["CUSTOM"], "preserved")

    def test_atomic_write_retries_transient_replace_errors(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            real_replace = os.replace
            attempts = 0

            def transient_then_success(source, destination):
                nonlocal attempts
                attempts += 1
                if attempts < 3:
                    raise PermissionError(13, "temporarily locked")
                real_replace(source, destination)

            with patch(
                "src.sitegroup_state.os.replace",
                side_effect=transient_then_success,
            ), patch("src.sitegroup_state.time.sleep") as sleep:
                _atomic_write({"VALUE": "saved"}, path)

            self.assertEqual(attempts, 3)
            self.assertEqual(sleep.call_count, 2)
            self.assertEqual(
                json.loads(path.read_text(encoding="utf-8")),
                {"VALUE": "saved"},
            )

    def test_atomic_write_exhausts_retry_and_cleans_temporary_file(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            path.write_text('{"VALUE": "original"}\n', encoding="utf-8")
            with patch(
                "src.sitegroup_state.os.replace",
                side_effect=PermissionError(13, "locked"),
            ), patch(
                "src.sitegroup_state.time.monotonic", side_effect=[0.0, 6.0]
            ):
                with self.assertRaises(PermissionError):
                    _atomic_write({"VALUE": "not saved"}, path)

            self.assertEqual(
                json.loads(path.read_text(encoding="utf-8")),
                {"VALUE": "original"},
            )
            self.assertEqual(list(Path(directory).glob("*.tmp")), [])

    def test_atomic_write_does_not_retry_non_transient_error(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            with patch(
                "src.sitegroup_state.os.replace",
                side_effect=OSError(22, "invalid argument"),
            ) as replace, patch("src.sitegroup_state.time.sleep") as sleep:
                with self.assertRaises(OSError):
                    _atomic_write({"VALUE": "not saved"}, path)

            replace.assert_called_once()
            sleep.assert_not_called()


if __name__ == "__main__":
    unittest.main()

import unittest

from src.sitegroup_state import (
    DEFAULT_EXCEPTION_DISCOUNT_GOLD_CODES,
    _normalize_state,
)


class SiteGroupStateTest(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()

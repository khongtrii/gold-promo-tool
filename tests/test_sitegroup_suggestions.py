import unittest
from unittest.mock import patch

import pandas as pd

from src.service.template_service import Template_ETL


class SiteGroupSuggestionTest(unittest.TestCase):
    def make_etl(self, rows):
        etl = object.__new__(Template_ETL)
        etl.src = pd.DataFrame(rows)
        etl.master_sitegroup_codes = {"12345"}
        etl.activated_sitegroup_codes = {"54321"}
        etl.non_suggested_sitegroup_codes = {"99999"}
        etl.sitegroup_members = {"12345": ("101", "102")}
        return etl

    def test_complete_so_requires_every_row(self):
        etl = self.make_etl([{"SO": "C01D01-01"}, {"SO": "  "}])
        self.assertFalse(etl.has_complete_so())

        etl.src.loc[1, "SO"] = "C01D01-02"
        self.assertTrue(etl.has_complete_so())

    @patch.object(Template_ETL, "_generate_sitegroup_code", side_effect=["60000", "60001"])
    def test_every_mismatch_gets_a_distinct_new_code(self, generate):
        etl = self.make_etl([
            {
                "SITE GROUP": "",
                "GOLD PROMO NETWORK EXPANDED": "101;103",
                "GOLD PROMO NETWORK": "8001",
                "STRUCTURE": "1",
            },
            {
                "SITE GROUP": "",
                "GOLD PROMO NETWORK EXPANDED": "104;105",
                "GOLD PROMO NETWORK": "8002",
                "STRUCTURE": "1",
            },
        ])

        suggestions = etl.get_sitegroup_suggestions()

        self.assertEqual([item["suggested_code"] for item in suggestions], ["60000", "60001"])
        self.assertTrue(all(item["original_suggested_code"] == "" for item in suggestions))
        unavailable_second_call = generate.call_args_list[1].args[0]
        self.assertIn("60000", unavailable_second_call)
        self.assertIn("99999", unavailable_second_call)

    def test_exact_match_is_not_reviewed(self):
        etl = self.make_etl([
            {
                "SITE GROUP": "12345",
                "GOLD PROMO NETWORK EXPANDED": "101;102",
                "GOLD PROMO NETWORK": "8001",
                "STRUCTURE": "1",
            }
        ])

        self.assertEqual(etl.get_sitegroup_suggestions(), [])


if __name__ == "__main__":
    unittest.main()

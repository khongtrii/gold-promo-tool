import unittest
from unittest.mock import patch

import pandas as pd

from src.service.template_service import Template_ETL


class DiscountNetworkTest(unittest.TestCase):
    def setUp(self):
        self.etl = Template_ETL([])
        network = pd.DataFrame({
            "SITE": ["101", "102", "103", "104"],
            "NATIONAL_SITE": ["8200"] * 4,
            "GROUP_SITE": ["8210"] * 4,
            "REGION_SITE": ["8220"] * 4,
            "ACTIVE": ["1", "0", "1", "0"],
            "DISCOUNT": ["0", "1", "1", "0"],
        })
        with patch("src.service.template_service.pd.read_excel", return_value=network):
            self.etl._load_network()

    def test_discount_filter_is_independent_of_active_for_all_groups(self):
        for group in ("8200", "8210", "8220"):
            self.assertEqual(self.etl.dict_network["network"][group], "101;103")
            self.assertEqual(self.etl.dict_network["DISCOUNT_NETWORK"][group], "102;103")

    def test_discount_expansion_uses_same_expression_rules(self):
        self.assertEqual(self.etl._extract_discount_network("8210(-103);105"), "102;105")
        self.assertEqual(self.etl._extract_discount_network("8210;8220"), "102;103")
        self.assertEqual(self.etl._extract_discount_network("8210-8220"), "")
        self.assertEqual(self.etl._extract_discount_network(""), "")

    def test_validation_populates_both_expansions_from_purchase_network(self):
        data = pd.DataFrame({
            "FILE NAME": ["source.xlsx"],
            "GOLD CODE": ["GC1"], "LV": ["1"], "LU": ["1"],
            "PURCHASE NETWORK": ["8210"], "GOLD PROMO NETWORK": ["8210"],
        })
        result = self.etl._check_network(data)
        self.assertEqual(result.at[0, "PURCHASE NETWORK EXPANDED"], "101;103")
        self.assertEqual(result.at[0, "DISCOUNT_NETWORK_EXPANDED"], "102;103")


if __name__ == "__main__":
    unittest.main()

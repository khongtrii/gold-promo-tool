import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from src.service.template_service import Template_ETL
from src.service.template_mapping import Discount
from src.constant.required import required_wh_discount


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

    def test_stage1_only_accepts_8200_sites_for_purchase_and_promo(self):
        self.etl.dict_network["network"]["8710"] = "201"
        self.etl.dict_network["store_minigo"] = ["201"]
        self.etl.dict_network["store"].append("201")
        data = pd.DataFrame({
            "FILE NAME": ["source.xlsx"] * 3,
            "GOLD CODE": ["GC1", "GC2", "GC3"],
            "LV": ["1"] * 3, "LU": ["1"] * 3,
            "PURCHASE NETWORK": ["8200", "8710", "8200"],
            "GOLD PROMO NETWORK": ["8200", "8200", "8710"],
        })
        result = self.etl._check_network(data)
        self.assertEqual(result.at[0, "NOTE ERR FROM MASTER DATA"], "")
        for index in (1, 2):
            self.assertIn("201 không thuộc network 8200", result.at[index, "NOTE ERR FROM MASTER DATA"])

    def test_stage3_requires_wh_sites_from_discount_network(self):
        network = pd.DataFrame({
            "SITE": ["801", "802", "101"],
            "NATIONAL_SITE": ["8300", "8300", "8200"],
            "GROUP_SITE": ["8310", "8310", "8210"],
            "REGION_SITE": ["8320", "8320", "8220"],
            "ACTIVE": ["0", "1", "1"],
            "DISCOUNT": ["1", "0", "1"],
        })
        expressions = ["8300", "801", "802", "8300+8200", "999", "8300-801"]
        data = pd.DataFrame({column: ["1"] * len(expressions) for column in required_wh_discount})
        data["PURCHASE NETWORK"] = expressions
        data["PP START YEAR"] = "2099"
        data["PP END YEAR"] = "2099"
        data["PP END DAY"] = "2"
        etl = Template_ETL([Path("source.xlsx")])
        etl.dept = {"source.xlsx": "110"}
        with patch.object(etl, "_load_source_metadata"), patch(
            "src.service.template_service.pd.read_excel", side_effect=[network, data]
        ):
            etl._load_network()._load_src_wh_discount()
        notes = etl.src["NOTE ERR FROM MASTER DATA"].fillna("").tolist()
        self.assertEqual(notes[:2], ["", ""])
        for index, site in ((2, "802"), (3, "101"), (4, "999")):
            self.assertIn(f"Site {site} không thuộc WH (8300) của DISCOUNT_NETWORK", notes[index])
        self.assertIn("bị rỗng", notes[5])
        self.assertEqual(etl.src.at[0, "DISCOUNT_NETWORK_EXPANDED"], "801")

        # Missing discount WH configuration must not fall back to ACTIVE WH.
        etl.dict_network["DISCOUNT_NETWORK"].pop("8300")
        with patch.object(etl, "_load_source_metadata"), patch(
            "src.service.template_service.pd.read_excel", return_value=data
        ):
            etl._load_src_wh_discount()
        self.assertIn("không thuộc WH", etl.src.at[1, "NOTE ERR FROM MASTER DATA"])

    def _load_wh_networks(self, expressions, wh_sites="801;802;809"):
        data = pd.DataFrame({column: ["1"] * len(expressions) for column in required_wh_discount})
        data["PURCHASE NETWORK"] = expressions
        data["GOLD CODE"] = [str(index + 1) for index in range(len(expressions))]
        data["PP START YEAR"] = "2099"
        data["PP END YEAR"] = "2099"
        data["PP END DAY"] = "2"
        etl = Template_ETL([Path("source.xlsx")])
        etl.dept = {"source.xlsx": "110"}
        etl.dict_network = {"DISCOUNT_NETWORK": {"8300": wh_sites}, "wh8": []}
        with patch.object(etl, "_load_source_metadata"), patch(
            "src.service.template_service.pd.read_excel", return_value=data
        ):
            etl._load_src_wh_discount()
        return etl

    def test_stage3_adds_839_after_expansion_and_passes_it_to_ag(self):
        expressions = ["802", "809", "802;809", "801", "8300", "8300(-839)"]
        etl = self._load_wh_networks(expressions)
        expected = ["802;839", "809;839", "802;809;839", "801",
                    "801;802;809;839", "801;802;809;839"]
        self.assertEqual(etl.src["DISCOUNT_NETWORK_EXPANDED"].tolist(), expected)
        self.assertEqual(etl.src["NOTE ERR FROM MASTER DATA"].tolist(), [""] * len(expressions))
        for column in ("PURCHASE NETWORK", "PURCHASE NETWORK EXPANDED"):
            self.assertEqual(etl.src[column].tolist(), expressions)
        discount = Discount(etl, ag_type="WH")._create_ag_raw()._create_ag()
        for frame in (discount.template_ag_raw, discount.template_ag):
            for index, sites in enumerate(expected):
                self.assertEqual(
                    sorted(frame.loc[frame["GOLD CODE"].eq(str(index + 1)), "SITE"].tolist()),
                    sites.split(";"),
                )

    def test_stage3_deduplicates_839_and_adds_it_back_after_exclusion(self):
        etl = self._load_wh_networks(
            ["802;839", "8300", "8300(-839)", "839"], wh_sites="802;809;839"
        )
        self.assertEqual(etl.src["DISCOUNT_NETWORK_EXPANDED"].tolist(),
                         ["802;839", "802;809;839", "802;809;839", "839"])
        self.assertEqual(etl.src["NOTE ERR FROM MASTER DATA"].tolist(), [""] * 4)

    def test_stage3_still_validates_original_sites_before_adding_839(self):
        etl = self._load_wh_networks(["839", "802;839", "802;999", "8300-8300", "809"],
                                     wh_sites="801;802")
        notes = etl.src["NOTE ERR FROM MASTER DATA"].tolist()
        for index, site in ((0, "839"), (1, "839"), (2, "999"), (4, "809")):
            self.assertIn(f"Site {site} không thuộc WH", notes[index])
        self.assertIn("bị rỗng", notes[3])
        etl = self._load_wh_networks(["8300"], wh_sites="802")
        etl.dict_network["DISCOUNT_NETWORK"]["8310"] = "802;839"
        with patch.object(etl, "_load_source_metadata"), patch(
            "src.service.template_service.pd.read_excel",
            return_value=etl.src.assign(**{"PURCHASE NETWORK": "8310"}),
        ):
            etl._load_src_wh_discount()
        self.assertIn("Site 839 không thuộc WH", etl.src.at[0, "NOTE ERR FROM MASTER DATA"])

    def test_stage1_network_check_does_not_add_839(self):
        data = pd.DataFrame({
            "FILE NAME": ["source.xlsx"], "GOLD CODE": ["GC1"], "LV": ["1"], "LU": ["1"],
            "PURCHASE NETWORK": ["802;809"], "GOLD PROMO NETWORK": ["802;809"],
        })
        result = self.etl._check_network(data)
        self.assertEqual(result.at[0, "DISCOUNT_NETWORK_EXPANDED"], "802;809")

    def test_stage3_rejects_plain_number_plus_number_discount(self):
        values = ["10+5", "10T+5TH", "10", "10%"]
        data = pd.DataFrame({column: ["1"] * len(values) for column in required_wh_discount})
        data["DISCOUNT (% OR VALUE)"] = values
        data["PURCHASE NETWORK"] = "101"
        data["PP START YEAR"] = "2099"
        data["PP END YEAR"] = "2099"
        data["PP END DAY"] = "2"
        etl = Template_ETL([Path("source.xlsx")])
        etl.dept = {"source.xlsx": "110"}
        with patch.object(etl, "_load_source_metadata"), patch(
            "src.service.template_service.pd.read_excel", return_value=data
        ):
            etl.dict_network = {"DISCOUNT_NETWORK": {"8300": "101"}}
            etl._load_src_wh_discount()
        notes = etl.src["NOTE ERR FROM MASTER DATA"].fillna("").tolist()
        self.assertIn("số+số", notes[0])
        self.assertEqual(notes[1:], ["", "", ""])


if __name__ == "__main__":
    unittest.main()

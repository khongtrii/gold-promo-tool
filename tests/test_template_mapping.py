import unittest
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd

from src.constant.template import (
    column_dc,
    column_po_commitment,
    column_promotion_plan,
    column_purchase,
)
from src.service.template_mapping import AttributeMapMixin, Discount, SalePrice, Template_Mapping
from src.service.template_service import Template_ETL


class SalePriceAttributeTest(unittest.TestCase):
    def test_note_count_is_added_after_attribute_columns_are_selected(self):
        etl = SimpleNamespace(
            src_listoff=None,
            src_attr=pd.DataFrame({
                "GOLD CODE": ["GC1", "GC1"],
                "SV": ["1", "1"],
                "CLASS": ["C", "C"],
                "Alphanum": ["A", "B"],
                "START DATE": ["01/01/2026"] * 2,
                "END DATE": ["31/12/2026"] * 2,
            }),
            cata="C01",
        )
        result = SalePrice(etl)._create_attr().template_attr
        self.assertEqual(result["NOTE COUNT"].tolist(), [2, 2])

    def test_bf_catalogue_code_has_no_d_suffix(self):
        etl = SimpleNamespace(
            src_listoff=None,
            src_attr=pd.DataFrame({
                "GOLD CODE": ["02094686"],
                "SV": ["1"],
                "CLASS": ["CATAP"],
                "Alphanum": ["OTHER.P.3.8200"],
                "START DATE": ["08/10/2026"],
                "END DATE": ["21/10/2026"],
            }),
            cata="BF621",
        )
        result = SalePrice(etl)._create_attr().template_attr
        self.assertEqual(result["CODE"].tolist(), ["BF621"])


class AttributeMediumMappingTest(unittest.TestCase):
    def test_maps_all_supported_medium_labels_to_numeric_codes(self):
        mapper = AttributeMapMixin()
        expected = {
            "FRONT PAGE": 4,
            "BACK PAGE": 4,
            "UNBEAT": 4,
            "HERO": 4,
            "STAR": 5,
            "MODEL": 6,
            "CATA": 7,
            "COMPLE": 8,
        }

        for label, code in expected.items():
            with self.subTest(label=label):
                self.assertEqual(mapper.attribute_map(label), code)

    def test_delete_rows_are_split_from_add_attribute_template(self):
        etl = SimpleNamespace(
            src=pd.DataFrame(
                {
                    "SO": ["SO1", "SO2", "SO2", "SO3"],
                    "GOLD CODE": ["GC1", "GC2", "GC2", "GC3"],
                    "LV": ["1", "2", "2", "3"],
                    "LU": ["1", "1", "1", "1"],
                    "ATTRIBUTE MARKETING": ["Hero", "please DELETE", "please delete", "delete hero"],
                    "FREE PRODUCT": ["", "", "", ""],
                }
            ),
            dict_network={},
            plan=pd.DataFrame(),
            cata="C01",
            cata_description="",
            cata_period="",
        )

        mapping = Template_Mapping(etl)._create_add_attribute_marketing()

        self.assertEqual(mapping.gold_code_delete.to_dict("records"), [
            {"GOLD CODE": "GC2", "LV": "2", "SO": "SO2"},
            {"GOLD CODE": "GC3", "LV": "3", "SO": "SO3"},
        ])
        self.assertEqual(
            mapping.template_add_attribute_marketing["GOLD CODE"].tolist(),
            ["GC1"],
        )


class PromotionPlanMappingTest(unittest.TestCase):
    def test_description_contains_only_catalogue_description(self):
        plan_dates = {
            "CATALOGUE START DATE": pd.Timestamp("2026-10-02"),
            "CATALOGUE END DATE": pd.Timestamp("2026-10-10"),
            "GLOBAL PERIOD START": pd.Timestamp("2026-10-01"),
            "GLOBAL PERIOD END": pd.Timestamp("2026-10-10"),
            "SHOP ACTIVATION": pd.Timestamp("2026-09-09"),
            "COMMITMENT DEADLINE": pd.Timestamp("2026-09-15"),
            "COMMITMENT CLOSING": pd.Timestamp("2026-09-15"),
            "ORDER WAREHOUSE START": pd.Timestamp("2026-09-18"),
            "ORDER WAREHOUSE END": pd.Timestamp("2026-10-10"),
        }
        etl = SimpleNamespace(
            src=pd.DataFrame({"SO": ["SO1"], "SITE GROUP": ["SG1"]}),
            dict_network={},
            plan=pd.DataFrame([plan_dates]),
            cata="C01",
            cata_description="Catalogue description",
            cata_period="Period",
        )

        mapping = Template_Mapping(etl)._create_promotion_plan()

        self.assertEqual(
            mapping.template_promotion_plan[column_promotion_plan[1]].iat[0],
            "Catalogue description",
        )


class SOCalendarMappingTest(unittest.TestCase):
    def test_pct_weight_drops_percent_sign(self):
        plan = pd.DataFrame(
            [
                {
                    "DIRECT | ORDER DATE 1": pd.Timestamp("2026-10-01"),
                    "DIRECT | ORDER DATE 2": pd.Timestamp("2026-10-02"),
                    "DIRECT | ORDER DATE 3": pd.Timestamp("2026-10-03"),
                    "DELIVERY DATE 1": pd.Timestamp("2026-10-04"),
                    "DELIVERY DATE 2": pd.Timestamp("2026-10-05"),
                    "DELIVERY DATE 3": pd.Timestamp("2026-10-06"),
                }
            ]
        )
        etl = SimpleNamespace(
            src=pd.DataFrame(
                {
                    "SO": ["SO1"],
                    "PURCHASE NETWORK EXPANDED": ["1"],
                    "SUPPLIER CODE": ["SUP"],
                    "COMMERCIAL CONTRACT": ["CONT"],
                    "GOLD CODE": ["02043862"],
                    "LV": ["1"],
                    "LU": ["1"],
                    "% DELIVERY 1": ["25%"],
                    "% DELIVERY 2": ["35%"],
                    "% DELIVERY 3": ["40%"],
                    "DELIVERY TYPE": ["DIRECT"],
                }
            ),
            dict_network={"wh8": []},
            plan=plan,
            cata="C01",
            cata_description="Catalogue description",
            cata_period="Period",
        )

        mapping = Template_Mapping(etl)._create_so_calendar()

        self.assertEqual(
            mapping.template_so_calendar["PCT WEIGHT"].tolist(),
            ["25", "35", "40"],
        )


class POCommitmentMappingTest(unittest.TestCase):
    def test_creates_supplier_free_template_and_supplier_report(self):
        etl = SimpleNamespace(
            src=pd.DataFrame(
                {
                    "SO": ["SO1"],
                    "GOLD CODE": ["02043862"],
                    "LV": ["1"],
                    "LU": ["1"],
                    "PURCHASE NETWORK EXPANDED": ["S1;S2"],
                    "SUPPLIER CODE": ["SUP"],
                    "S1": [12],
                    "S2": [24],
                }
            ),
            dict_network={"store": ["S1", "S2"]},
            plan=pd.DataFrame(),
            cata="C01",
            cata_description="Catalogue description",
            cata_period="Period",
        )

        mapping = Template_Mapping(etl)._create_po_commitment()

        self.assertEqual(
            mapping.template_po_commitment.columns.tolist(),
            ["NO", *column_po_commitment],
        )
        self.assertNotIn("SUPPLIER", mapping.template_po_commitment.columns)
        self.assertEqual(
            mapping.template_po_commitment_report.columns.tolist(),
            ["NO", *column_po_commitment, "SUPPLIER"],
        )
        self.assertEqual(mapping.template_po_commitment["QUANTITY"].tolist(), [12, 24])
        self.assertEqual(mapping.template_po_commitment_report["SUPPLIER"].tolist(), ["SUP", "SUP"])


class PurchaseMappingTest(unittest.TestCase):
    def test_always_adds_minigo_and_warehouse_sites(self):
        etl = SimpleNamespace(
            src=pd.DataFrame(
                {
                    "GOLD CODE": ["02043862"],
                    "LV": ["1"],
                    "NORMAL PURCHASE PRICE": [45100],
                    "PURCHASE NETWORK EXPANDED": ["P1;P2"],
                    "DISCOUNT_NETWORK_EXPANDED": ["S1;S2"],
                    "PP START DATE": [pd.Timestamp("2026-10-01")],
                    "PP END DATE": [pd.Timestamp("2026-10-10")],
                    "COMMERCIAL CONTRACT": ["CONT"],
                    "PURCHASE VAT": ["10%"],
                    "SUPPLIER CODE": ["SUP"],
                }
            ),
            dict_network={
                "DISCOUNT_NETWORK": {"8710": "M1;M2", "8300": "W1;W2"},
                "store_minigo": ["M1", "M2"],
                "wh": ["W1", "W2"],
                "wh8": [],
            },
            plan=pd.DataFrame(),
            cata="C01",
            cata_description="Catalogue description",
            cata_period="Period",
        )

        mapping = Template_Mapping(etl)._create_purchase()

        self.assertEqual(
            mapping.template_purchase["SITE"].tolist(),
            ["S1", "S2", "M1", "M2", "W1", "W2"],
        )

    def test_formats_whole_prices_and_removes_float_noise(self):
        etl = SimpleNamespace(
            src=pd.DataFrame(
                {
                    "GOLD CODE": ["02043862", "02043863", "02043864"],
                    "LV": ["1", "1", "1"],
                    "NORMAL PURCHASE PRICE": [45100.0, 45100.0000000001, 45100.25],
                    "PURCHASE NETWORK EXPANDED": ["S1", "S1", "S1"],
                    "DISCOUNT_NETWORK_EXPANDED": ["S1", "S1", "S1"],
                    "PP START DATE": [pd.Timestamp("2026-10-01")] * 3,
                    "PP END DATE": [pd.Timestamp("2026-10-10")] * 3,
                    "COMMERCIAL CONTRACT": ["CONT"] * 3,
                    "PURCHASE VAT": ["10%"] * 3,
                    "SUPPLIER CODE": ["SUP"] * 3,
                }
            ),
            dict_network={"wh8": []},
            plan=pd.DataFrame(),
            cata="C01",
            cata_description="Catalogue description",
            cata_period="Period",
        )

        mapping = Template_Mapping(etl)._create_purchase()

        self.assertEqual(
            mapping.template_purchase[column_purchase[3]].tolist(),
            ["45100", "45100", "45100.25"],
        )


class DiscountSupplementNetworkTest(unittest.TestCase):
    def setUp(self):
        self.etl = Template_ETL([])
        network = pd.DataFrame({
            "SITE": ["201", "202", "839", "801"],
            "NATIONAL_SITE": ["8710", "8710", "8300", "8300"],
            "GROUP_SITE": ["8710", "8710", "8300", "8300"],
            "REGION_SITE": ["8710", "8710", "8300", "8300"],
            "ACTIVE": ["0", "1", "0", "1"],
            "DISCOUNT": ["1", "0", "1", "0"],
        })
        with patch("src.service.template_service.pd.read_excel", return_value=network):
            self.etl._load_network()
        self.etl.src = pd.DataFrame({
            "GOLD CODE": ["GC1"], "LV": ["1"], "STRUCTURE": ["110"],
            "NORMAL PURCHASE PRICE": [100], "DISCOUNT_NETWORK_EXPANDED": ["101"],
            "PP START DATE": [pd.Timestamp("2099-01-01")],
            "PP END DATE": [pd.Timestamp("2099-01-02")],
            "COMMERCIAL CONTRACT": ["CONT"], "PURCHASE VAT": ["10%"],
            "SUPPLIER CODE": ["12345"], "DISCOUNT (% OR VALUE)": ["10%"],
        })

    def assert_sites(self, expected_purchase, expected_discount):
        purchase = Template_Mapping(self.etl)._create_purchase().template_purchase
        discount = Discount(self.etl)._create_ag_raw()._create_ag()
        self.assertCountEqual(purchase["SITE"].tolist(), expected_purchase)
        for frame in (discount.template_ag_raw, discount.template_ag):
            self.assertCountEqual(frame["SITE"].tolist(), expected_discount)

    def test_supplements_use_discount_flags_independently_of_active(self):
        self.assert_sites(["101", "201", "839"], ["101", "201", "839"])
        self.assertEqual(self.etl.dict_network["store_minigo"], ["202"])
        self.assertEqual(self.etl.dict_network["wh"], ["801"])

    def test_warehouse_discount_conditions_and_zero_discount_are_preserved(self):
        for value, exception, sites in (
            ("10", False, ["101", "201", "839"]),
            ("10+5", False, ["101", "201"]),
            ("10T+5TH", False, ["101", "201", "839"]),
            ("10+5", True, ["101", "201", "839"]),
            ("0", False, []),
        ):
            with self.subTest(value=value, exception=exception):
                self.etl.src["DISCOUNT (% OR VALUE)"] = value
                self.etl.exception_discount_gold_codes = {"GC1"} if exception else set()
                self.assert_sites(["101", "201", "839"], sites)

    def test_missing_or_empty_groups_do_not_fall_back_to_active_sites(self):
        for groups in (None, {}, {"8710": " ; ", "8300": ""},
                       {"8710": None, "8300": None}):
            with self.subTest(groups=groups):
                if groups is None:
                    self.etl.dict_network.pop("DISCOUNT_NETWORK", None)
                else:
                    self.etl.dict_network["DISCOUNT_NETWORK"] = groups
                self.assert_sites(["101"], ["101"])

    def test_existing_sites_and_empty_tokens_do_not_create_extra_rows(self):
        self.etl.src["DISCOUNT_NETWORK_EXPANDED"] = "101;201;839"
        self.etl.dict_network["DISCOUNT_NETWORK"] = {
            "8710": " ;201;; ", "8300": "839; ;839;",
        }
        self.assert_sites(["101", "201", "839"], ["101", "201", "839"])

    def test_wh_ag_does_not_add_configured_supplements(self):
        discount = Discount(self.etl, ag_type="WH")._create_ag_raw()._create_ag()
        self.assertEqual(discount.template_ag["SITE"].tolist(), ["101"])


class DiscountRawRestoreTest(unittest.TestCase):
    def test_ag_number_normalization_pads_to_thirteen_characters(self):
        self.assertEqual(Discount._normalize_ag_number(" 12345 "), "0000000012345")
        self.assertEqual(Discount._normalize_ag_number("1234567890123"), "1234567890123")
        self.assertEqual(Discount._normalize_ag_number("12345678901234"), "12345678901234")
        self.assertEqual(Discount._normalize_ag_number("   "), "")
        self.assertEqual(Discount._normalize_ag_number(None), "")

    def test_created_ag_and_ag_raw_use_padded_ag_number(self):
        etl = SimpleNamespace(
            src=pd.DataFrame(
                {
                    "STRUCTURE": ["1"],
                    "PURCHASE NETWORK EXPANDED": ["999"],
                    "DISCOUNT_NETWORK_EXPANDED": ["101"],
                    "SUPPLIER CODE": ["12345"],
                    "COMMERCIAL CONTRACT": ["CONT1234"],
                    "GOLD CODE": ["GC1"],
                    "LV": ["1"],
                    "PP START DATE": [pd.Timestamp("2026-10-01")],
                    "PP END DATE": [pd.Timestamp("2026-10-02")],
                    "DISCOUNT (% OR VALUE)": ["10%"],
                }
            ),
            dict_network={"store_minigo": [], "wh": [], "wh8": []},
            exception_discount_gold_codes=set(),
            cata="C01",
        )

        discount = Discount(etl)._create_ag_raw()._create_ag()

        self.assertEqual(discount.template_ag_raw["SITE"].tolist(), ["101"])
        self.assertEqual(discount.template_ag_raw["AG NO"].tolist(), ["0012345010101"])
        self.assertEqual(discount.template_ag["AG NO"].tolist(), ["0012345010101"])

    def test_update_normalizes_report_agno_and_keeps_row_for_dc_and_de(self):
        discount = Discount(None)
        discount.template_ag_raw = pd.DataFrame(
            {
                "ACTION": ["0"],
                "DEPARTMENT": ["010"],
                "SITE": ["0101"],
                "SUPPLIER": ["SUP"],
                "CONTRACT": ["CONT"],
                "AG NO": [" 12345 "],
                "AG CODE": [""],
                "AG DESCRIPTION": ["Description"],
                "AG START DATE": ["01/10/2026"],
                "AG END DATE": ["02/10/2026"],
                "GOLD CODE": ["GC1"],
                "LV": ["1"],
                "ARTICLE START DATE": ["01/10/2026"],
                "ARTICLE END DATE": ["02/10/2026"],
                "LEVEL OF VALUE": ["0"],
                "ERROR": [""],
                "DISCOUNT VALUE": ["10%"],
                "RAW START DATE": ["01/10/2026"],
                "RAW END DATE": ["02/10/2026"],
                "DISCOUNT TYPE": ["1"],
            }
        )
        report = pd.DataFrame(
            {
                "ERRORMESS": [""],
                "AG_CODE": ["AG-CODE-1"],
                "ACTION": ["0"],
                "DEPT": ["010"],
                "SITE": ["0101"],
                "SUPPLIER_CODE": ["SUP"],
                "COMERCIAL_CONTRACT": ["CONT"],
                "AGNO1": ["12345"],
                "AG_DESC": ["Description"],
                "AG_START_DATE": ["01/10/2026"],
                "AG_END_DATE": ["02/10/2026"],
                "ARTICLE_CODE": ["GC1"],
                "LV": ["1"],
                "ARTICLE_START_DATE": ["01/10/2026"],
                "ARTICLE_END_DATE": ["02/10/2026"],
            }
        )

        with patch("src.service.template_mapping.pd.read_excel", return_value=report):
            discount._update("report.xlsx")._create_dc()._create_de()

        self.assertEqual(discount.template_ag_raw["AG NO"].tolist(), ["0000000012345"])
        self.assertEqual(discount.template_ag_raw["AG CODE"].tolist(), ["AG-CODE-1"])
        self.assertEqual(len(discount.template_dc), 1)
        self.assertEqual(len(discount.template_de), 1)

    def test_de_removes_insignificant_decimal_zeros(self):
        discount = Discount(None)
        discount.template_ag_raw = pd.DataFrame(
            {
                "DISCOUNT TYPE": ["2", "1", "2"],
                "SITE": ["101", "101", "101"],
                "AG CODE": ["AG1", "AG2", "AG3"],
                "GOLD CODE": ["GC1", "GC2", "GC3"],
                "LV": ["1", "1", "1"],
                "DISCOUNT VALUE": ["25000.00", "10.50%", "25000"],
                "RAW START DATE": ["01/01/2026"] * 3,
                "RAW END DATE": ["02/01/2026"] * 3,
            }
        )

        discount._create_de()

        self.assertEqual(
            discount.template_de["VALUE ON INVOICE"].tolist(),
            ["25000", "10.5", "25000"],
        )

    def test_dc_combines_501_and_201_rows(self):
        discount = Discount(None)
        discount.template_ag_raw = pd.DataFrame(
            {
                "DISCOUNT TYPE": ["3", "2"],
                "SITE": ["101", "102"],
                "SUPPLIER": ["SUP", "SUP"],
                "CONTRACT": ["CON", "CON"],
                "AG CODE": ["AG1", "AG2"],
                "AG DESCRIPTION": ["Free", "Money"],
                "ARTICLE START DATE": ["01/01/2026"] * 2,
                "ARTICLE END DATE": ["02/01/2026"] * 2,
                "GOLD CODE": ["GC1", "GC2"],
                "LV": ["1", "1"],
                "RAW START DATE": ["01/01/2026"] * 2,
                "RAW END DATE": ["02/01/2026"] * 2,
                "DISCOUNT VALUE": ["1T+1T", "25000"],
            }
        )

        discount._create_dc()

        self.assertEqual(discount.template_dc.columns.tolist(), ["NO", *column_dc])
        self.assertEqual(set(discount.template_dc["DISCOUNT TYPE"]), {"501", "201"})
        self.assertEqual(len(discount.template_dc), 2)

    def test_restores_exported_ag_raw_for_dc_de_generation(self):
        raw = pd.DataFrame(
            {
                column: ["value"]
                for column in Discount.RAW_COLUMNS
            }
        )

        with patch("src.service.template_mapping.pd.read_excel", return_value=raw):
            discount = Discount.from_ag_raw_file("template_ag_raw.xls")

        self.assertIsNone(discount.etl)
        self.assertTrue(discount.template_ag_raw.equals(raw))

    def test_rejects_ag_raw_missing_processing_columns(self):
        with patch("src.service.template_mapping.pd.read_excel", return_value=pd.DataFrame()):
            with self.assertRaisesRegex(ValueError, "AG raw file is missing required columns"):
                Discount.from_ag_raw_file("template_ag_raw.xls")


if __name__ == "__main__":
    unittest.main()

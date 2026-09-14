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
from src.service.template_mapping import AttributeMapMixin, Discount, Template_Mapping


class AttributeMediumMappingTest(unittest.TestCase):
    def test_maps_all_supported_medium_labels_to_numeric_codes(self):
        mapper = AttributeMapMixin()
        expected = {
            "FRONT PAGE": 1,
            "BACK PAGE": 2,
            "UNBEAT": 3,
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
            {"GOLD CODE": "GC2", "LV": "2", "SO": "SO2"}
        ])
        self.assertEqual(
            mapping.template_add_attribute_marketing["GOLD CODE"].tolist(),
            ["GC1", "GC3"],
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
    def test_formats_whole_prices_and_removes_float_noise(self):
        etl = SimpleNamespace(
            src=pd.DataFrame(
                {
                    "GOLD CODE": ["02043862", "02043863", "02043864"],
                    "LV": ["1", "1", "1"],
                    "NORMAL PURCHASE PRICE": [45100.0, 45100.0000000001, 45100.25],
                    "PURCHASE NETWORK EXPANDED": ["S1", "S1", "S1"],
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


class DiscountRawRestoreTest(unittest.TestCase):
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

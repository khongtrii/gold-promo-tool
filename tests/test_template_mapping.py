import unittest
from types import SimpleNamespace

import pandas as pd

from src.constant.template import column_po_commitment, column_promotion_plan
from src.service.template_mapping import AttributeMapMixin, Template_Mapping


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


if __name__ == "__main__":
    unittest.main()

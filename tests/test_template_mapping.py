import unittest
from types import SimpleNamespace

import pandas as pd

from src.constant.template import column_promotion_plan
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


if __name__ == "__main__":
    unittest.main()

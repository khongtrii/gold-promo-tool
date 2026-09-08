import unittest

import pandas as pd

from src.service.template_service import Template_ETL


class DiscountParsingTest(unittest.TestCase):
    def test_accepts_supported_discount_formats(self):
        valid_values = [
            "10",
            "10.5",
            "10,5",
            "10%",
            "10 + 2",
            "10T + 2T",
            "10th+2TH",
            "10t+2th",
        ]

        for value in valid_values:
            normalized = value.replace(" ", "")
            with self.subTest(value=value):
                self.assertIsNotNone(
                    Template_ETL.VALID_DISCOUNT.fullmatch(normalized)
                )

    def test_rejects_percentage_free_goods_and_partial_unit_formats(self):
        invalid_values = [
            "10%+2%",
            "10%+2",
            "10+2%",
            "10T+2",
            "10+2TH",
            "abc",
            "",
        ]

        for value in invalid_values:
            with self.subTest(value=value):
                self.assertIsNone(Template_ETL.VALID_DISCOUNT.fullmatch(value))

    def test_only_plain_number_plus_number_is_non_warehouse(self):
        data = pd.DataFrame(
            {
                "DISCOUNT (% OR VALUE)": [
                    "10+2",
                    "10 + 2",
                    "10T+2T",
                    "10TH+2TH",
                    "10%+2%",
                    "10%",
                    "10",
                ]
            }
        )

        result = Template_ETL._get_non_warehouse_src(data)

        self.assertEqual(result.index.tolist(), [0, 1])


if __name__ == "__main__":
    unittest.main()

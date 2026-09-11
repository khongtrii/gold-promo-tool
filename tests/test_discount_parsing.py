import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import Workbook

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
                "GOLD CODE": [
                    "NORMAL",
                    "NORMAL",
                    "NORMAL",
                    "NORMAL",
                    "NORMAL",
                    "NORMAL",
                    "NORMAL",
                ],
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

    def test_warehouse_exception_is_not_reported_as_non_warehouse(self):
        data = pd.DataFrame(
            {
                "GOLD CODE": ["02043862", "99999999"],
                "DISCOUNT (% OR VALUE)": ["10+2", "10+2"],
            }
        )

        result = Template_ETL._get_non_warehouse_src(data, {"02043862"})

        self.assertEqual(result["GOLD CODE"].tolist(), ["99999999"])

    def test_invalid_site_rule_skips_following_network_checks_for_its_group(self):
        etl = Template_ETL([])
        etl.dict_network = {"network": {}, "store": ["1001"]}
        data = pd.DataFrame(
            {
                "FILE NAME": ["source.xlsx"],
                "GOLD CODE": ["02043862"],
                "LV": ["1"],
                "LU": ["1"],
                "PURCHASE NETWORK": ["1001++"],
                "GOLD PROMO NETWORK": ["9999"],
            }
        )

        result = etl._check_network(data)
        result = etl._ppNetwork_gpNetwork(result)
        note = result.at[0, "NOTE ERR FROM MASTER DATA"]

        self.assertIn("PURCHASE NETWORK không đúng SITE RULE.", note)
        self.assertNotIn("Cửa hàng", note)
        self.assertNotIn("Thiếu cửa hàng", note)
        self.assertNotIn("Dư cửa hàng", note)

    def test_network_8000_skips_all_network_validation_for_its_group(self):
        etl = Template_ETL([])
        etl.dict_network = {"network": {}, "store": ["1001"]}
        data = pd.DataFrame(
            {
                "FILE NAME": ["source.xlsx", "source.xlsx"],
                "GOLD CODE": ["02043862", "02043863"],
                "LV": ["1", "1"],
                "LU": ["1", "1"],
                "PURCHASE NETWORK": ["8000", "1001"],
                "GOLD PROMO NETWORK": ["9999", "8000"],
            }
        )

        result = etl._check_network(data)
        result = etl._ppNetwork_gpNetwork(result)
        for index in result.index:
            self.assertEqual(result.at[index, "NOTE ERR FROM MASTER DATA"], "Không được sử dụng network 8000.")
            self.assertEqual(result.at[index, "PURCHASE NETWORK EXPANDED"], "")
            self.assertEqual(result.at[index, "GOLD PROMO NETWORK EXPANDED"], "")

    def test_restores_percentage_formatted_source_cells(self):
        columns = [
            "DISCOUNT (% OR VALUE)",
            "% DELIVERY 1",
            "% DELIVERY 2",
            "% DELIVERY 3",
        ]
        with TemporaryDirectory() as directory:
            path = Path(directory) / "source.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Template"
            for column_index, column in enumerate(columns, start=1):
                sheet.cell(7, column_index, column)
            for column_index, value in enumerate([0.5, 0.5, 0.3, 0.2], start=1):
                cell = sheet.cell(8, column_index, value)
                cell.number_format = "0%"
            sheet.cell(9, 1, "50%")
            workbook.save(path)

            data = pd.DataFrame(
                {
                    "DISCOUNT (% OR VALUE)": ["0.5", "50%"],
                    "% DELIVERY 1": ["0.5", ""],
                    "% DELIVERY 2": ["0.3", ""],
                    "% DELIVERY 3": ["0.2", ""],
                    "NORMAL DISCOUNT ON PURCHASE (%)": ["0.5", "0.5"],
                }
            )
            Template_ETL._restore_percentage_cells(data, path)

        self.assertEqual(data.loc[0, "DISCOUNT (% OR VALUE)"], "50%")
        self.assertEqual(data.loc[0, "% DELIVERY 1"], "50%")
        self.assertEqual(data.loc[0, "% DELIVERY 2"], "30%")
        self.assertEqual(data.loc[0, "% DELIVERY 3"], "20%")
        self.assertEqual(data.loc[1, "DISCOUNT (% OR VALUE)"], "50%")
        self.assertEqual(data.loc[0, "NORMAL DISCOUNT ON PURCHASE (%)"], "0.5")
        self.assertIsNotNone(Template_ETL.VALID_DISCOUNT.fullmatch(data.loc[0, "DISCOUNT (% OR VALUE)"]))

        allocation = data.iloc[[0]].copy()
        allocation["GOLD CODE"] = "02043862"
        allocation["LV"] = "1"
        allocation["LU"] = "1"
        allocation["SUPPLIER CODE"] = "SUP"
        allocation["COMMERCIAL CONTRACT"] = "CONT"
        allocation["PURCHASE NETWORK EXPANDED"] = ""
        result = Template_ETL([])._check_allocation(allocation)
        self.assertNotIn("Tổng % DELIVERY", result.at[0, "NOTE ERR FROM MASTER DATA"])


if __name__ == "__main__":
    unittest.main()

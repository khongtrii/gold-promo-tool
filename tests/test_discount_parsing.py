import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import Workbook

from src.constant.required import required_stage1
from src.service.template_service import Template_ETL


class DiscountParsingTest(unittest.TestCase):
    def test_pp_dates_follow_delivery_type_order_dates_from_plan(self):
        etl = Template_ETL([])
        etl.cata = "C01"
        etl.plan = pd.DataFrame(
            {
                "CROSS-DOCKING | ORDER DATE 1": [pd.Timestamp("2026-10-01")],
                "CROSS-DOCKING | ORDER DATE 3": [pd.Timestamp("2026-10-03")],
                "DIRECT | ORDER DATE 1": [pd.Timestamp("2026-10-04")],
                "DIRECT | ORDER DATE 3": [pd.Timestamp("2026-10-06")],
                "VINAMILK | ORDER DATE 1": [pd.Timestamp("2026-10-07")],
                "VINAMILK | ORDER DATE 3": [pd.Timestamp("2026-10-09")],
            }
        )
        data = pd.DataFrame(
            {
                "DELIVERY TYPE": [
                    "CROSS-DOCKING",
                    "DIRECT",
                    "VINAMILK",
                    "DIRECT",
                ],
                "PP START DATE": pd.to_datetime(
                    ["2026-10-01", "2026-10-05", "2026-10-07", None]
                ),
                "PP END DATE": pd.to_datetime(
                    ["2026-10-03", "2026-10-05", "2026-10-10", None]
                ),
            }
        )

        result = etl._validate_pp_dates_against_plan(data)

        self.assertEqual(result.at[0, "NOTE ERR FROM MASTER DATA"], "")
        self.assertIn(
            "PP START DATE phải bằng DIRECT | ORDER DATE 1 (04/10/2026).",
            result.at[1, "NOTE ERR FROM MASTER DATA"],
        )
        self.assertIn(
            "PP END DATE phải lớn hơn hoặc bằng DIRECT | ORDER DATE 3 (06/10/2026).",
            result.at[1, "NOTE ERR FROM MASTER DATA"],
        )
        self.assertEqual(result.at[2, "NOTE ERR FROM MASTER DATA"], "")
        self.assertEqual(result.at[3, "NOTE ERR FROM MASTER DATA"], "")

    def test_blank_contract_reports_required_data_and_skips_contract_normalization(self):
        etl = Template_ETL([])
        etl.dict_network = {"wh": []}
        data = pd.DataFrame(
            {
                "COMMERCIAL CONTRACT": [None, "   ", "ABCD1234"],
                "NOTE ERR FROM MASTER DATA": ["", "", ""],
            }
        )

        etl._check_required_data(data, ["COMMERCIAL CONTRACT"])
        data["COMMERCIAL CONTRACT"] = data["COMMERCIAL CONTRACT"].map(
            etl._contract_checking
        )

        self.assertEqual(data["COMMERCIAL CONTRACT"].tolist(), ["", "", "ABCD"])
        self.assertEqual(
            data["NOTE ERR FROM MASTER DATA"].tolist(),
            [
                "Các cột bắt buộc đang để trống: COMMERCIAL CONTRACT",
                "Các cột bắt buộc đang để trống: COMMERCIAL CONTRACT",
                "",
            ],
        )

    def test_discount_difference_allows_only_type_1_or_2_with_type_3(self):
        allowed = [
            pd.Series(["10%", "10+2"]),
            pd.Series(["10000", "10TH+2TH"]),
        ]
        rejected = [
            pd.Series(["10%", "20%"]),
            pd.Series(["10000", "20000"]),
            pd.Series(["10+2", "20TH+3TH"]),
            pd.Series(["10%", "10000"]),
            pd.Series(["10%", "20%", "10+2"]),
        ]

        for discounts in allowed:
            with self.subTest(discounts=discounts.tolist()):
                self.assertFalse(Template_ETL._has_invalid_discount_difference(discounts))
        for discounts in rejected:
            with self.subTest(discounts=discounts.tolist()):
                self.assertTrue(Template_ETL._has_invalid_discount_difference(discounts))

    def test_overlapping_network_applies_discount_type_exception(self):
        etl = Template_ETL([])
        base = {
            "GOLD CODE": ["02043862", "02043862"],
            "LV": ["1", "1"],
            "PURCHASE NETWORK EXPANDED": ["1001", "1001"],
            "NORMAL PURCHASE PRICE": ["10000", "10000"],
        }

        allowed = etl._validate_overlapping_price_or_discount(
            pd.DataFrame(base | {"DISCOUNT (% OR VALUE)": ["10%", "10+2"]})
        )
        rejected = etl._validate_overlapping_price_or_discount(
            pd.DataFrame(base | {"DISCOUNT (% OR VALUE)": ["10%", "10000"]})
        )

        self.assertTrue(allowed["NOTE ERR FROM MASTER DATA"].eq("").all())
        self.assertTrue(
            rejected["NOTE ERR FROM MASTER DATA"].str.contains(
                "Thông tin mua hàng và chiết khấu bị trùng", regex=False
            ).all()
        )

    def test_missing_allocations_are_combined_for_every_row_in_group(self):
        etl = Template_ETL([])
        etl.dict_network = {"store": ["101", "102"]}
        data = pd.DataFrame(
            {
                "GOLD CODE": ["GC1", "GC1"],
                "LV": ["1", "1"],
                "LU": ["1", "1"],
                "SUPPLIER CODE": ["SUP1", "SUP2"],
                "COMMERCIAL CONTRACT": ["CON1", "CON2"],
                "PURCHASE NETWORK EXPANDED": ["101", "102"],
                "% DELIVERY 1": ["100", "100"],
                "% DELIVERY 2": ["", ""],
                "% DELIVERY 3": ["", ""],
                # Values on another row must not satisfy the current row's
                # PURCHASE NETWORK allocation requirement.
                "101": ["", "1"],
                "102": ["1", ""],
            }
        )

        result = etl._check_allocation(data)

        expected = (
            "Thiếu phân bổ đối với các cửa hàng: 101;102 "
            "dựa trên PURCHASE NETWORK EXPANDED."
        )
        self.assertEqual(result["NOTE ERR FROM MASTER DATA"].tolist(), [expected, expected])

    def test_blank_discount_defaults_to_zero_percent(self):
        data = pd.DataFrame({"DISCOUNT (% OR VALUE)": [None, "", "  ", "10%"]})

        result = Template_ETL._default_blank_discounts(data)

        self.assertEqual(
            result["DISCOUNT (% OR VALUE)"].tolist(),
            ["0%", "0%", "0%", "10%"],
        )

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

    def test_stage2_validates_vat_and_normalizes_integral_sale_price(self):
        data = pd.DataFrame(
            {
                "SALE VAT": ["10%", " kkkt ", "invalid", "0.1"],
                "PROMOTION SALE PRICE": [
                    "45100.0",
                    "45100.0000000001",
                    "45100.5",
                    "not a number",
                ],
            }
        )

        result = Template_ETL([])._validate_stage2_sale_values(data)

        self.assertEqual(result["SALE VAT"].tolist()[:2], ["10%", "KKKT"])
        self.assertEqual(result["PROMOTION SALE PRICE"].tolist()[:2], ["45100", "45100"])
        self.assertIn("SALE VAT chỉ được phép", result.at[2, "NOTE ERR FROM MASTER DATA"])
        self.assertIn("SALE VAT chỉ được phép", result.at[3, "NOTE ERR FROM MASTER DATA"])
        self.assertIn("PROMOTION SALE PRICE phải là số nguyên", result.at[2, "NOTE ERR FROM MASTER DATA"])
        self.assertIn("PROMOTION SALE PRICE phải là số nguyên", result.at[3, "NOTE ERR FROM MASTER DATA"])

    def test_free_product_is_required_only_with_check_attribute(self):
        required_without_free_product = [
            column
            for column in required_stage1
            if column != "FREE PRODUCT"
        ]
        data = pd.DataFrame(
            {column: ["value"] for column in required_without_free_product}
            | {"FREE PRODUCT": [""]}
        )
        data["NOTE ERR FROM MASTER DATA"] = ""

        Template_ETL._check_required_data(data, required_without_free_product)
        self.assertEqual(data.at[0, "NOTE ERR FROM MASTER DATA"], "")

        Template_ETL._check_required_data(data, required_stage1)
        self.assertIn("FREE PRODUCT", data.at[0, "NOTE ERR FROM MASTER DATA"])


if __name__ == "__main__":
    unittest.main()

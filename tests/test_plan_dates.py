import unittest
from unittest.mock import patch

import pandas as pd

from src.constant.required import date_columns_plan
from src.service.template_service import Template_ETL


class PlanDateTest(unittest.TestCase):
    def test_load_plan_keeps_dates_and_parses_day_first(self):
        source = {
            column: ["02/10/2026"]
            for column in date_columns_plan
        }
        source["SHOP ACTIVE START"] = ["05/10/2026"]
        source.update(
            {
                "CATALOGUE": ["C01"],
                "CATALOGUE DESCRIPTION": ["Catalogue"],
            }
        )
        etl = Template_ETL([], path_plan="plan.xlsx")
        etl.cata = "C01"

        with patch(
            "src.service.template_service.pd.read_excel",
            return_value=pd.DataFrame(source),
        ):
            etl._load_plan()

        catalogue_date = etl.plan["CATALOGUE START DATE"].iat[0]
        self.assertIsInstance(catalogue_date, pd.Timestamp)
        self.assertEqual((catalogue_date.day, catalogue_date.month), (2, 10))
        shop_active_start = etl.plan["SHOP ACTIVE START"].iat[0]
        shop_activation = etl.plan["SHOP ACTIVATION"].iat[0]
        self.assertIsInstance(shop_activation, pd.Timestamp)
        self.assertEqual(shop_activation, shop_active_start)

if __name__ == "__main__":
    unittest.main()

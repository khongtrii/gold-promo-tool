import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import Workbook, load_workbook

from src.desktop_app import GoldPromoApp, WorkbookExporter


class WorkbookExporterTest(unittest.TestCase):
    def test_discount_output_numbers_continue_after_other_templates(self):
        output = Path("output")
        timestamp = "160926_120000"

        expected = {
            "template_ag": "8_template_ag_160926_120000.xls",
            "template_ag_raw": "8_template_ag_raw_160926_120000.xls",
            "report_ag_errors": "8_report_ag_errors_160926_120000.xls",
            "template_dc": "9_template_dc_160926_120000.xls",
            "template_de": "10_template_de_160926_120000.xls",
        }
        for name, file_name in expected.items():
            with self.subTest(name=name):
                self.assertEqual(
                    GoldPromoApp._output_file(output, name, timestamp),
                    output / file_name,
                )

    def test_source_error_notes_use_yellow_fill_and_red_font(self):
        with TemporaryDirectory(dir=Path.cwd()) as directory:
            source = Path(directory) / "source.xlsx"
            output = Path(directory) / "errors.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Template"
            sheet.cell(7, 1, "MULTI")
            workbook.save(source)
            workbook.close()

            WorkbookExporter.write_source_errors(
                source,
                pd.DataFrame(
                    {
                        "_SOURCE_ROW": [8],
                        "NOTE ERR FROM MASTER DATA": ["Validation error"],
                    }
                ),
                output,
            )

            result = load_workbook(output)
            try:
                result_sheet = result["Template"]
                for row in (7, 8):
                    cell = result_sheet.cell(row, 2)
                    self.assertEqual(cell.fill.fill_type, "solid")
                    self.assertEqual(cell.fill.fgColor.rgb, "00FFFF00")
                    self.assertEqual(cell.font.color.type, "rgb")
                    self.assertEqual(cell.font.color.rgb, "FFFF0000")
            finally:
                result.close()


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import importlib.util
import io
import sys
import types
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook


def load_app_module():
    candidates = [
        path for path in Path(__file__).parent.glob("global_inventory_streamlit*.py")
        if not path.name.startswith("test_")
    ]
    if len(candidates) != 1:
        raise RuntimeError(f"실행 파일을 하나만 두어야 합니다: {candidates}")
    if "streamlit" not in sys.modules:
        try:
            import streamlit  # noqa: F401
        except ModuleNotFoundError:
            sys.modules["streamlit"] = types.ModuleType("streamlit")
    spec = importlib.util.spec_from_file_location("inventory_app_for_test", candidates[0])
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


APP = load_app_module()


def workbook_bytes(headers, rows, sheet_name):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for row in rows:
        ws.append(row)
    output = io.BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def default_settings(locations=("DU", "GD", "GL")):
    return APP.Settings(
        locations=locations,
        pass_threshold=0.66,
        keep_missing_expiry=True,
        keep_missing_master=True,
        include_zero_in_details=False,
        expiry_ascending=True,
        quantity_ascending=True,
        lot_ascending=True,
        stock_sheet="Sheet1",
        master_sheet="유효기간",
        columns=APP.ColumnNames(),
    )


class InventoryAutomationTests(unittest.TestCase):
    def setUp(self):
        self.master = workbook_bytes(
            ["바코드", "개봉 전 사용기한"],
            [["00123", 30], ["456", "2년"], ["789", None]],
            "유효기간",
        )

    def test_supported_date_formats(self):
        expected = date(2028, 8, 1)
        for value in ("20280801", "2028년 8월 1일", "2028-08-01 00:00:00"):
            with self.subTest(value=value):
                self.assertEqual(APP.to_date_only(value), expected)
        self.assertIsNone(APP.to_date_only("미확인"))
        self.assertIsNone(APP.parse_months("N/A"))

    def test_invalid_quantity_is_blocked(self):
        for value in ("NOT-A-NUMBER", -1, 1.5, "NaN", "inf", None):
            with self.subTest(value=value):
                if value is None:
                    self.assertIsNone(APP.safe_number(value))
                else:
                    with self.assertRaises(ValueError):
                        APP.safe_number(value)

    def test_unicode_arrow_location_input(self):
        self.assertEqual(APP.parse_locations("DU → GD → GL"), ("DU", "GD", "GL"))

    def test_generation_and_quantity_reconciliation(self):
        stock = workbook_bytes(
            ["저장 위치", "자재", "자재내역", "가용재고(ASN)", "유통기한", "로트번호(생산)"],
            [
                ["DU", "00123", "+HYPERLINK(\"https://example.com\")", 100, date(2028, 8, 1), "L2"],
                ["DU", "00123", "A 변경", 10, "20280801", "L1"],
                ["GD", "456", "B", "1,234", "2028년 8월 1일", "L3"],
                ["GL", "789", "C", 0, None, None],
            ],
            "Sheet1",
        )
        output, source_rows, kept_rows, summary = APP.create_inventory_workbook(
            stock, self.master, "stock_0820.xlsx", "master.xlsx",
            date(2026, 8, 15), "test", date(2026, 8, 20), default_settings(),
        )
        self.assertEqual(summary.source_quantity, 1344)
        self.assertEqual(summary.kept_quantity, 1344)
        self.assertEqual(len(source_rows), 4)
        self.assertEqual(len(kept_rows), 4)

        wb = load_workbook(io.BytesIO(output), data_only=True)
        try:
            self.assertEqual(wb["글로벌"]["C2"].data_type, "s")
            self.assertIsInstance(wb["글로벌"]["G2"].value, float)
            self.assertEqual(wb["글로벌"]["I2"].value, "통과")
            pivot = wb["글로벌 pv"]
            self.assertEqual(pivot.cell(pivot.max_row, 6).value, 1344)
        finally:
            wb.close()

    def test_malformed_rows_are_blocked(self):
        stock = workbook_bytes(
            ["저장 위치", "자재", "자재내역", "가용재고(ASN)", "유통기한", "로트번호(생산)"],
            [
                ["DU", "00123", "A", "NOT-A-NUMBER", "2028-08-01", "L1"],
                ["DU", "00123", "A", 1, "not-a-date", "L1,L2"],
            ],
            "Sheet1",
        )
        with self.assertRaises(APP.DataValidationError):
            APP.create_inventory_workbook(
                stock, self.master, "bad.xlsx", "master.xlsx",
                date(2026, 8, 15), "test", None, default_settings(),
            )

    def test_table_names_do_not_collide(self):
        stock = workbook_bytes(
            ["저장 위치", "자재", "자재내역", "가용재고(ASN)", "유통기한", "로트번호(생산)"],
            [
                ["D-U", "00123", "A", 1, "2028-08-01", "L1"],
                ["DU", "00123", "A", 1, "2028-08-01", "L1"],
            ],
            "Sheet1",
        )
        output, *_ = APP.create_inventory_workbook(
            stock, self.master, "stock.xlsx", "master.xlsx",
            date(2026, 8, 15), "test", None, default_settings(("D-U", "DU")),
        )
        self.assertTrue(output)


if __name__ == "__main__":
    unittest.main()

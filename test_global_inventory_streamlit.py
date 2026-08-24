from __future__ import annotations

import importlib.util
import inspect
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
            streamlit_stub = types.ModuleType("streamlit")
            streamlit_stub.cache_data = lambda **_kwargs: (lambda func: func)
            sys.modules["streamlit"] = streamlit_stub
    spec = importlib.util.spec_from_file_location("inventory_app_for_test", candidates[0])
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


APP = load_app_module()


def workbook_bytes(headers, rows, sheet_name, header_row=1):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for _ in range(1, header_row):
        ws.append([])
    ws.append(headers)
    for row in rows:
        ws.append(row)
    output = io.BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def default_settings(locations=("GD", "DU", "GL")):
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
            header_row=12,
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

    def test_default_warehouse_order_and_master_header_row(self):
        sidebar_source = inspect.getsource(APP.render_sidebar)
        self.assertIn('value="GD, DU, GL"', sidebar_source)
        self.assertEqual(APP.MASTER_HEADER_ROW, 12)
        summary = APP.ValidationSummary()
        mapping = APP.load_lifetime_months(self.master, default_settings(), summary)
        self.assertEqual(mapping["00123"], 30)
        self.assertEqual(mapping["456"], 24)
        self.assertEqual(summary.master_rows, 3)

        wrong_header_master = workbook_bytes(
            ["바코드", "개봉 전 사용기한"],
            [["00123", 30]],
            "유효기간",
            header_row=1,
        )
        with self.assertRaises(APP.InputFormatError):
            APP.load_lifetime_months(
                wrong_header_master,
                default_settings(),
                APP.ValidationSummary(),
            )

    def test_streamlit_options_are_batched_in_forms(self):
        sidebar_source = inspect.getsource(APP.render_sidebar)
        main_source = inspect.getsource(APP.main)
        self.assertIn('st.form("inventory_options_form"', sidebar_source)
        self.assertIn('st.form_submit_button(', sidebar_source)
        self.assertNotIn("disabled=base_mode", sidebar_source)
        self.assertIn('st.form("inventory_generation_form"', main_source)
        self.assertIn("generate_submitted and not base_confirmed", main_source)
        self.assertIn('on_click="ignore"', main_source)

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
            self.assertEqual(wb["글로벌"]["D2"].number_format, "#,##0")
            self.assertEqual(wb["글로벌"]["E2"].number_format, "yyyy-mm-dd")
            self.assertEqual(wb["글로벌"]["G2"].number_format, "0.0%")
            pivot = wb["글로벌 pv"]
            self.assertEqual(
                [pivot.cell(2, column).value for column in range(3, 6)],
                ["GD", "DU", "GL"],
            )
            self.assertEqual(pivot.cell(pivot.max_row, 6).value, 1344)
            self.assertEqual(pivot.cell(pivot.max_row, 6).number_format, "#,##0")
            detail = wb["글로벌 유통기한 DU"]
            self.assertEqual(detail.cell(detail.max_row, 3).number_format, "#,##0")
        finally:
            wb.close()

    def test_all_location_aggregation_preserves_totals(self):
        rows = [
            {
                "저장 위치": "DU", "자재": "100", "자재내역": "A",
                "가용재고(ASN)": 3, "유통기한": date(2028, 1, 1),
                "로트번호(생산)": "L1", "잔여 유통기한율": 0.8,
                "개봉 전 사용기한(개월)": 30,
            },
            {
                "저장 위치": "DU", "자재": "100", "자재내역": "A",
                "가용재고(ASN)": 2, "유통기한": date(2028, 1, 1),
                "로트번호(생산)": "L1", "잔여 유통기한율": 0.8,
                "개봉 전 사용기한(개월)": 30,
            },
            {
                "저장 위치": "GD", "자재": "200", "자재내역": "B",
                "가용재고(ASN)": 7, "유통기한": date(2028, 2, 1),
                "로트번호(생산)": "L2", "잔여 유통기한율": 0.9,
                "개봉 전 사용기한(개월)": 30,
            },
        ]
        result = APP.aggregate_all_location_exp_rows(rows, default_settings(("DU", "GD")))
        self.assertEqual(result["DU"][0]["수량"], 5)
        self.assertEqual(result["GD"][0]["수량"], 7)
        self.assertEqual(sum(item["수량"] for values in result.values() for item in values), 12)

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

    def test_detail_sort_priority_is_preserved(self):
        rows = []
        for expiry, quantity, lot in (
            (date(2028, 2, 1), 5, "L2"),
            (date(2028, 1, 1), 10, "L3"),
            (date(2028, 1, 1), 2, "L2"),
            (date(2028, 1, 1), 2, "L1"),
        ):
            rows.append({
                "저장 위치": "DU",
                "자재": "00123",
                "자재내역": "A",
                "가용재고(ASN)": quantity,
                "유통기한": expiry,
                "로트번호(생산)": lot,
                "잔여 유통기한율": 0.8,
                "개봉 전 사용기한(개월)": 30,
                "판정": "통과",
                "원본 행": 2,
            })
        result = APP.aggregate_location_exp_rows(rows, "DU", default_settings())
        self.assertEqual(
            [(item["EXP"], item["수량"], item["로트번호(생산)"]) for item in result],
            [
                (date(2028, 1, 1), 2, "L1"),
                (date(2028, 1, 1), 2, "L2"),
                (date(2028, 1, 1), 10, "L3"),
                (date(2028, 2, 1), 5, "L2"),
            ],
        )


if __name__ == "__main__":
    unittest.main()

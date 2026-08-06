from __future__ import annotations

import calendar
import hashlib
import io
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


APP_VERSION = "1.0.0"


class InputFormatError(ValueError):
    """입력 엑셀의 시트 또는 필수 열 구조가 예상과 다를 때 발생한다."""


@dataclass(frozen=True)
class ColumnNames:
    location: str = "저장 위치"
    material: str = "자재"
    product_name: str = "자재내역"
    available_stock: str = "가용재고(ASN)"
    expiry: str = "유통기한"
    lot_no: str = "로트번호(생산)"
    master_barcode: str = "바코드"
    master_months: str = "개봉 전 사용기한"


@dataclass(frozen=True)
class Settings:
    locations: tuple[str, ...]
    pass_threshold: float
    keep_missing_expiry: bool
    keep_missing_master: bool
    include_zero_in_details: bool
    expiry_ascending: bool
    quantity_ascending: bool
    lot_ascending: bool
    stock_sheet: str
    master_sheet: str
    columns: ColumnNames


def normalize_code(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def normalize_lot_number(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def find_required_columns(ws, required_headers: dict[str, str]) -> dict[str, int]:
    normalized_to_indexes: dict[str, list[int]] = defaultdict(list)
    for index, cell in enumerate(ws[1]):
        normalized_to_indexes[normalize_header(cell.value)].append(index)

    result: dict[str, int] = {}
    missing: list[str] = []
    duplicated: list[str] = []
    for key, expected_header in required_headers.items():
        matches = normalized_to_indexes.get(normalize_header(expected_header), [])
        if not matches:
            missing.append(expected_header)
        elif len(matches) > 1:
            duplicated.append(expected_header)
        else:
            result[key] = matches[0]

    if missing or duplicated:
        details = []
        if missing:
            details.append(f"누락 열: {', '.join(missing)}")
        if duplicated:
            details.append(f"중복 열: {', '.join(duplicated)}")
        raise InputFormatError(
            f"'{ws.title}' 시트의 필수 열 구조가 올바르지 않습니다. " + " / ".join(details)
        )
    return result


def parse_months(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        months = float(value)
        return months if months > 0 else None
    match = re.search(r"\d+(?:\.\d+)?", str(value))
    if not match:
        return None
    months = float(match.group())
    return months if months > 0 else None


def safe_number(value: Any) -> float | int | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        number = float(str(value).replace(",", "").strip())
        return int(number) if number.is_integer() else number
    except ValueError:
        return None


def to_date_only(value: Any) -> Optional[date]:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        from openpyxl.utils.datetime import from_excel

        try:
            converted = from_excel(value)
            return converted.date() if isinstance(converted, datetime) else converted
        except Exception:
            return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_filename_date(filename: str, fallback_year: int) -> Optional[date]:
    name = Path(filename).stem
    for token in re.findall(r"(?<!\d)(20\d{6})(?!\d)", name):
        try:
            return datetime.strptime(token, "%Y%m%d").date()
        except ValueError:
            pass
    for token in re.findall(r"(?<!\d)(\d{4})(?!\d)", name):
        month, day = int(token[:2]), int(token[2:])
        try:
            return date(fallback_year, month, day)
        except ValueError:
            pass
    return None


def auto_base_date(file_date: date, split_day: int) -> date:
    last_day = calendar.monthrange(file_date.year, file_date.month)[1]
    if file_date.day <= min(split_day, last_day):
        return date(file_date.year, file_date.month, min(split_day, last_day))
    return date(file_date.year, file_date.month, last_day)


def read_existing_base_date(pantos_bytes: bytes) -> Optional[date]:
    wb = load_workbook(io.BytesIO(pantos_bytes), data_only=True, read_only=True)
    try:
        for sheet_name in ("글로벌", "글로벌 유통기한", "글로벌유통기한"):
            if sheet_name in wb.sheetnames:
                candidate = to_date_only(wb[sheet_name]["G1"].value)
                if candidate:
                    return candidate
        return None
    finally:
        wb.close()


def resolve_base_date(
    pantos_bytes: bytes,
    filename: str,
    mode: str,
    manual_date: date,
    fallback_year: int,
    split_day: int,
    reuse_existing: bool,
) -> tuple[date, str, Optional[date]]:
    filename_date = parse_filename_date(filename, fallback_year)
    if mode == "직접 지정":
        return manual_date, "화면 직접 지정", filename_date

    if reuse_existing:
        existing = read_existing_base_date(pantos_bytes)
        if existing:
            return existing, "입력 파일 기존 글로벌 시트 G1", filename_date

    source_date = filename_date or date.today()
    source_text = "파일명 날짜 자동 계산" if filename_date else "파일명 날짜 없음 → 오늘 날짜 자동 계산"
    return auto_base_date(source_date, split_day), source_text, filename_date


def load_lifetime_months(master_bytes: bytes, settings: Settings) -> dict[str, Optional[float]]:
    wb = load_workbook(io.BytesIO(master_bytes), data_only=True, read_only=True)
    try:
        if settings.master_sheet not in wb.sheetnames:
            raise InputFormatError(
                f"Product Master에 '{settings.master_sheet}' 시트가 없습니다. "
                f"현재 시트: {', '.join(wb.sheetnames)}"
            )
        ws = wb[settings.master_sheet]
        cols = find_required_columns(
            ws,
            {
                "barcode": settings.columns.master_barcode,
                "months": settings.columns.master_months,
            },
        )
        mapping: dict[str, Optional[float]] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            barcode = normalize_code(row[cols["barcode"]] if len(row) > cols["barcode"] else None)
            months = parse_months(row[cols["months"]] if len(row) > cols["months"] else None)
            if not barcode:
                continue
            if barcode in mapping and mapping[barcode] != months:
                raise InputFormatError(
                    f"Product Master 바코드 {barcode}에 서로 다른 개봉 전 사용기한이 "
                    f"중복 등록되어 있습니다: {mapping[barcode]} / {months}"
                )
            mapping[barcode] = months
        return mapping
    finally:
        wb.close()


def build_source_rows(
    pantos_bytes: bytes,
    lifetime_map: dict[str, Optional[float]],
    base: date,
    settings: Settings,
) -> tuple[list[dict], list[dict]]:
    wb = load_workbook(io.BytesIO(pantos_bytes), data_only=True, read_only=True)
    try:
        if settings.stock_sheet not in wb.sheetnames:
            raise InputFormatError(
                f"판토스 파일에 '{settings.stock_sheet}' 시트가 없습니다. "
                f"현재 시트: {', '.join(wb.sheetnames)}"
            )
        ws = wb[settings.stock_sheet]
        c = settings.columns
        cols = find_required_columns(
            ws,
            {
                "location": c.location,
                "material": c.material,
                "product_name": c.product_name,
                "available_stock": c.available_stock,
                "expiry": c.expiry,
                "lot_no": c.lot_no,
            },
        )

        global_rows: list[dict] = []
        filtered_rows: list[dict] = []
        allowed = set(settings.locations)

        for row_number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            location = str(row[cols["location"]] or "").strip().upper()
            if location not in allowed:
                continue

            material = normalize_code(row[cols["material"]])
            product_name = row[cols["product_name"]]
            available_stock = safe_number(row[cols["available_stock"]])
            expiry = to_date_only(row[cols["expiry"]])
            lot_no = row[cols["lot_no"]]
            lifetime = lifetime_map.get(material)

            ratio = None
            if expiry and lifetime:
                ratio = (expiry - base).days / ((lifetime / 12) * 365)

            if expiry is None:
                status = "유지(유통기한 없음)" if settings.keep_missing_expiry else "제외(유통기한 없음)"
                keep = settings.keep_missing_expiry
            elif lifetime is None:
                status = "유지(마스터 누락)" if settings.keep_missing_master else "제외(마스터 누락)"
                keep = settings.keep_missing_master
            elif ratio is not None and ratio >= settings.pass_threshold:
                status, keep = "통과", True
            else:
                status, keep = "제외", False

            item = {
                "저장 위치": location,
                "자재": material,
                "자재내역": product_name,
                "가용재고(ASN)": available_stock,
                "유통기한": expiry,
                "로트번호(생산)": lot_no,
                "잔여 유통기한율": ratio,
                "개봉 전 사용기한(개월)": lifetime,
                "판정": status,
                "원본 행": row_number,
            }
            global_rows.append(item)
            if keep:
                filtered_rows.append(item)

        return global_rows, filtered_rows
    finally:
        wb.close()


def apply_standard_style(ws, max_col: int, freeze_cell: str = "A2") -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    body_font = Font(name="맑은 고딕", size=10)
    side = Side(style="thin", color="D9E2F3")
    border = Border(left=side, right=side, top=side, bottom=side)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=max_col):
        for cell in row:
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = freeze_cell


def style_summary_sheet(ws, max_col: int, header_row: int = 1, freeze_cell: str = "A2") -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    light_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    bold_font = Font(name="맑은 고딕", bold=True, size=10)
    body_font = Font(name="맑은 고딕", size=10)
    side = Side(style="thin", color="D9E2F3")
    border = Border(left=side, right=side, top=side, bottom=side)
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, max_col=max_col):
        for cell in row:
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    if ws.max_row > header_row and ws.cell(ws.max_row, 1).value == "총합계":
        for cell in ws[ws.max_row]:
            cell.font = bold_font
            cell.fill = light_fill
            cell.border = border
    ws.freeze_panes = freeze_cell


def add_table(ws, display_name: str, reference: str) -> None:
    table = Table(displayName=display_name, ref=reference)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


def write_data_sheet(
    wb: Workbook,
    title: str,
    rows: list[dict],
    settings: Settings,
    table_name: str,
) -> None:
    ws = wb.create_sheet(title)
    headers = [
        "저장 위치", "자재", "자재내역", "가용재고(ASN)", "유통기한", "로트번호(생산)",
        "잔여 유통기한율", "개봉 전 사용기한(개월)", "판정",
    ]
    ws.append(headers)
    for row_index, item in enumerate(rows, start=2):
        ws.append([
            item["저장 위치"], item["자재"], item["자재내역"], item["가용재고(ASN)"],
            item["유통기한"], item["로트번호(생산)"], None,
            item["개봉 전 사용기한(개월)"], None,
        ])
        ws.cell(row_index, 7).value = (
            f'=IF(OR(E{row_index}="",H{row_index}="",H{row_index}<=0),"",'
            f'(E{row_index}-\'기준_자동화\'!$B$4)/((H{row_index}/12)*365))'
        )
        missing_expiry_result = "유지(유통기한 없음)" if settings.keep_missing_expiry else "제외(유통기한 없음)"
        missing_master_result = "유지(마스터 누락)" if settings.keep_missing_master else "제외(마스터 누락)"
        ws.cell(row_index, 9).value = (
            f'=IF(E{row_index}="","{missing_expiry_result}",'
            f'IF(OR(H{row_index}="",H{row_index}<=0),"{missing_master_result}",'
            f'IF(G{row_index}>={settings.pass_threshold},"통과","제외")))'
        )

    apply_standard_style(ws, len(headers))
    for col, width in {
        "A": 12, "B": 18, "C": 45, "D": 15, "E": 13,
        "F": 18, "G": 16, "H": 20, "I": 18,
    }.items():
        ws.column_dimensions[col].width = width
    for cell in ws["D"]:
        cell.number_format = "#,##0"
    for cell in ws["E"]:
        cell.number_format = "yyyy-mm-dd"
    for cell in ws["G"]:
        cell.number_format = "0.0%"
    for cell in ws["H"]:
        cell.number_format = "0"

    if rows:
        add_table(ws, table_name, f"A1:I{len(rows) + 1}")
        ws.conditional_formatting.add(
            f"G2:G{len(rows) + 1}",
            CellIsRule(
                operator="lessThan",
                formula=[str(settings.pass_threshold)],
                fill=PatternFill("solid", fgColor="FCE4D6"),
            ),
        )
    else:
        ws.auto_filter.ref = "A1:I1"


def aggregate_by_product(rows: Iterable[dict]) -> dict[tuple[str, Any], dict]:
    grouped: dict[tuple[str, Any], dict] = {}
    for item in rows:
        key = (item["자재"], item["자재내역"])
        if key not in grouped:
            grouped[key] = {
                "자재": item["자재"],
                "자재내역": item["자재내역"],
                "qty": defaultdict(float),
            }
        grouped[key]["qty"][item["저장 위치"]] += item["가용재고(ASN)"] or 0
    return grouped


def product_sort_key(key: tuple[str, Any]):
    material = key[0]
    if str(material).isdigit():
        return 0, int(material), str(key[1])
    return 1, str(material), str(key[1])


def write_pivot_sheet(wb: Workbook, rows: list[dict], settings: Settings) -> None:
    ws = wb.create_sheet("글로벌 pv")
    ws["A1"] = "합계 : 가용재고(ASN)"
    ws["C1"] = "저장 위치"
    headers = ["자재", "자재내역", *settings.locations, "총합계"]
    ws.append(headers)
    grouped = aggregate_by_product(rows)
    start_row = 3
    for row_number, key in enumerate(sorted(grouped, key=product_sort_key), start=start_row):
        item = grouped[key]
        ws.cell(row_number, 1).value = item["자재"]
        ws.cell(row_number, 2).value = item["자재내역"]
        total = 0.0
        for column, location in enumerate(settings.locations, start=3):
            quantity = item["qty"].get(location, 0)
            ws.cell(row_number, column).value = int(quantity) if float(quantity).is_integer() else quantity
            total += quantity
        total_column = 3 + len(settings.locations)
        ws.cell(row_number, total_column).value = int(total) if float(total).is_integer() else total

    total_row = start_row + len(grouped)
    ws.cell(total_row, 1).value = "총합계"
    for column, location in enumerate(settings.locations, start=3):
        total = sum(item["qty"].get(location, 0) for item in grouped.values())
        ws.cell(total_row, column).value = int(total) if float(total).is_integer() else total
    grand_total = sum(sum(item["qty"].values()) for item in grouped.values())
    total_column = 3 + len(settings.locations)
    ws.cell(total_row, total_column).value = int(grand_total) if float(grand_total).is_integer() else grand_total

    light_fill = PatternFill("solid", fgColor="D9EAF7")
    bold_font = Font(name="맑은 고딕", bold=True, size=10)
    side = Side(style="thin", color="D9E2F3")
    border = Border(left=side, right=side, top=side, bottom=side)
    for cell in ws[1]:
        cell.font, cell.fill, cell.border = bold_font, light_fill, border

    style_summary_sheet(ws, len(headers), header_row=2, freeze_cell="A3")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 45
    for column in range(3, len(headers) + 1):
        letter = ws.cell(1, column).column_letter
        ws.column_dimensions[letter].width = 14
        for cell in ws[letter]:
            cell.number_format = "#,##0"
    if grouped:
        add_table(ws, "GlobalExpiryPVTable", f"A2:{ws.cell(2, len(headers)).column_letter}{total_row - 1}")
    else:
        ws.auto_filter.ref = f"A2:{ws.cell(2, len(headers)).column_letter}2"


def aggregate_location_exp_rows(rows: list[dict], location: str, settings: Settings) -> list[dict]:
    grouped: dict[tuple, dict] = {}
    for item in rows:
        if item["저장 위치"] != location:
            continue
        quantity = item["가용재고(ASN)"] or 0
        if quantity == 0 and not settings.include_zero_in_details:
            continue
        lot_no = normalize_lot_number(item["로트번호(생산)"])
        key = (
            item["자재"], item["자재내역"], item["유통기한"], lot_no,
            item["잔여 유통기한율"], item["개봉 전 사용기한(개월)"],
        )
        if key not in grouped:
            grouped[key] = {
                "자재": item["자재"],
                "자재내역": item["자재내역"],
                "수량": 0,
                "EXP": item["유통기한"],
                "잔여 유통기한율": item["잔여 유통기한율"],
                "개봉 전 사용기한(개월)": item["개봉 전 사용기한(개월)"],
                "로트번호(생산)": lot_no,
            }
        grouped[key]["수량"] += quantity

    result = list(grouped.values())

    def base_key(item: dict):
        material = item["자재"]
        material_key = (0, int(material)) if str(material).isdigit() else (1, str(material))
        return material_key, str(item["자재내역"] or "")

    # 안정 정렬을 역순으로 적용해 각 기준의 오름/내림차순을 독립적으로 제어한다.
    result.sort(key=lambda x: x["로트번호(생산)"], reverse=not settings.lot_ascending)
    result.sort(key=lambda x: x["수량"], reverse=not settings.quantity_ascending)
    result.sort(
        key=lambda x: (
            x["EXP"] is None,
            (x["EXP"] or date.min).toordinal() * (1 if settings.expiry_ascending else -1),
        ),
    )
    result.sort(key=base_key)
    return result


def write_location_sheets(wb: Workbook, rows: list[dict], settings: Settings) -> None:
    headers = [
        "자재", "자재내역", "수량", "EXP", "잔여 유통기한율",
        "개봉 전 사용기한(개월)", "로트번호(생산)",
    ]
    for location in settings.locations:
        ws = wb.create_sheet(f"글로벌 유통기한 {location}")
        ws.append(headers)
        detail_rows = aggregate_location_exp_rows(rows, location, settings)
        for item in detail_rows:
            quantity = item["수량"]
            ws.append([
                item["자재"], item["자재내역"],
                int(quantity) if float(quantity).is_integer() else quantity,
                item["EXP"], item["잔여 유통기한율"],
                item["개봉 전 사용기한(개월)"], item["로트번호(생산)"],
            ])
        total_row = ws.max_row + 1
        ws.cell(total_row, 1).value = "총합계"
        ws.cell(total_row, 3).value = sum((ws.cell(r, 3).value or 0) for r in range(2, total_row))
        style_summary_sheet(ws, len(headers))
        for col, width in {
            "A": 18, "B": 45, "C": 14, "D": 13, "E": 16, "F": 20, "G": 28,
        }.items():
            ws.column_dimensions[col].width = width
        for cell in ws["C"]:
            cell.number_format = "#,##0"
        for cell in ws["D"]:
            cell.number_format = "yyyy-mm-dd"
        for cell in ws["E"]:
            cell.number_format = "0.0%"
        for cell in ws["F"]:
            cell.number_format = "0"
        if detail_rows:
            safe_location = re.sub(r"[^A-Za-z0-9_]", "", location) or "LOC"
            add_table(ws, f"{safe_location}ExpTable", f"A1:G{total_row - 1}")
            ws.conditional_formatting.add(
                f"E2:E{total_row - 1}",
                CellIsRule(
                    operator="lessThan",
                    formula=[str(settings.pass_threshold)],
                    fill=PatternFill("solid", fgColor="FCE4D6"),
                ),
            )
        else:
            ws.auto_filter.ref = "A1:G1"


def write_meta_sheet(
    wb: Workbook,
    pantos_name: str,
    master_name: str,
    base: date,
    base_source: str,
    filename_date: Optional[date],
    global_rows: list[dict],
    filtered_rows: list[dict],
    settings: Settings,
) -> None:
    ws = wb.create_sheet("기준_자동화")
    counts = defaultdict(int)
    pass_counts = defaultdict(int)
    for item in global_rows:
        counts[item["저장 위치"]] += 1
    for item in filtered_rows:
        pass_counts[item["저장 위치"]] += 1
    direction = lambda ascending: "빠른/작은 순" if ascending else "늦은/큰 순"
    meta = [
        ["항목", "값"],
        ["판토스 파일", pantos_name],
        ["Product Master 파일", master_name],
        ["기준일", base],
        ["기준일 출처", base_source],
        ["파일명 기준 날짜", filename_date],
        ["대상 창고 및 출력 순서", " → ".join(settings.locations)],
        ["유통기한 통과 기준", f"잔여 유통기한율 >= {settings.pass_threshold:.1%}"],
        ["유통기한 없음 처리", "유지" if settings.keep_missing_expiry else "제외"],
        ["Product Master 사용기한 누락 처리", "유지" if settings.keep_missing_master else "제외"],
        ["창고별 상세 정렬", (
            f"자재별 → EXP {direction(settings.expiry_ascending)} → "
            f"수량 {direction(settings.quantity_ascending)} → "
            f"LOT {'오름차순' if settings.lot_ascending else '내림차순'}"
        )],
        ["글로벌 행 수", len(global_rows)],
        ["글로벌 유통기한 행 수", len(filtered_rows)],
        ["원본 창고별 행 수", " / ".join(f"{loc} {counts[loc]:,}" for loc in settings.locations)],
        ["유지 창고별 행 수", " / ".join(f"{loc} {pass_counts[loc]:,}" for loc in settings.locations)],
        ["주의: 유통기한 없음", sum(item["유통기한"] is None for item in global_rows)],
        ["주의: Product Master 사용기한 누락", sum(
            item["유통기한"] is not None and item["개봉 전 사용기한(개월)"] is None
            for item in global_rows
        )],
        ["주의: 로트번호 없음", sum(
            normalize_lot_number(item["로트번호(생산)"]) == "" for item in global_rows
        )],
        ["참고: 가용재고 0", sum((item["가용재고(ASN)"] or 0) == 0 for item in global_rows)],
        ["생성 앱 버전", APP_VERSION],
    ]
    for row in meta:
        ws.append(row)
    apply_standard_style(ws, 2)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 90
    ws["B4"].number_format = "yyyy-mm-dd"
    ws["B6"].number_format = "yyyy-mm-dd"
    for row in ws.iter_rows(min_row=2, max_col=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def create_inventory_workbook(
    pantos_bytes: bytes,
    master_bytes: bytes,
    pantos_name: str,
    master_name: str,
    base: date,
    base_source: str,
    filename_date: Optional[date],
    settings: Settings,
) -> tuple[bytes, list[dict], list[dict]]:
    lifetime_map = load_lifetime_months(master_bytes, settings)
    global_rows, filtered_rows = build_source_rows(pantos_bytes, lifetime_map, base, settings)

    wb = Workbook()
    wb.remove(wb.active)
    write_data_sheet(wb, "글로벌", global_rows, settings, "GlobalInventoryTable")
    write_data_sheet(wb, "글로벌 유통기한", filtered_rows, settings, "GlobalExpiryDataTable")
    write_pivot_sheet(wb, filtered_rows, settings)
    write_location_sheets(wb, filtered_rows, settings)
    write_meta_sheet(
        wb, pantos_name, master_name, base, base_source, filename_date,
        global_rows, filtered_rows, settings,
    )
    order = [
        "글로벌", "글로벌 유통기한", "글로벌 pv",
        *(f"글로벌 유통기한 {location}" for location in settings.locations),
        "기준_자동화",
    ]
    wb._sheets = [wb[name] for name in order]
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    output = io.BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue(), global_rows, filtered_rows


def parse_locations(text: str) -> tuple[str, ...]:
    locations: list[str] = []
    for token in re.split(r"[,>\s]+", text.upper().strip()):
        if token and token not in locations:
            locations.append(token)
    if not locations:
        raise ValueError("대상 창고를 하나 이상 입력해 주세요.")
    if len(locations) > 10:
        raise ValueError("대상 창고는 최대 10개까지 설정할 수 있습니다.")
    return tuple(locations)


def sanitize_output_name(value: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', "_", value.strip())
    if not name:
        name = "글로벌재고_자동생성.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def result_signature(
    pantos_bytes: bytes,
    master_bytes: bytes,
    pantos_name: str,
    master_name: str,
    base: date,
    settings: Settings,
) -> str:
    digest = hashlib.sha256()
    digest.update(pantos_bytes)
    digest.update(master_bytes)
    digest.update(pantos_name.encode("utf-8"))
    digest.update(master_name.encode("utf-8"))
    digest.update(base.isoformat().encode("utf-8"))
    digest.update(repr(settings).encode("utf-8"))
    return digest.hexdigest()


def render_sidebar() -> dict[str, Any]:
    with st.sidebar:
        st.header("생성 기준")
        base_mode = st.radio("기준일", ["자동", "직접 지정"], horizontal=True)
        manual_date = st.date_input("직접 기준일", value=date.today(), disabled=base_mode != "직접 지정")
        threshold_percent = st.number_input(
            "잔여 유통기한 통과율 (%)", min_value=0.0, max_value=100.0,
            value=66.0, step=1.0,
        )
        location_text = st.text_input("대상 창고 및 순서", value="DU, GD, GL")

        st.subheader("예외 처리")
        keep_missing_expiry = st.checkbox("유통기한 없는 재고 유지", value=True)
        keep_missing_master = st.checkbox("마스터 사용기한 누락 재고 유지", value=True)
        include_zero = st.checkbox("창고별 상세에 0재고 표시", value=False)

        st.subheader("창고별 상세 정렬")
        expiry_order = st.selectbox("EXP", ["빠른 순", "늦은 순"])
        quantity_order = st.selectbox("동일 EXP 내 수량", ["적은 순", "많은 순"])
        lot_order = st.selectbox("동일 조건 내 LOT", ["오름차순", "내림차순"])

        with st.expander("자동 기준일 상세"):
            fallback_year = st.number_input(
                "파일명이 MMDD일 때 연도", min_value=2000, max_value=2100,
                value=date.today().year, step=1,
            )
            split_day = st.number_input(
                "전반기 기준일", min_value=1, max_value=28, value=15, step=1,
                help="파일 날짜가 이 날짜 이하면 해당 날짜, 이후면 월말을 기준일로 사용합니다.",
            )
            reuse_existing = st.checkbox("기존 글로벌 시트 G1 기준일 우선 사용", value=True)

        with st.expander("입력 시트·열 이름"):
            stock_sheet = st.text_input("판토스 원본 시트", value="Sheet1")
            master_sheet = st.text_input("Product Master 시트", value="유효기간")
            location_col = st.text_input("창고 열", value="저장 위치")
            material_col = st.text_input("자재 열", value="자재")
            product_col = st.text_input("제품명 열", value="자재내역")
            stock_col = st.text_input("가용재고 열", value="가용재고(ASN)")
            expiry_col = st.text_input("유통기한 열", value="유통기한")
            lot_col = st.text_input("LOT 열", value="로트번호(생산)")
            barcode_col = st.text_input("마스터 바코드 열", value="바코드")
            months_col = st.text_input("마스터 사용기한 열", value="개봉 전 사용기한")

    return {
        "base_mode": base_mode,
        "manual_date": manual_date,
        "threshold": threshold_percent / 100,
        "location_text": location_text,
        "keep_missing_expiry": keep_missing_expiry,
        "keep_missing_master": keep_missing_master,
        "include_zero": include_zero,
        "expiry_ascending": expiry_order == "빠른 순",
        "quantity_ascending": quantity_order == "적은 순",
        "lot_ascending": lot_order == "오름차순",
        "fallback_year": int(fallback_year),
        "split_day": int(split_day),
        "reuse_existing": reuse_existing,
        "stock_sheet": stock_sheet.strip(),
        "master_sheet": master_sheet.strip(),
        "columns": ColumnNames(
            location=location_col, material=material_col, product_name=product_col,
            available_stock=stock_col, expiry=expiry_col, lot_no=lot_col,
            master_barcode=barcode_col, master_months=months_col,
        ),
    }


def main() -> None:
    st.set_page_config(page_title="글로벌 재고 자동화", page_icon="📦", layout="wide")
    controls = render_sidebar()

    st.title("글로벌 재고 파일 만들기")
    left, right = st.columns(2)
    with left:
        pantos_file = st.file_uploader("판토스 재고 파일", type=["xlsx"], key="pantos")
    with right:
        master_file = st.file_uploader("Product Master", type=["xlsx"], key="master")

    output_name = st.text_input("결과 파일명", value="글로벌재고_자동생성.xlsx")

    if not pantos_file or not master_file:
        st.info("두 엑셀 파일을 업로드하면 재고 파일을 생성할 수 있습니다.")
        return

    pantos_bytes = pantos_file.getvalue()
    master_bytes = master_file.getvalue()

    try:
        locations = parse_locations(controls["location_text"])
        base, base_source, filename_date = resolve_base_date(
            pantos_bytes=pantos_bytes,
            filename=pantos_file.name,
            mode=controls["base_mode"],
            manual_date=controls["manual_date"],
            fallback_year=controls["fallback_year"],
            split_day=controls["split_day"],
            reuse_existing=controls["reuse_existing"],
        )
        settings = Settings(
            locations=locations,
            pass_threshold=controls["threshold"],
            keep_missing_expiry=controls["keep_missing_expiry"],
            keep_missing_master=controls["keep_missing_master"],
            include_zero_in_details=controls["include_zero"],
            expiry_ascending=controls["expiry_ascending"],
            quantity_ascending=controls["quantity_ascending"],
            lot_ascending=controls["lot_ascending"],
            stock_sheet=controls["stock_sheet"],
            master_sheet=controls["master_sheet"],
            columns=controls["columns"],
        )
        current_signature = result_signature(
            pantos_bytes, master_bytes, pantos_file.name, master_file.name, base, settings
        )
    except (ValueError, InputFormatError) as exc:
        st.error(str(exc))
        return

    a, b, c = st.columns(3)
    a.metric("기준일", base.strftime("%Y-%m-%d"))
    b.metric("통과율", f"{settings.pass_threshold:.0%}")
    c.metric("창고 순서", " → ".join(settings.locations))
    st.caption(f"기준일 출처: {base_source}")

    if st.button("재고 파일 만들기", type="primary", use_container_width=True):
        try:
            with st.spinner("재고 데이터를 처리하고 있습니다..."):
                output_bytes, global_rows, filtered_rows = create_inventory_workbook(
                    pantos_bytes=pantos_bytes,
                    master_bytes=master_bytes,
                    pantos_name=pantos_file.name,
                    master_name=master_file.name,
                    base=base,
                    base_source=base_source,
                    filename_date=filename_date,
                    settings=settings,
                )
            st.session_state["inventory_result"] = {
                "bytes": output_bytes,
                "global_count": len(global_rows),
                "filtered_count": len(filtered_rows),
                "base": base,
                "preview": filtered_rows[:100],
                "signature": current_signature,
            }
        except (InputFormatError, ValueError, KeyError, TypeError) as exc:
            st.error(f"입력 파일을 처리할 수 없습니다: {exc}")
            return
        except Exception as exc:
            st.error(f"예상하지 못한 오류가 발생했습니다: {type(exc).__name__}: {exc}")
            return

    result = st.session_state.get("inventory_result")
    if result and result.get("signature") == current_signature:
        st.success(
            f"완료: 전체 {result['global_count']:,}행 중 "
            f"{result['filtered_count']:,}행이 유통기한 결과에 반영되었습니다."
        )
        st.download_button(
            "결과 엑셀 다운로드",
            data=result["bytes"],
            file_name=sanitize_output_name(output_name),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        with st.expander("결과 미리보기 (최대 100행)"):
            preview = []
            for item in result["preview"]:
                preview.append({
                    "창고": item["저장 위치"],
                    "자재": item["자재"],
                    "제품명": item["자재내역"],
                    "수량": item["가용재고(ASN)"],
                    "EXP": item["유통기한"],
                    "LOT": item["로트번호(생산)"],
                    "잔여율": item["잔여 유통기한율"],
                    "판정": item["판정"],
                })
            st.dataframe(preview, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()

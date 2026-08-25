from __future__ import annotations

import calendar
import hashlib
import io
import math
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional
from zoneinfo import ZoneInfo
from zipfile import BadZipFile, ZipFile

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


APP_VERSION = "1.4.1"
MAX_UPLOAD_BYTES = 50 * 1024 * 1024
MAX_SOURCE_ROWS = 200_000
MAX_VALIDATION_MESSAGES = 100
STOCK_HEADER_ROW = 1
MASTER_HEADER_ROW = 12
KST = ZoneInfo("Asia/Seoul")
MISSING_MARKERS = {"-", "N/A", "NA", "NONE", "NULL", "UNKNOWN", "미확인", "미정"}


class InputFormatError(ValueError):
    """입력 엑셀의 시트 또는 필수 열 구조가 예상과 다를 때 발생한다."""


class DataValidationError(ValueError):
    """입력 데이터가 업무 규칙상 안전하게 처리될 수 없을 때 발생한다."""


@dataclass
class ValidationSummary:
    source_rows: int = 0
    selected_rows: int = 0
    kept_rows: int = 0
    source_quantity: float = 0
    kept_quantity: float = 0
    missing_expiry_rows: int = 0
    missing_master_rows: int = 0
    missing_lot_rows: int = 0
    missing_product_name_rows: int = 0
    zero_quantity_rows: int = 0
    master_rows: int = 0
    master_blank_barcode_rows: int = 0
    master_blank_months_rows: int = 0
    master_match_rows: int = 0
    product_name_conflicts: int = 0
    warnings: list[str] = field(default_factory=list)
    location_rows: dict[str, int] = field(default_factory=dict)
    kept_location_rows: dict[str, int] = field(default_factory=dict)

    @property
    def master_match_ratio(self) -> float:
        return self.master_match_rows / self.selected_rows if self.selected_rows else 0.0


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


def today_kst() -> date:
    return datetime.now(KST).date()


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def is_missing_marker(value: Any) -> bool:
    return isinstance(value, str) and value.strip().upper() in MISSING_MARKERS


def is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def row_value(row: tuple[Any, ...], index: int) -> Any:
    return row[index] if index < len(row) else None


def write_safe_text(cell, value: Any) -> None:
    """업로드 데이터의 문자열이 엑셀 수식으로 재해석되지 않도록 텍스트로 기록한다."""
    if value is None:
        cell.value = None
    else:
        cell.value = str(value)
        cell.data_type = "s"


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


def make_batch_number(expiry: Optional[date], lot_no: Any) -> Optional[str]:
    """유통기한과 LOT가 모두 있을 때 YYYYMMDD|LOT 형식의 배치번호를 만든다."""
    normalized_lot = normalize_lot_number(lot_no)
    if expiry is None or not normalized_lot:
        return None
    return f"{expiry:%Y%m%d}|{normalized_lot}"


def find_required_columns(
    ws,
    required_headers: dict[str, str],
    header_row: int = 1,
) -> dict[str, int]:
    normalized_to_indexes: dict[str, list[int]] = defaultdict(list)
    header_cells = next(
        ws.iter_rows(min_row=header_row, max_row=header_row),
        (),
    )
    for index, cell in enumerate(header_cells):
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
            f"'{ws.title}' 시트 {header_row}행의 필수 열 구조가 올바르지 않습니다. "
            + " / ".join(details)
        )
    return result


def parse_months(value: Any) -> Optional[float]:
    if is_blank(value) or is_missing_marker(value):
        return None
    if isinstance(value, bool) or is_formula(value):
        raise ValueError("숫자 값이 아닙니다")
    if isinstance(value, (int, float)):
        months = float(value)
    else:
        text = str(value).strip().replace(",", "")
        year_match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*년", text)
        month_match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(?:개월|개월간|month(?:s)?)?", text, re.I)
        if year_match:
            months = float(year_match.group(1)) * 12
        elif month_match:
            months = float(month_match.group(1))
        else:
            raise ValueError("개월 수 형식이 아닙니다")
    if not math.isfinite(months) or months <= 0 or months > 120:
        raise ValueError("사용기한은 0 초과 120개월 이하여야 합니다")
    return months


def safe_number(value: Any) -> float | int | None:
    if is_blank(value):
        return None
    if isinstance(value, bool) or is_formula(value):
        raise ValueError("숫자 값이 아닙니다")
    try:
        number = float(value) if isinstance(value, (int, float)) else float(str(value).replace(",", "").strip())
    except (TypeError, ValueError) as exc:
        raise ValueError("숫자 형식이 아닙니다") from exc
    if not math.isfinite(number):
        raise ValueError("유한한 숫자가 아닙니다")
    if number < 0:
        raise ValueError("음수 재고는 허용되지 않습니다")
    if not number.is_integer():
        raise ValueError("재고수량은 정수여야 합니다")
    return int(number)


def to_date_only(value: Any) -> Optional[date]:
    if is_blank(value) or is_missing_marker(value):
        return None
    if isinstance(value, bool) or is_formula(value):
        raise ValueError("날짜 값이 아닙니다")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        from openpyxl.utils.datetime import from_excel

        number = float(value)
        if not math.isfinite(number):
            raise ValueError("유한한 날짜 값이 아닙니다")
        if number.is_integer() and 19000101 <= int(number) <= 21001231:
            try:
                return datetime.strptime(str(int(number)), "%Y%m%d").date()
            except ValueError as exc:
                raise ValueError("YYYYMMDD 날짜가 올바르지 않습니다") from exc
        try:
            converted = from_excel(value)
            return converted.date() if isinstance(converted, datetime) else converted
        except Exception as exc:
            raise ValueError("엑셀 날짜 일련번호가 올바르지 않습니다") from exc
    text = str(value).strip()
    korean = re.fullmatch(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일?", text)
    if korean:
        try:
            return date(int(korean.group(1)), int(korean.group(2)), int(korean.group(3)))
        except ValueError as exc:
            raise ValueError("한글 날짜가 올바르지 않습니다") from exc
    normalized = text.replace("T", " ").removesuffix("Z")
    for fmt in (
        "%Y%m%d", "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
        "%m/%d/%Y", "%m-%d-%Y",
    ):
        try:
            return datetime.strptime(normalized, fmt).date()
        except ValueError:
            continue
    raise ValueError("지원하지 않는 날짜 형식입니다")


def add_months(start: date, months: int) -> date:
    year = start.year + (start.month - 1 + months) // 12
    month = (start.month - 1 + months) % 12 + 1
    day = min(start.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def calculate_remaining_ratio(expiry: date, base: date, lifetime_months: float) -> float:
    integer_months = int(lifetime_months)
    if float(lifetime_months).is_integer():
        denominator = (add_months(base, integer_months) - base).days
    else:
        denominator = round(lifetime_months * 365.2425 / 12)
    if denominator <= 0:
        raise ValueError("사용기한으로 계산한 일수가 0 이하입니다")
    return (expiry - base).days / denominator


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


@st.cache_data(show_spinner=False, max_entries=8)
def read_existing_base_date(pantos_bytes: bytes) -> Optional[date]:
    wb = load_workbook(io.BytesIO(pantos_bytes), data_only=True, read_only=True)
    try:
        for sheet_name in ("글로벌", "글로벌 유통기한", "글로벌유통기한"):
            if sheet_name in wb.sheetnames:
                try:
                    candidate = to_date_only(wb[sheet_name]["G1"].value)
                except ValueError:
                    candidate = None
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

    source_date = filename_date or today_kst()
    source_text = "파일명 날짜 자동 계산" if filename_date else "파일명 날짜 없음 → 오늘 날짜 자동 계산"
    return auto_base_date(source_date, split_day), source_text, filename_date


def load_lifetime_months(
    master_bytes: bytes,
    settings: Settings,
    summary: ValidationSummary,
) -> dict[str, Optional[float]]:
    wb = load_workbook(io.BytesIO(master_bytes), data_only=False, read_only=True)
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
            header_row=MASTER_HEADER_ROW,
        )
        mapping: dict[str, Optional[float]] = {}
        errors: list[str] = []
        master_data_row = MASTER_HEADER_ROW + 1
        for row_number, row in enumerate(
            ws.iter_rows(min_row=master_data_row, values_only=True),
            start=master_data_row,
        ):
            if row_number - MASTER_HEADER_ROW > MAX_SOURCE_ROWS:
                raise InputFormatError(f"Product Master는 최대 {MAX_SOURCE_ROWS:,}행까지 처리할 수 있습니다.")
            summary.master_rows += 1
            raw_barcode = row_value(row, cols["barcode"])
            raw_months = row_value(row, cols["months"])
            if is_formula(raw_barcode):
                errors.append(f"마스터 {row_number}행: 바코드가 수식입니다. 값으로 붙여넣어 주세요.")
                continue
            barcode = normalize_code(raw_barcode)
            if not barcode:
                summary.master_blank_barcode_rows += 1
                continue
            if isinstance(raw_barcode, (int, float)) and not isinstance(raw_barcode, bool) and abs(raw_barcode) >= 10**15:
                errors.append(f"마스터 {row_number}행: 15자리 초과 숫자 바코드는 엑셀 정밀도 손실 위험이 있습니다. 텍스트로 저장해 주세요.")
                continue
            try:
                months = parse_months(raw_months)
            except ValueError as exc:
                errors.append(f"마스터 {row_number}행 바코드 {barcode}: 개봉 전 사용기한 오류 ({exc})")
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue
            if months is None:
                summary.master_blank_months_rows += 1
            if barcode in mapping and mapping[barcode] != months:
                errors.append(
                    f"마스터 바코드 {barcode}: 서로 다른 사용기한 중복 ({mapping[barcode]} / {months})"
                )
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue
            mapping[barcode] = months
        if errors:
            raise DataValidationError("Product Master 데이터 오류:\n- " + "\n- ".join(errors))
        if not mapping:
            raise DataValidationError("Product Master에서 사용할 수 있는 바코드를 찾지 못했습니다.")
        return mapping
    finally:
        wb.close()


def build_source_rows(
    pantos_bytes: bytes,
    lifetime_map: dict[str, Optional[float]],
    base: date,
    settings: Settings,
    summary: ValidationSummary,
) -> tuple[list[dict], list[dict]]:
    wb = load_workbook(io.BytesIO(pantos_bytes), data_only=False, read_only=True)
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
            header_row=STOCK_HEADER_ROW,
        )

        global_rows: list[dict] = []
        filtered_rows: list[dict] = []
        allowed = set(settings.locations)
        seen_locations: set[str] = set()
        errors: list[str] = []
        names_by_material: dict[str, set[str]] = defaultdict(set)

        stock_data_row = STOCK_HEADER_ROW + 1
        for row_number, row in enumerate(
            ws.iter_rows(min_row=stock_data_row, values_only=True),
            start=stock_data_row,
        ):
            if row_number - STOCK_HEADER_ROW > MAX_SOURCE_ROWS:
                raise InputFormatError(f"판토스 원본은 최대 {MAX_SOURCE_ROWS:,}행까지 처리할 수 있습니다.")
            summary.source_rows += 1
            raw_location = row_value(row, cols["location"])
            if is_formula(raw_location):
                errors.append(f"판토스 {row_number}행: 창고 코드가 수식입니다. 값으로 붙여넣어 주세요.")
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue
            location = str(raw_location or "").strip().upper()
            if location:
                seen_locations.add(location)
            if location not in allowed:
                continue

            summary.selected_rows += 1
            raw_material = row_value(row, cols["material"])
            raw_product_name = row_value(row, cols["product_name"])
            raw_stock = row_value(row, cols["available_stock"])
            raw_expiry = row_value(row, cols["expiry"])
            raw_lot = row_value(row, cols["lot_no"])

            if is_formula(raw_material):
                errors.append(f"판토스 {row_number}행: 자재코드가 수식입니다. 값으로 붙여넣어 주세요.")
                continue
            material = normalize_code(raw_material)
            if not material:
                errors.append(f"판토스 {row_number}행: 자재코드가 비어 있습니다.")
                continue
            if isinstance(raw_material, (int, float)) and not isinstance(raw_material, bool) and abs(raw_material) >= 10**15:
                errors.append(f"판토스 {row_number}행: 15자리 초과 숫자 자재코드는 텍스트로 저장해 주세요.")
                continue

            if is_formula(raw_product_name):
                errors.append(f"판토스 {row_number}행 자재 {material}: 제품명이 수식입니다. 값으로 붙여넣어 주세요.")
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue
            if is_formula(raw_lot):
                errors.append(f"판토스 {row_number}행 자재 {material}: LOT가 수식입니다. 값으로 붙여넣어 주세요.")
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue

            try:
                available_stock = safe_number(raw_stock)
            except ValueError as exc:
                errors.append(f"판토스 {row_number}행 자재 {material}: 가용재고 오류 ({exc})")
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue
            if available_stock is None:
                errors.append(f"판토스 {row_number}행 자재 {material}: 가용재고가 비어 있습니다.")
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue

            try:
                expiry = to_date_only(raw_expiry)
            except ValueError as exc:
                errors.append(f"판토스 {row_number}행 자재 {material}: 유통기한 오류 ({raw_expiry!r}, {exc})")
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue

            product_name = None if is_blank(raw_product_name) else str(raw_product_name).strip()
            lot_no = None if is_blank(raw_lot) or is_missing_marker(raw_lot) else normalize_lot_number(raw_lot)
            if lot_no and re.search(r"[,;\n\r]", lot_no):
                errors.append(
                    f"판토스 {row_number}행 자재 {material}: 복수 LOT로 보이는 값({lot_no!r})은 "
                    "LOT별 수량을 확인할 수 없어 처리할 수 없습니다. LOT별 행으로 분리해 주세요."
                )
                if len(errors) >= MAX_VALIDATION_MESSAGES:
                    break
                continue

            lifetime = lifetime_map.get(material)

            ratio = None
            if expiry and lifetime:
                ratio = calculate_remaining_ratio(expiry, base, lifetime)

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
                "배치번호": make_batch_number(expiry, lot_no),
                "잔여 유통기한율": ratio,
                "개봉 전 사용기한(개월)": lifetime,
                "판정": status,
                "원본 행": row_number,
            }
            global_rows.append(item)
            summary.location_rows[location] = summary.location_rows.get(location, 0) + 1
            summary.source_quantity += available_stock
            if expiry is None:
                summary.missing_expiry_rows += 1
            if lifetime is None:
                summary.missing_master_rows += 1
            else:
                summary.master_match_rows += 1
            if not lot_no:
                summary.missing_lot_rows += 1
            if not product_name:
                summary.missing_product_name_rows += 1
            else:
                names_by_material[material].add(product_name)
            if available_stock == 0:
                summary.zero_quantity_rows += 1
            if keep:
                filtered_rows.append(item)
                summary.kept_location_rows[location] = summary.kept_location_rows.get(location, 0) + 1
                summary.kept_quantity += available_stock

        if errors:
            suffix = "\n- 오류가 더 있을 수 있습니다." if len(errors) >= MAX_VALIDATION_MESSAGES else ""
            raise DataValidationError("판토스 데이터 오류:\n- " + "\n- ".join(errors) + suffix)
        if not global_rows:
            raise DataValidationError(
                "선택한 창고에 처리할 재고 행이 없습니다. 원본 창고 코드와 대상 창고 입력을 확인해 주세요."
            )
        if summary.master_match_rows == 0 and any(item["유통기한"] is not None for item in global_rows):
            raise DataValidationError(
                "유통기한이 입력된 재고가 있지만 Product Master와 매칭된 행이 0건입니다. "
                "자재·바코드 형식과 파일 조합을 확인해 주세요."
            )

        summary.kept_rows = len(filtered_rows)
        requested_missing = [location for location in settings.locations if location not in seen_locations]
        if requested_missing:
            summary.warnings.append("원본에 없는 선택 창고: " + ", ".join(requested_missing))
        summary.product_name_conflicts = sum(len(names) > 1 for names in names_by_material.values())
        if summary.product_name_conflicts:
            examples = [
                f"{material} ({' / '.join(sorted(names)[:3])})"
                for material, names in names_by_material.items() if len(names) > 1
            ][:10]
            summary.warnings.append(
                f"동일 자재코드의 제품명 불일치 {summary.product_name_conflicts:,}건: " + "; ".join(examples)
            )
        if summary.missing_expiry_rows:
            summary.warnings.append(
                f"유통기한 공란 {summary.missing_expiry_rows:,}행은 설정에 따라 "
                f"{'유지' if settings.keep_missing_expiry else '제외'}되었습니다."
            )
        if summary.missing_master_rows:
            summary.warnings.append(
                f"마스터 사용기한 미매칭/공란 {summary.missing_master_rows:,}행은 설정에 따라 "
                f"{'유지' if settings.keep_missing_master else '제외'}되었습니다."
            )
        if summary.master_match_ratio < 0.5:
            summary.warnings.append(f"마스터 매칭률이 {summary.master_match_ratio:.1%}로 낮습니다. 파일 조합을 확인해 주세요.")

        return global_rows, filtered_rows
    finally:
        wb.close()


def ensure_named_style(
    wb: Workbook,
    name: str,
    *,
    font: Font,
    border: Border,
    alignment: Alignment,
    fill: Optional[PatternFill] = None,
    number_format: str = "General",
) -> str:
    if name not in wb.named_styles:
        wb.add_named_style(NamedStyle(
            name=name,
            font=font,
            border=border,
            alignment=alignment,
            fill=fill or PatternFill(),
            number_format=number_format,
        ))
    return name


def format_style_name(prefix: str, column: int, number_format: str) -> str:
    suffix = hashlib.sha1(number_format.encode("utf-8")).hexdigest()[:6]
    return f"{prefix}_{column}_{suffix}"


def apply_standard_style(
    ws,
    max_col: int,
    freeze_cell: str = "A2",
    number_formats: Optional[dict[int, str]] = None,
) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    body_font = Font(name="맑은 고딕", size=10)
    side = Side(style="thin", color="D9E2F3")
    border = Border(left=side, right=side, top=side, bottom=side)
    header_style = ensure_named_style(
        ws.parent,
        "InvStandardHeader",
        font=header_font,
        border=border,
        alignment=Alignment(horizontal="center", vertical="center"),
        fill=header_fill,
    )
    body_style = ensure_named_style(
        ws.parent,
        "InvStandardBody",
        font=body_font,
        border=border,
        alignment=Alignment(vertical="center"),
    )
    formatted_styles = {
        column: ensure_named_style(
            ws.parent,
            format_style_name("InvStandardBody", column, number_format),
            font=body_font,
            border=border,
            alignment=Alignment(vertical="center"),
            number_format=number_format,
        )
        for column, number_format in (number_formats or {}).items()
    }
    for cell in ws[1]:
        cell.style = header_style
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=max_col):
        for cell in row:
            cell.style = formatted_styles.get(cell.column, body_style)
    ws.freeze_panes = freeze_cell


def style_summary_sheet(
    ws,
    max_col: int,
    header_row: int = 1,
    freeze_cell: str = "A2",
    number_formats: Optional[dict[int, str]] = None,
) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    light_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    bold_font = Font(name="맑은 고딕", bold=True, size=10)
    body_font = Font(name="맑은 고딕", size=10)
    side = Side(style="thin", color="D9E2F3")
    border = Border(left=side, right=side, top=side, bottom=side)
    header_style = ensure_named_style(
        ws.parent,
        "InvSummaryHeader",
        font=header_font,
        border=border,
        alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
        fill=header_fill,
    )
    body_style = ensure_named_style(
        ws.parent,
        "InvSummaryBody",
        font=body_font,
        border=border,
        alignment=Alignment(vertical="center", wrap_text=True),
    )
    formatted_styles = {
        column: ensure_named_style(
            ws.parent,
            format_style_name("InvSummaryBody", column, number_format),
            font=body_font,
            border=border,
            alignment=Alignment(vertical="center", wrap_text=True),
            number_format=number_format,
        )
        for column, number_format in (number_formats or {}).items()
    }
    total_style = ensure_named_style(
        ws.parent,
        "InvSummaryTotal",
        font=bold_font,
        border=border,
        alignment=Alignment(vertical="center", wrap_text=True),
        fill=light_fill,
    )
    formatted_total_styles = {
        column: ensure_named_style(
            ws.parent,
            format_style_name("InvSummaryTotal", column, number_format),
            font=bold_font,
            border=border,
            alignment=Alignment(vertical="center", wrap_text=True),
            fill=light_fill,
            number_format=number_format,
        )
        for column, number_format in (number_formats or {}).items()
    }
    for cell in ws[header_row]:
        cell.style = header_style
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, max_col=max_col):
        for cell in row:
            cell.style = formatted_styles.get(cell.column, body_style)
    if ws.max_row > header_row and ws.cell(ws.max_row, 1).value == "총합계":
        for cell in ws[ws.max_row]:
            cell.style = formatted_total_styles.get(cell.column, total_style)
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
        "배치번호", "잔여 유통기한율", "개봉 전 사용기한(개월)", "판정",
    ]
    ws.append(headers)
    for row_index, item in enumerate(rows, start=2):
        write_safe_text(ws.cell(row_index, 1), item["저장 위치"])
        write_safe_text(ws.cell(row_index, 2), item["자재"])
        write_safe_text(ws.cell(row_index, 3), item["자재내역"])
        ws.cell(row_index, 4).value = item["가용재고(ASN)"]
        ws.cell(row_index, 5).value = item["유통기한"]
        write_safe_text(ws.cell(row_index, 6), item["로트번호(생산)"])
        write_safe_text(
            ws.cell(row_index, 7),
            item.get("배치번호") or make_batch_number(item["유통기한"], item["로트번호(생산)"]),
        )
        ws.cell(row_index, 8).value = item["잔여 유통기한율"]
        ws.cell(row_index, 9).value = item["개봉 전 사용기한(개월)"]
        write_safe_text(ws.cell(row_index, 10), item["판정"])

    apply_standard_style(
        ws,
        len(headers),
        number_formats={4: "#,##0", 5: "yyyy-mm-dd", 8: "0.0%", 9: "0.##"},
    )
    for col, width in {
        "A": 12, "B": 18, "C": 45, "D": 15, "E": 13,
        "F": 18, "G": 30, "H": 16, "I": 20, "J": 18,
    }.items():
        ws.column_dimensions[col].width = width
    if rows:
        add_table(ws, table_name, f"A1:J{len(rows) + 1}")
        ws.conditional_formatting.add(
            f"H2:H{len(rows) + 1}",
            CellIsRule(
                operator="lessThan",
                formula=[str(settings.pass_threshold)],
                fill=PatternFill("solid", fgColor="FCE4D6"),
            ),
        )
    else:
        ws.auto_filter.ref = "A1:J1"


def aggregate_by_product(rows: Iterable[dict]) -> dict[str, dict]:
    grouped: dict[str, dict] = {}
    for item in rows:
        key = item["자재"]
        if key not in grouped:
            grouped[key] = {
                "자재": item["자재"],
                "자재내역": item["자재내역"],
                "qty": defaultdict(float),
            }
        elif not grouped[key]["자재내역"] and item["자재내역"]:
            grouped[key]["자재내역"] = item["자재내역"]
        grouped[key]["qty"][item["저장 위치"]] += item["가용재고(ASN)"] or 0
    return grouped


def product_sort_key(material: str):
    if str(material).isdigit():
        return 0, int(material)
    return 1, str(material)


def write_pivot_sheet(wb: Workbook, rows: list[dict], settings: Settings) -> float:
    ws = wb.create_sheet("글로벌 pv")
    ws["A1"] = "합계 : 가용재고(ASN)"
    ws["C1"] = "저장 위치"
    headers = ["자재", "자재내역", *settings.locations, "총합계"]
    ws.append(headers)
    grouped = aggregate_by_product(rows)
    location_totals = defaultdict(float)
    grand_total = 0.0
    start_row = 3
    for row_number, key in enumerate(sorted(grouped, key=product_sort_key), start=start_row):
        item = grouped[key]
        write_safe_text(ws.cell(row_number, 1), item["자재"])
        write_safe_text(ws.cell(row_number, 2), item["자재내역"])
        total = 0.0
        for column, location in enumerate(settings.locations, start=3):
            quantity = item["qty"].get(location, 0)
            ws.cell(row_number, column).value = int(quantity) if float(quantity).is_integer() else quantity
            total += quantity
            location_totals[location] += quantity
        total_column = 3 + len(settings.locations)
        ws.cell(row_number, total_column).value = int(total) if float(total).is_integer() else total
        grand_total += total

    total_row = start_row + len(grouped)
    ws.cell(total_row, 1).value = "총합계"
    for column, location in enumerate(settings.locations, start=3):
        total = location_totals[location]
        ws.cell(total_row, column).value = int(total) if float(total).is_integer() else total
    total_column = 3 + len(settings.locations)
    ws.cell(total_row, total_column).value = int(grand_total) if float(grand_total).is_integer() else grand_total

    light_fill = PatternFill("solid", fgColor="D9EAF7")
    bold_font = Font(name="맑은 고딕", bold=True, size=10)
    side = Side(style="thin", color="D9E2F3")
    border = Border(left=side, right=side, top=side, bottom=side)
    for cell in ws[1]:
        cell.font, cell.fill, cell.border = bold_font, light_fill, border

    style_summary_sheet(
        ws,
        len(headers),
        header_row=2,
        freeze_cell="A3",
        number_formats={column: "#,##0" for column in range(3, len(headers) + 1)},
    )
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 45
    for column in range(3, len(headers) + 1):
        letter = ws.cell(1, column).column_letter
        ws.column_dimensions[letter].width = 14
    if grouped:
        add_table(ws, "GlobalExpiryPVTable", f"A2:{ws.cell(2, len(headers)).column_letter}{total_row - 1}")
    else:
        ws.auto_filter.ref = f"A2:{ws.cell(2, len(headers)).column_letter}2"
    return grand_total


def sort_location_rows(result: list[dict], settings: Settings) -> None:
    def material_key(item: dict):
        material = item["자재"]
        return (0, int(material)) if str(material).isdigit() else (1, str(material))

    # LOT 정렬을 먼저 적용한 뒤 나머지 기준을 하나의 안정 정렬로 묶는다.
    result.sort(key=lambda x: x["로트번호(생산)"], reverse=not settings.lot_ascending)
    result.sort(
        key=lambda x: (
            material_key(x),
            x["EXP"] is None,
            (x["EXP"] or date.min).toordinal() * (1 if settings.expiry_ascending else -1),
            x["수량"] * (1 if settings.quantity_ascending else -1),
        )
    )


def aggregate_all_location_exp_rows(
    rows: list[dict],
    settings: Settings,
) -> dict[str, list[dict]]:
    grouped_by_location: dict[str, dict[tuple, dict]] = {
        location: {} for location in settings.locations
    }
    for item in rows:
        location = item["저장 위치"]
        grouped = grouped_by_location.get(location)
        if grouped is None:
            continue
        quantity = item["가용재고(ASN)"] or 0
        if quantity == 0 and not settings.include_zero_in_details:
            continue
        lot_no = normalize_lot_number(item["로트번호(생산)"])
        key = (
            item["자재"], item["유통기한"], lot_no,
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
                "배치번호": make_batch_number(item["유통기한"], lot_no),
            }
        elif not grouped[key]["자재내역"] and item["자재내역"]:
            grouped[key]["자재내역"] = item["자재내역"]
        grouped[key]["수량"] += quantity

    result_by_location: dict[str, list[dict]] = {}
    for location, grouped in grouped_by_location.items():
        result = list(grouped.values())
        sort_location_rows(result, settings)
        result_by_location[location] = result
    return result_by_location


def aggregate_location_exp_rows(rows: list[dict], location: str, settings: Settings) -> list[dict]:
    """단일 창고 호출과 기존 테스트 호환을 위한 래퍼."""
    if location not in settings.locations:
        return []
    return aggregate_all_location_exp_rows(rows, settings)[location]


def write_location_sheets(wb: Workbook, rows: list[dict], settings: Settings) -> float:
    headers = [
        "자재", "자재내역", "수량", "EXP", "잔여 유통기한율",
        "개봉 전 사용기한(개월)", "로트번호(생산)", "배치번호",
    ]
    details_by_location = aggregate_all_location_exp_rows(rows, settings)
    grand_total = 0.0
    for location in settings.locations:
        ws = wb.create_sheet(f"글로벌 유통기한 {location}")
        ws.append(headers)
        detail_rows = details_by_location[location]
        for item in detail_rows:
            quantity = item["수량"]
            row_index = ws.max_row + 1
            write_safe_text(ws.cell(row_index, 1), item["자재"])
            write_safe_text(ws.cell(row_index, 2), item["자재내역"])
            ws.cell(row_index, 3).value = int(quantity) if float(quantity).is_integer() else quantity
            ws.cell(row_index, 4).value = item["EXP"]
            ws.cell(row_index, 5).value = item["잔여 유통기한율"]
            ws.cell(row_index, 6).value = item["개봉 전 사용기한(개월)"]
            write_safe_text(ws.cell(row_index, 7), item["로트번호(생산)"])
            write_safe_text(ws.cell(row_index, 8), item["배치번호"])
        total_row = ws.max_row + 1
        location_total = sum(item["수량"] for item in detail_rows)
        grand_total += location_total
        ws.cell(total_row, 1).value = "총합계"
        ws.cell(total_row, 3).value = location_total
        style_summary_sheet(
            ws,
            len(headers),
            number_formats={3: "#,##0", 4: "yyyy-mm-dd", 5: "0.0%", 6: "0.##"},
        )
        for col, width in {
            "A": 18, "B": 45, "C": 14, "D": 13, "E": 16, "F": 20, "G": 28, "H": 30,
        }.items():
            ws.column_dimensions[col].width = width
        if detail_rows:
            safe_location = re.sub(r"[^A-Za-z0-9_]", "", location)[:16] or "LOC"
            suffix = hashlib.sha1(location.encode("utf-8")).hexdigest()[:8]
            add_table(ws, f"T_{safe_location}_{suffix}", f"A1:H{total_row - 1}")
            ws.conditional_formatting.add(
                f"E2:E{total_row - 1}",
                CellIsRule(
                    operator="lessThan",
                    formula=[str(settings.pass_threshold)],
                    fill=PatternFill("solid", fgColor="FCE4D6"),
                ),
            )
        else:
            ws.auto_filter.ref = "A1:H1"
    return grand_total


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
    summary: ValidationSummary,
) -> None:
    ws = wb.create_sheet("기준_자동화")
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
        ["선택 창고 전체 수량", summary.source_quantity],
        ["유지 결과 수량", summary.kept_quantity],
        ["마스터 매칭률", summary.master_match_ratio],
        ["원본 창고별 행 수", " / ".join(
            f"{loc} {summary.location_rows.get(loc, 0):,}" for loc in settings.locations
        )],
        ["유지 창고별 행 수", " / ".join(
            f"{loc} {summary.kept_location_rows.get(loc, 0):,}" for loc in settings.locations
        )],
        ["주의: 유통기한 없음", summary.missing_expiry_rows],
        ["주의: Product Master 사용기한 누락", summary.missing_master_rows],
        ["주의: 로트번호 없음", summary.missing_lot_rows],
        ["참고: 가용재고 0", summary.zero_quantity_rows],
        ["생성 앱 버전", APP_VERSION],
    ]
    for row in meta:
        ws.append(row)
    apply_standard_style(ws, 2)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 90
    ws["B4"].number_format = "yyyy-mm-dd"
    ws["B6"].number_format = "yyyy-mm-dd"
    ws["B16"].number_format = "0.0%"
    for row in ws.iter_rows(min_row=2, max_col=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def write_validation_sheet(wb: Workbook, summary: ValidationSummary) -> None:
    ws = wb.create_sheet("검증결과")
    rows = [
        ["검증 항목", "결과"],
        ["판토스 전체 원본 행", summary.source_rows],
        ["선택 창고 처리 행", summary.selected_rows],
        ["결과 유지 행", summary.kept_rows],
        ["선택 창고 전체 수량", summary.source_quantity],
        ["결과 유지 수량", summary.kept_quantity],
        ["마스터 전체 행", summary.master_rows],
        ["마스터 빈 바코드 행", summary.master_blank_barcode_rows],
        ["마스터 사용기한 공란 행", summary.master_blank_months_rows],
        ["재고 마스터 매칭 행", summary.master_match_rows],
        ["재고 마스터 매칭률", summary.master_match_ratio],
        ["유통기한 공란 행", summary.missing_expiry_rows],
        ["마스터 미매칭/공란 행", summary.missing_master_rows],
        ["LOT 공란 행", summary.missing_lot_rows],
        ["제품명 공란 행", summary.missing_product_name_rows],
        ["0재고 행", summary.zero_quantity_rows],
        ["동일 자재 제품명 불일치", summary.product_name_conflicts],
        ["오류 검증", "통과"],
    ]
    for row in rows:
        ws.append(row)
    warning_start = ws.max_row + 2
    ws.cell(warning_start, 1).value = "주의사항"
    ws.cell(warning_start, 2).value = "내용"
    for warning in summary.warnings or ["없음"]:
        ws.append(["주의", warning])
    apply_standard_style(ws, 2)
    for cell in ws[warning_start]:
        cell.fill = PatternFill("solid", fgColor="BF9000")
        cell.font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 100
    ws["B11"].number_format = "0.0%"
    for row in ws.iter_rows(min_row=2, max_col=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def validate_output_workbook(
    output_bytes: bytes,
    global_rows: list[dict],
    filtered_rows: list[dict],
    settings: Settings,
) -> None:
    expected_sheets = [
        "글로벌", "글로벌 유통기한", "글로벌 pv",
        *(f"글로벌 유통기한 {location}" for location in settings.locations),
        "검증결과", "기준_자동화",
    ]
    with ZipFile(io.BytesIO(output_bytes)) as archive:
        required_parts = {"[Content_Types].xml", "xl/workbook.xml"}
        if not required_parts.issubset(archive.namelist()):
            raise RuntimeError("출력 XLSX의 필수 구성요소가 누락되었습니다.")

    wb = load_workbook(io.BytesIO(output_bytes), data_only=True, read_only=True)
    try:
        if wb.sheetnames != expected_sheets:
            raise RuntimeError("출력 시트 구성이 예상과 다릅니다.")
        if wb["글로벌"].max_row - 1 != len(global_rows):
            raise RuntimeError("글로벌 시트 행 수 검증에 실패했습니다.")
        if wb["글로벌 유통기한"].max_row - 1 != len(filtered_rows):
            raise RuntimeError("글로벌 유통기한 시트 행 수 검증에 실패했습니다.")

        if global_rows and wb["글로벌"]["G2"].value != global_rows[0]["배치번호"]:
            raise RuntimeError("배치번호 저장값 검증에 실패했습니다.")
        if global_rows and (wb["글로벌"]["H2"].value != global_rows[0]["잔여 유통기한율"]):
            raise RuntimeError("잔여 유통기한율 저장값 검증에 실패했습니다.")
        if global_rows and wb["글로벌"]["J2"].value != global_rows[0]["판정"]:
            raise RuntimeError("판정 저장값 검증에 실패했습니다.")
        if filtered_rows and wb["글로벌 유통기한"]["G2"].value != filtered_rows[0]["배치번호"]:
            raise RuntimeError("글로벌 유통기한 배치번호 저장값 검증에 실패했습니다.")

        detail_rows_by_location = aggregate_all_location_exp_rows(filtered_rows, settings)
        for location, detail_rows in detail_rows_by_location.items():
            if not detail_rows:
                continue
            sheet_name = f"글로벌 유통기한 {location}"
            if wb[sheet_name]["H2"].value != detail_rows[0]["배치번호"]:
                raise RuntimeError(f"{sheet_name} 배치번호 저장값 검증에 실패했습니다.")
    finally:
        wb.close()


def create_inventory_workbook(
    pantos_bytes: bytes,
    master_bytes: bytes,
    pantos_name: str,
    master_name: str,
    base: date,
    base_source: str,
    filename_date: Optional[date],
    settings: Settings,
) -> tuple[bytes, list[dict], list[dict], ValidationSummary]:
    summary = ValidationSummary()
    lifetime_map = load_lifetime_months(master_bytes, settings, summary)
    global_rows, filtered_rows = build_source_rows(pantos_bytes, lifetime_map, base, settings, summary)

    wb = Workbook()
    try:
        wb.remove(wb.active)
        write_data_sheet(wb, "글로벌", global_rows, settings, "GlobalInventoryTable")
        write_data_sheet(wb, "글로벌 유통기한", filtered_rows, settings, "GlobalExpiryDataTable")
        pivot_total = write_pivot_sheet(wb, filtered_rows, settings)
        detail_total = write_location_sheets(wb, filtered_rows, settings)
        if not math.isclose(float(pivot_total), float(summary.kept_quantity), abs_tol=1e-9):
            raise RuntimeError("피벗 총수량 대사에 실패했습니다.")
        if not math.isclose(float(detail_total), float(summary.kept_quantity), abs_tol=1e-9):
            raise RuntimeError("창고별 상세 총수량 대사에 실패했습니다.")
        write_validation_sheet(wb, summary)
        write_meta_sheet(
            wb, pantos_name, master_name, base, base_source, filename_date,
            global_rows, filtered_rows, settings, summary,
        )
        output = io.BytesIO()
        wb.save(output)
    finally:
        wb.close()
    output_bytes = output.getvalue()
    validate_output_workbook(output_bytes, global_rows, filtered_rows, settings)
    return output_bytes, global_rows, filtered_rows, summary


def parse_locations(text: str) -> tuple[str, ...]:
    locations: list[str] = []
    for token in re.split(r"[,>→\s]+", text.upper().strip()):
        if token and token not in locations:
            if not re.fullmatch(r"[A-Z0-9_-]{1,12}", token):
                raise ValueError(
                    f"창고 코드 '{token}'가 올바르지 않습니다. 영문, 숫자, 밑줄, 하이픈만 사용할 수 있습니다."
                )
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
    stem = name[:-5].rstrip(". ")[:175] or "글로벌재고_자동생성"
    return stem + ".xlsx"


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
        st.caption("옵션을 변경한 뒤 아래 적용 버튼을 눌러 주세요.")
        with st.form("inventory_options_form", clear_on_submit=False):
            base_mode = st.radio("기준일", ["자동", "직접 지정"], horizontal=True)
            manual_date = st.date_input(
                "직접 기준일",
                value=today_kst(),
                help="기준일이 '자동'이면 이 값은 사용하지 않습니다.",
            )
            threshold_percent = st.number_input(
                "잔여 유통기한 통과율 (%)", min_value=0.0, max_value=100.0,
                value=66.0, step=1.0,
            )
            location_text = st.text_input("대상 창고 및 순서", value="GD, DU, GL")

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
                    value=today_kst().year, step=1,
                )
                split_day = st.number_input(
                    "전반기 기준일", min_value=1, max_value=28, value=15, step=1,
                    help="파일 날짜가 이 날짜 이하면 해당 날짜, 이후면 월말을 기준일로 사용합니다.",
                )
                reuse_existing = st.checkbox("기존 글로벌 시트 G1 기준일 우선 사용", value=True)

            with st.expander("입력 시트·열 이름"):
                stock_sheet = st.text_input("판토스 원본 시트", value="Sheet1")
                master_sheet = st.text_input("Product Master 시트", value="FG (제조 발주)")
                st.caption("판토스 열 제목은 1행, Product Master 열 제목은 12행에서 인식합니다.")
                location_col = st.text_input("창고 열", value="저장 위치")
                material_col = st.text_input("자재 열", value="자재")
                product_col = st.text_input("제품명 열", value="자재내역")
                stock_col = st.text_input("가용재고 열", value="가용재고(ASN)")
                expiry_col = st.text_input("유통기한 열", value="유통기한")
                lot_col = st.text_input("LOT 열", value="로트번호(생산)")
                barcode_col = st.text_input("마스터 바코드 열", value="바코드")
                months_col = st.text_input("마스터 사용기한 열", value="개봉 전 사용기한")

            st.subheader("출력")
            output_name = st.text_input("결과 파일명", value="글로벌재고_자동생성.xlsx")
            options_submitted = st.form_submit_button(
                "옵션 적용",
                type="primary",
                width="stretch",
            )

        if options_submitted:
            st.success("옵션이 적용되었습니다.")

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
        "output_name": output_name,
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

    output_name = controls["output_name"]

    if not pantos_file or not master_file:
        st.info("두 엑셀 파일을 업로드하면 재고 파일을 생성할 수 있습니다.")
        return

    pantos_bytes = pantos_file.getvalue()
    master_bytes = master_file.getvalue()

    oversized = []
    if len(pantos_bytes) > MAX_UPLOAD_BYTES:
        oversized.append(f"판토스 파일 ({len(pantos_bytes) / 1024 / 1024:.1f}MB)")
    if len(master_bytes) > MAX_UPLOAD_BYTES:
        oversized.append(f"Product Master ({len(master_bytes) / 1024 / 1024:.1f}MB)")
    if oversized:
        st.error(f"파일당 최대 크기는 {MAX_UPLOAD_BYTES / 1024 / 1024:.0f}MB입니다: " + ", ".join(oversized))
        return

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
    except (BadZipFile, InvalidFileException, OSError, ValueError, TypeError, InputFormatError) as exc:
        st.error(f"업로드 파일 또는 설정을 확인해 주세요: {exc}")
        return
    except Exception as exc:
        st.error(f"업로드 파일을 확인할 수 없습니다: {type(exc).__name__}: {exc}")
        return

    a, b, c = st.columns(3)
    a.metric("기준일", base.strftime("%Y-%m-%d"))
    b.metric("통과율", f"{settings.pass_threshold:.0%}")
    c.metric("창고 순서", " → ".join(settings.locations))
    st.caption(f"기준일 출처: {base_source}")
    if base_source == "입력 파일 기존 글로벌 시트 G1" and filename_date:
        filename_base = auto_base_date(filename_date, controls["split_day"])
        if filename_base != base:
            st.warning(
                f"기존 G1 기준일({base:%Y-%m-%d})과 파일명 자동 기준일({filename_base:%Y-%m-%d})이 다릅니다."
            )
    elif filename_date is None and controls["base_mode"] == "자동":
        st.warning("파일명에서 날짜를 찾지 못해 한국시간 오늘 날짜로 기준일을 계산했습니다.")

    confirmation_key = f"base_confirmed_{base.isoformat()}_{hashlib.sha1(base_source.encode()).hexdigest()[:8]}"
    with st.form("inventory_generation_form", clear_on_submit=False):
        base_confirmed = st.checkbox(
            f"기준일 {base:%Y-%m-%d}을 확인했습니다.",
            key=confirmation_key,
        )
        generate_submitted = st.form_submit_button(
            "재고 파일 만들기",
            type="primary",
            width="stretch",
        )

    if generate_submitted and not base_confirmed:
        st.error("표시된 기준일을 확인한 뒤 체크해 주세요.")
    elif generate_submitted:
        try:
            with st.spinner("재고 데이터를 처리하고 있습니다..."):
                output_bytes, global_rows, filtered_rows, summary = create_inventory_workbook(
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
                "source_quantity": summary.source_quantity,
                "filtered_quantity": summary.kept_quantity,
                "master_match_ratio": summary.master_match_ratio,
                "warnings": summary.warnings,
                "base": base,
                "preview": filtered_rows[:100],
                "signature": current_signature,
            }
        except (InputFormatError, DataValidationError, ValueError, KeyError, TypeError) as exc:
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
        q1, q2, q3 = st.columns(3)
        q1.metric("선택 창고 전체 수량", f"{result['source_quantity']:,.0f}")
        q2.metric("결과 유지 수량", f"{result['filtered_quantity']:,.0f}")
        q3.metric("마스터 매칭률", f"{result['master_match_ratio']:.1%}")
        for warning in result.get("warnings", []):
            st.warning(warning)
        st.download_button(
            "결과 엑셀 다운로드",
            data=result["bytes"],
            file_name=sanitize_output_name(output_name),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
            on_click="ignore",
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
            st.dataframe(preview, width="stretch", hide_index=True)


if __name__ == "__main__":
    main()

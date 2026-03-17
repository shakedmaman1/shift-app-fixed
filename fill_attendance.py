#!/usr/bin/env python3
"""
מילוי קובץ נוכחות מתוך קובץ סידור עבודה.

ספריות נדרשות:
    pip install pandas openpyxl xlrd beautifulsoup4

דוגמת הרצה:
    python fill_attendance.py ^
        --source "C:\\Users\\Shaked\\Downloads\\מחלקת שימור.xls" ^
        --template "C:\\Users\\Shaked\\OneDrive\\Desktop\\נוכחות ריק.xlsx" ^
        --output "C:\\Users\\Shaked\\OneDrive\\Desktop\\נוכחות ממולא.xlsx"

הערות:
1. הכתיבה לקובץ היעד נעשית עם openpyxl בלבד, כדי לא לפגוע בעיצוב ובנוסחאות.
2. הקריאה לקובץ המקור תומכת בשני מצבים:
   - קובץ Excel "רגיל" עם שמות בשורות וימים בעמודות.
   - קובץ .xls שהוא בפועל HTML (כמו WeekDisplay / Partner Shift).
"""

from __future__ import annotations

import argparse
import importlib
import re
import sys
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable, Optional


TARGET_HEADERS = ("צוות", "שם", "משמרת", "משך זמן", "הערה", "הערה")
IGNORE_DAY_VALUES = ("חופש", "מחלה", "off", "vacation", "חופשתי")
SPECIAL_STATUS_KEYWORDS = (
    "טיפולים",
    "כתף",
    "ריחוף",
    "אימון",
    "הדרכה",
    "חניכה",
    "ליווי",
    "ממ",
    "אחמ",
    "איחור",
)
SHEET_BY_DAY = {
    "א": "יום א",
    "ב": "יום ב",
    "ג": "יום ג",
    "ד": "יום ד",
    "ה": "יום ה",
    "ו": "יום ו",
}
SHIFT_RE = re.compile(
    r"(?P<start>\d{1,2}:\d{2})\s*[-–—]\s*(?P<end>\d{1,2}:\d{2})"
)
DATE_RE = re.compile(r"(?P<day>\d{1,2})[./-](?P<month>\d{1,2})(?:[./-](?P<year>\d{2,4}))?")
EVENING_START_TIME = time(15, 30)
DIVIDER_TITLES = {"משמרת ערב", "משמרת בוקר"}


def import_or_raise(module_name: str, install_hint: str) -> Any:
    """טוען ספרייה דינמית ומחזיר שגיאה ברורה אם היא חסרה."""
    try:
        return importlib.import_module(module_name)
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            f"חסרה הספרייה '{module_name}'. התקן אותה עם: {install_hint}"
        ) from exc


@dataclass
class ShiftRecord:
    day_sheet: str
    team: str
    name: str
    shift_text: str
    duration_hours: int | float
    break_text: str
    special_status: str
    start_time: time


@dataclass
class DisplayRow:
    row_type: str
    record: Optional[ShiftRecord] = None
    title: str = ""


def clean_text(value: object) -> str:
    """מנרמל כל ערך טבלאי למחרוזת נקייה."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    try:
        if value != value:
            return ""
    except Exception:
        pass

    text = str(value)
    text = text.replace("\xa0", " ").replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_label(value: object) -> str:
    """מנרמל כותרות כדי להשוות בלי רווחים/סימנים."""
    text = clean_text(value).lower()
    text = text.replace("׳", "'").replace("״", '"')
    return re.sub(r"[\s\"'`.,:;!?\-_/\\()\[\]{}]+", "", text)


def unique_join(values: Iterable[str]) -> str:
    """מאחד טקסטים לא-ריקים בלי כפילויות."""
    seen: set[str] = set()
    output: list[str] = []
    for value in values:
        text = clean_text(value)
        if not text:
            continue
        key = normalize_label(text)
        if key in seen:
            continue
        seen.add(key)
        output.append(text)
    return " | ".join(output)


def infer_default_team(source_path: Path) -> str:
    """מנסה להבין את שם הצוות מתוך שם הקובץ."""
    source_name = clean_text(source_path.stem).lower()
    if "שימור" in source_name or "retention" in source_name:
        return "שימור"
    if "שירות" in source_name or "service" in source_name:
        return "שירות"
    if "תמיכה" in source_name or "support" in source_name:
        return "תמיכה"
    return ""


def infer_team_from_role(role_text: str, fallback_team: str) -> str:
    """ממפה תיאור תפקיד/שורה לצוות כללי ככל האפשר."""
    role = clean_text(role_text)
    if "שימור" in role:
        return "שימור"
    if "שירות" in role:
        return "שירות"
    if "תמיכה" in role:
        return "תמיכה"
    return fallback_team


def map_header_to_sheet(header_value: object) -> Optional[str]:
    """ממפה כותרת יום/תאריך לשם הגיליון בתבנית."""
    raw = clean_text(header_value)
    if not raw:
        return None

    normalized = raw.replace("׳", "'").replace("״", '"')
    normalized = re.sub(r"\s+", " ", normalized)

    explicit_map = {
        "יום א": "יום א",
        "ראשון": "יום א",
        "יום ב": "יום ב",
        "שני": "יום ב",
        "יום ג": "יום ג",
        "שלישי": "יום ג",
        "יום ד": "יום ד",
        "רביעי": "יום ד",
        "יום ה": "יום ה",
        "חמישי": "יום ה",
        "יום ו": "יום ו",
        "שישי": "יום ו",
    }
    for key, value in explicit_map.items():
        if key in normalized:
            return value

    day_letter_match = re.search(r"(?:^|[\s\-])'?(א|ב|ג|ד|ה|ו)(?:$|[\s\-])", normalized)
    if day_letter_match:
        return SHEET_BY_DAY.get(day_letter_match.group(1))

    date_match = DATE_RE.search(normalized)
    if date_match:
        day_num = int(date_match.group("day"))
        month_num = int(date_match.group("month"))
        year_num = date_match.group("year")
        if year_num is None:
            year = datetime.now().year
        else:
            year = int(year_num)
            if year < 100:
                year += 2000
        try:
            parsed_date = date(year, month_num, day_num)
        except ValueError:
            return None

        # Python: Monday=0 ... Sunday=6
        day_by_python_weekday = {
            6: "יום א",
            0: "יום ב",
            1: "יום ג",
            2: "יום ד",
            3: "יום ה",
            4: "יום ו",
        }
        return day_by_python_weekday.get(parsed_date.weekday())

    return None


def parse_shift_range(text: str) -> tuple[str, time, time]:
    """שולף טווח שעות מתוך טקסט חופשי."""
    match = SHIFT_RE.search(clean_text(text))
    if not match:
        raise ValueError(f"לא נמצא טווח שעות חוקי בתוך: {text!r}")

    start_str = match.group("start")
    end_str = match.group("end")
    start_time = datetime.strptime(start_str, "%H:%M").time()
    end_time = datetime.strptime(end_str, "%H:%M").time()
    return f"{start_str}-{end_str}", start_time, end_time


def calculate_duration_and_break(start_time: time, end_time: time) -> tuple[int | float, str]:
    """מחשב משך משמרת ושעת הפסקה לפי 7 דקות לכל שעה מלאה."""
    anchor_day = date(2000, 1, 1)
    start_dt = datetime.combine(anchor_day, start_time)
    end_dt = datetime.combine(anchor_day, end_time)
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)

    total_minutes = int((end_dt - start_dt).total_seconds() // 60)
    total_hours = total_minutes / 60

    if total_hours.is_integer():
        duration_hours: int | float = int(total_hours)
    else:
        duration_hours = round(total_hours, 2)

    break_minutes = (total_minutes // 60) * 7
    return duration_hours, f"{break_minutes} דקות"


def should_ignore_day_cell(text: str) -> bool:
    """קובע אם תא יומי צריך להידלג."""
    normalized = clean_text(text).lower()
    if not normalized:
        return True
    return any(keyword in normalized for keyword in IGNORE_DAY_VALUES)


def extract_shift_and_inline_note(cell_text: str) -> tuple[Optional[str], str]:
    """שולף טווח שעות והערה מתוך תא יומי."""
    text = clean_text(cell_text)
    match = SHIFT_RE.search(text)
    if not match:
        return None, ""

    shift_text = f"{match.group('start')}-{match.group('end')}"
    inline_note = text[match.end() :].strip(" -–—")
    inline_note = clean_text(inline_note)
    return shift_text, inline_note


def build_record(
    *,
    day_sheet: str,
    team: str,
    name: str,
    shift_text: str,
    special_status: str = "",
) -> ShiftRecord:
    """יוצר אובייקט אחיד עבור שיבוץ בודד."""
    shift_text, start_time, end_time = parse_shift_range(shift_text)
    duration_hours, break_text = calculate_duration_and_break(start_time, end_time)
    return ShiftRecord(
        day_sheet=day_sheet,
        team=team,
        name=clean_text(name),
        shift_text=shift_text,
        duration_hours=duration_hours,
        break_text=break_text,
        special_status=clean_text(special_status),
        start_time=start_time,
    )


def is_html_disguised_excel(source_path: Path) -> bool:
    """בודק אם הקובץ הוא HTML שמתחפש ל-.xls."""
    try:
        head = source_path.read_text(encoding="utf-8", errors="ignore")[:1024].lower()
    except OSError:
        return False
    return "<html" in head or "<table" in head


def parse_html_schedule(source_path: Path) -> list[ShiftRecord]:
    """קורא קובץ HTML/‏XLS בתצוגת שבוע ומוציא ממנו שיבוצים."""
    bs4 = import_or_raise("bs4", "pip install beautifulsoup4")
    BeautifulSoup = bs4.BeautifulSoup

    html = source_path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if table is None:
        return []

    rows = []
    for row in table.find_all("tr"):
        cells = row.find_all("td")
        if not cells:
            continue
        rows.append((row, [clean_text(cell.get_text(" ", strip=True)) for cell in cells]))

    if not rows:
        return []

    header_labels = rows[0][1]
    normalized_headers = {normalize_label(label) for label in header_labels}
    flat_export_headers = {
        "תאריך",
        "שעתהתחלה",
        "שעתסיום",
        "עובדמשובץ",
    }
    if flat_export_headers.issubset(normalized_headers):
        return parse_html_flat_export(rows, source_path)

    return parse_html_week_display(rows, source_path)


def parse_html_flat_export(rows, source_path: Path) -> list[ShiftRecord]:
    """קורא קובץ HTML שטוח שבו כל שורה מייצגת משמרת אחת."""
    headers = rows[0][1]
    header_map = {normalize_label(label): index for index, label in enumerate(headers)}

    date_index = header_map.get("תאריך")
    day_index = header_map.get("יום")
    start_index = header_map.get("שעתהתחלה")
    end_index = header_map.get("שעתסיום")
    name_index = header_map.get("עובדמשובץ")
    role_index = header_map.get("תפקיד")
    category_index = header_map.get("קטגוריה")
    description_index = header_map.get("תאור") or header_map.get("תיאור")

    required_indexes = [date_index, start_index, end_index, name_index]
    if any(index is None for index in required_indexes):
        return []

    default_team = infer_default_team(source_path)
    generic_roles = {"נציג שירות", "נציג שימור", "נציג תמיכה"}
    records: list[ShiftRecord] = []

    for _, row_values in rows[1:]:
        def value_at(index: Optional[int]) -> str:
            if index is None or index >= len(row_values):
                return ""
            return clean_text(row_values[index])

        scheduled_name = value_at(name_index)
        if not scheduled_name:
            continue

        start_time = value_at(start_index)
        end_time = value_at(end_index)
        if not start_time or not end_time:
            continue

        day_sheet = map_header_to_sheet(value_at(date_index)) or map_header_to_sheet(value_at(day_index))
        if not day_sheet:
            continue

        role_text = value_at(role_index)
        team = infer_team_from_role(role_text, default_team)
        role_as_status = role_text if role_text and role_text not in generic_roles else ""
        special_status = unique_join(
            [
                role_as_status,
                value_at(category_index),
                value_at(description_index),
            ]
        )

        records.append(
            build_record(
                day_sheet=day_sheet,
                team=team,
                name=scheduled_name,
                shift_text=f"{start_time}-{end_time}",
                special_status=special_status,
            )
        )

    return records


def parse_html_week_display(rows, source_path: Path) -> list[ShiftRecord]:
    """קורא קובץ HTML שבועי שבו כל תא יום מכיל כמה נציגים."""
    default_team = infer_default_team(source_path)
    day_sheets: list[Optional[str]] = []
    records: list[ShiftRecord] = []

    for row, row_texts in rows:
        cells = row.find_all("td")
        if not day_sheets and any("תפקיד" in text for text in row_texts):
            day_sheets = [map_header_to_sheet(text) for text in row_texts[1:]]
            continue

        # שורת כותרת / מפריד קבוצה (למשל "בוקר", "ערב")
        if len(cells) == 1 or cells[0].get("colspan"):
            continue

        if not day_sheets:
            continue

        role_text = clean_text(cells[0].get_text(" ", strip=True))
        if not role_text:
            continue

        for cell_index, day_sheet in enumerate(day_sheets, start=1):
            if day_sheet is None or cell_index >= len(cells):
                continue

            day_cell = cells[cell_index]
            top_level_entries = day_cell.find_all("div", recursive=False)
            if not top_level_entries:
                top_level_entries = [day_cell]

            for entry in top_level_entries:
                entry_text = clean_text(entry.get_text(" ", strip=True))
                if not entry_text:
                    continue

                # דוגמה: "08:00-15:00 איתי פזי - שלב אימון"
                match = re.match(
                    r"^(?P<shift>\d{1,2}:\d{2}\s*[-–—]\s*\d{1,2}:\d{2})\s+"
                    r"(?P<name>.*?)(?:\s+-\s+(?P<note>.*))?$",
                    entry_text,
                )
                if not match:
                    continue

                shift_text = clean_text(match.group("shift"))
                name = clean_text(match.group("name")).rstrip("*").strip()
                inline_note = clean_text(match.group("note"))
                if not name:
                    continue

                team = infer_team_from_role(role_text, default_team)
                records.append(
                    build_record(
                        day_sheet=day_sheet,
                        team=team,
                        name=name,
                        shift_text=shift_text,
                        special_status=inline_note,
                    )
                )

    return records


def find_header_row_and_day_columns(df: Any) -> tuple[Optional[int], dict[int, str]]:
    """מאתר שורה שבה מופיעים ימי השבוע/תאריכים."""
    best_row: Optional[int] = None
    best_day_columns: dict[int, str] = {}

    for row_index in range(min(len(df), 25)):
        day_columns: dict[int, str] = {}
        for col_index in range(df.shape[1]):
            day_sheet = map_header_to_sheet(df.iat[row_index, col_index])
            if day_sheet:
                day_columns[col_index] = day_sheet
        if len(day_columns) > len(best_day_columns):
            best_row = row_index
            best_day_columns = day_columns

    if len(best_day_columns) < 3:
        return None, {}
    return best_row, best_day_columns


def find_named_column(
    df: Any,
    candidate_rows: Iterable[int],
    predicates: Iterable[str],
) -> Optional[int]:
    """מחפש עמודה לפי מילות מפתח בכותרת."""
    normalized_needles = tuple(normalize_label(item) for item in predicates)
    for row_index in candidate_rows:
        if row_index < 0 or row_index >= len(df):
            continue
        for col_index in range(df.shape[1]):
            label = normalize_label(df.iat[row_index, col_index])
            if not label:
                continue
            if any(needle in label for needle in normalized_needles):
                return col_index
    return None


def collect_status_columns(df: Any, candidate_rows: Iterable[int]) -> list[int]:
    """מחזיר עמודות שנראות כמו סטטוס/הערות/סימונים צדדיים."""
    columns: list[int] = []
    for row_index in candidate_rows:
        if row_index < 0 or row_index >= len(df):
            continue
        for col_index in range(df.shape[1]):
            label = clean_text(df.iat[row_index, col_index])
            normalized = normalize_label(label)
            if not normalized:
                continue
            if (
                "סטטוס" in label
                or "הער" in label
                or any(keyword in label for keyword in SPECIAL_STATUS_KEYWORDS)
            ):
                if col_index not in columns:
                    columns.append(col_index)
    return columns


def fallback_row_status(df: Any, row_index: int, excluded_columns: set[int]) -> str:
    """כאשר אין עמודת סטטוס מזוהה, מחפש טקסט סטטוס בשאר השורה."""
    found_values: list[str] = []
    for col_index in range(df.shape[1]):
        if col_index in excluded_columns:
            continue
        text = clean_text(df.iat[row_index, col_index])
        if not text:
            continue
        if any(keyword in text for keyword in SPECIAL_STATUS_KEYWORDS):
            found_values.append(text)
    return unique_join(found_values)


def parse_weekly_grid_excel(source_path: Path) -> list[ShiftRecord]:
    """קורא קובץ Excel "רגיל" עם שמות בשורות וימים בעמודות."""
    pandas = import_or_raise("pandas", "pip install pandas")

    if source_path.suffix.lower() == ".xls":
        import_or_raise("xlrd", "pip install xlrd")
        excel = pandas.ExcelFile(source_path, engine="xlrd")
    else:
        import_or_raise("openpyxl", "pip install openpyxl")
        excel = pandas.ExcelFile(source_path, engine="openpyxl")

    all_records: list[ShiftRecord] = []
    default_team = infer_default_team(source_path)

    for sheet_name in excel.sheet_names:
        df = excel.parse(sheet_name, header=None)
        if df.empty:
            continue

        header_row, day_columns = find_header_row_and_day_columns(df)
        if header_row is None:
            continue

        candidate_rows = range(max(0, header_row - 2), min(len(df), header_row + 3))
        name_col = find_named_column(df, candidate_rows, ("שם נציג", "שם", "נציג"))
        if name_col is None:
            continue

        team_col = find_named_column(df, candidate_rows, ("צוות", "מחלקה", "department"))
        status_cols = collect_status_columns(df, candidate_rows)

        data_start_row = header_row + 1
        for row_index in range(data_start_row, len(df)):
            name = clean_text(df.iat[row_index, name_col])
            if not name:
                continue
            if normalize_label(name) in {"שם", "שםנציג", "נציג"}:
                continue

            team = clean_text(df.iat[row_index, team_col]) if team_col is not None else ""
            if not team:
                team = default_team

            row_status = unique_join(
                clean_text(df.iat[row_index, col_index]) for col_index in status_cols
            )
            if not row_status:
                excluded = set(day_columns) | {name_col}
                if team_col is not None:
                    excluded.add(team_col)
                row_status = fallback_row_status(df, row_index, excluded)

            for col_index, day_sheet in day_columns.items():
                raw_value = clean_text(df.iat[row_index, col_index])
                if should_ignore_day_cell(raw_value):
                    continue

                shift_text, inline_note = extract_shift_and_inline_note(raw_value)
                if not shift_text:
                    continue

                special_status = unique_join([row_status, inline_note])
                all_records.append(
                    build_record(
                        day_sheet=day_sheet,
                        team=team,
                        name=name,
                        shift_text=shift_text,
                        special_status=special_status,
                    )
                )

    return all_records


def load_source_records(source_path: Path) -> list[ShiftRecord]:
    """טוען את קובץ המקור לפי הסוג שזוהה."""
    if is_html_disguised_excel(source_path):
        records = parse_html_schedule(source_path)
        if records:
            return records
        raise RuntimeError(
            "קובץ המקור הוא HTML מחופש ל-.xls, אבל המבנה שלו לא נתמך עדיין."
        )

    records = parse_weekly_grid_excel(source_path)
    if records:
        return records

    raise RuntimeError(
        "לא הצלחתי לזהות שיבוצים מתוך קובץ המקור. "
        "ודא שהקובץ מכיל שמות/ימים/שעות או שהוא קובץ WeekDisplay/HTML תקין."
    )


def sort_records(records: list[ShiftRecord]) -> dict[str, list[ShiftRecord]]:
    """מקבץ רשומות לפי יום וממיין לפי שעת התחלה ואז שם."""
    grouped: dict[str, list[ShiftRecord]] = defaultdict(list)
    for record in records:
        grouped[record.day_sheet].append(record)

    for day_sheet in grouped:
        grouped[day_sheet].sort(key=lambda item: (item.start_time, item.name))
    return grouped


def build_display_rows(records: list[ShiftRecord]) -> list[DisplayRow]:
    """בונה את רצף השורות לתצוגה, כולל הפרדה למשמרת ערב."""
    morning_records = [record for record in records if record.start_time < EVENING_START_TIME]
    evening_records = [record for record in records if record.start_time >= EVENING_START_TIME]

    display_rows: list[DisplayRow] = [DisplayRow(row_type="record", record=record) for record in morning_records]
    if evening_records:
        display_rows.append(DisplayRow(row_type="divider", title="משמרת ערב"))
        display_rows.extend(DisplayRow(row_type="record", record=record) for record in evening_records)

    if not morning_records and not evening_records:
        return []
    return display_rows


def locate_target_table(ws) -> tuple[int, int]:
    """מאתר את הטבלה הראשית לפי ששת הכותרות הראשונות בלבד."""
    max_scan_row = min(ws.max_row, 20)
    max_scan_col = min(ws.max_column, 15)

    expected = tuple(normalize_label(header) for header in TARGET_HEADERS)
    for row_index in range(1, max_scan_row + 1):
        for col_index in range(1, max_scan_col - len(expected) + 2):
            labels = [
                normalize_label(ws.cell(row=row_index, column=col_index + offset).value)
                for offset in range(len(expected))
            ]
            if (
                labels[0] == expected[0]
                and labels[1] == expected[1]
                and labels[2] == expected[2]
                and labels[3].startswith(expected[3])
                and labels[4].startswith(expected[4])
                and labels[5].startswith(expected[5])
            ):
                return row_index, col_index

    raise RuntimeError(f"לא נמצאה טבלת היעד בגיליון {ws.title!r}")


def resolve_day_worksheets(workbook) -> dict[str, object]:
    """ממפה את שמות הגיליונות בפועל לימי העבודה בתבנית בצורה גמישה."""
    resolved: dict[str, object] = {}
    for worksheet in workbook.worksheets:
        mapped_day = map_header_to_sheet(worksheet.title)
        if mapped_day and mapped_day not in resolved:
            resolved[mapped_day] = worksheet

    return resolved


def detect_table_body_end(ws, header_row: int, start_col: int, width: int = 6) -> int:
    """מנסה לזהות היכן מסתיימות שורות הגוף בתבנית לפי גובה/עיצוב קיים."""
    body_start = header_row + 1
    template_height = ws.row_dimensions[body_start].height
    template_styles = [ws.cell(body_start, col)._style for col in range(start_col, start_col + width)]
    header_styles = [ws.cell(header_row, col)._style for col in range(start_col, start_col + width)]

    current_end = body_start
    max_scan_row = min(ws.max_row, body_start + 200)
    for row_index in range(body_start, max_scan_row + 1):
        row_title = clean_text(ws.cell(row=row_index, column=start_col).value)
        if row_index != body_start:
            if row_title in DIVIDER_TITLES:
                same_header_style = all(
                    ws.cell(row_index, col)._style == header_styles[offset]
                    for offset, col in enumerate(range(start_col, start_col + width))
                )
                if same_header_style:
                    current_end = row_index
                    continue

            row_height = ws.row_dimensions[row_index].height
            if template_height is not None and row_height not in (template_height, None):
                break

            same_style = all(
                ws.cell(row_index, col)._style == template_styles[offset]
                for offset, col in enumerate(range(start_col, start_col + width))
            )
            if not same_style:
                break

        current_end = row_index

    return current_end


def copy_row_format(ws, source_row: int, target_row: int) -> None:
    """מעתיק עיצוב שורה שלם בלי להעתיק ערכים."""
    merged_cell_module = import_or_raise("openpyxl.cell.cell", "pip install openpyxl")
    MergedCell = merged_cell_module.MergedCell

    for col_index in range(1, ws.max_column + 1):
        source_cell = ws.cell(source_row, col_index)
        target_cell = ws.cell(target_row, col_index)
        if isinstance(source_cell, MergedCell):
            continue
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.protection = copy(source_cell.protection)

    source_dim = ws.row_dimensions[source_row]
    target_dim = ws.row_dimensions[target_row]
    target_dim.height = source_dim.height
    target_dim.hidden = source_dim.hidden
    target_dim.outlineLevel = source_dim.outlineLevel
    target_dim.collapsed = source_dim.collapsed


def resize_table_body(ws, header_row: int, start_col: int, required_rows: int) -> int:
    """מגדיל/מקטין את הגוף של הטבלה תוך שימור עיצוב."""
    body_start = header_row + 1
    body_end = detect_table_body_end(ws, header_row, start_col, width=6)
    current_rows = max(body_end - body_start + 1, 0)

    if required_rows > current_rows:
        rows_to_add = required_rows - current_rows
        insert_at = body_end + 1
        style_row = body_end
        ws.insert_rows(insert_at, amount=rows_to_add)
        for row_index in range(insert_at, insert_at + rows_to_add):
            copy_row_format(ws, style_row, row_index)
        return body_start

    if required_rows < current_rows:
        delete_from = body_start + required_rows
        rows_to_delete = current_rows - required_rows
        ws.delete_rows(delete_from, rows_to_delete)

    return body_start


def clear_table_values(ws, header_row: int, start_col: int, from_row: int, to_row: int) -> None:
    """מנקה רק את הערכים באזור הטבלה, בלי לפגוע בעיצוב."""
    body_start = header_row + 1
    for row_index in range(max(from_row, body_start), max(to_row, body_start - 1) + 1):
        for col_index in range(start_col, start_col + 6):
            ws.cell(row=row_index, column=col_index).value = None


def apply_evening_divider_row(ws, header_row: int, start_col: int, target_row: int) -> None:
    """צובע שורת הפרדה למשמרת ערב לפי עיצוב הכותרת, בלי לגעת בשאר הגיליון."""
    for offset in range(6):
        source_cell = ws.cell(row=header_row, column=start_col + offset)
        target_cell = ws.cell(row=target_row, column=start_col + offset)
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.protection = copy(source_cell.protection)
        target_cell.value = None

    ws.row_dimensions[target_row].height = ws.row_dimensions[header_row].height
    ws.cell(row=target_row, column=start_col).value = "משמרת ערב"


def write_day_records(ws, records: list[ShiftRecord]) -> None:
    """כותב את כל הרשומות של יום מסוים לגיליון המתאים."""
    header_row, start_col = locate_target_table(ws)
    display_rows = build_display_rows(records)
    body_start = resize_table_body(ws, header_row, start_col, len(display_rows))
    if not display_rows:
        return
    clear_table_values(ws, header_row, start_col, body_start, body_start + max(len(display_rows) - 1, 0))

    for offset, display_row in enumerate(display_rows):
        row_index = body_start + offset
        if display_row.row_type == "divider":
            apply_evening_divider_row(ws, header_row, start_col, row_index)
            continue

        record = display_row.record
        if record is None:
            continue
        ws.cell(row=row_index, column=start_col + 0).value = record.team
        ws.cell(row=row_index, column=start_col + 1).value = record.name
        ws.cell(row=row_index, column=start_col + 2).value = record.shift_text
        ws.cell(row=row_index, column=start_col + 3).value = record.duration_hours
        ws.cell(row=row_index, column=start_col + 4).value = record.break_text
        ws.cell(row=row_index, column=start_col + 5).value = record.special_status


def write_day_records_fallback(ws, records: list[ShiftRecord]) -> None:
    """כתיבה ישירה לטבלה בלי מחיקת שורות, למקרה שהתבנית מתנהגת בצורה חריגה."""
    header_row, start_col = locate_target_table(ws)
    body_start = header_row + 1
    body_end = detect_table_body_end(ws, header_row, start_col, width=6)
    current_rows = max(body_end - body_start + 1, 0)
    display_rows = build_display_rows(records)

    if not display_rows:
        clear_table_values(ws, header_row, start_col, body_start, body_end)
        return

    if display_rows and len(display_rows) > current_rows:
        rows_to_add = len(display_rows) - current_rows
        insert_at = body_end + 1
        style_row = body_end
        ws.insert_rows(insert_at, amount=rows_to_add)
        for row_index in range(insert_at, insert_at + rows_to_add):
            copy_row_format(ws, style_row, row_index)
        body_end += rows_to_add

    clear_to_row = max(body_end, body_start + len(display_rows) - 1)
    clear_table_values(ws, header_row, start_col, body_start, clear_to_row)

    for offset, display_row in enumerate(display_rows):
        row_index = body_start + offset
        if display_row.row_type == "divider":
            apply_evening_divider_row(ws, header_row, start_col, row_index)
            continue

        record = display_row.record
        if record is None:
            continue
        ws.cell(row=row_index, column=start_col + 0).value = record.team
        ws.cell(row=row_index, column=start_col + 1).value = record.name
        ws.cell(row=row_index, column=start_col + 2).value = record.shift_text
        ws.cell(row=row_index, column=start_col + 3).value = record.duration_hours
        ws.cell(row=row_index, column=start_col + 4).value = record.break_text
        ws.cell(row=row_index, column=start_col + 5).value = record.special_status


def summarize_output_workbook(output_path: Path) -> dict[str, int]:
    """קורא את קובץ הפלט וסופר כמה שורות מולאו בפועל בכל גיליון."""
    openpyxl = import_or_raise("openpyxl", "pip install openpyxl")
    load_workbook = openpyxl.load_workbook

    workbook = load_workbook(output_path, data_only=False)
    resolved_sheets = resolve_day_worksheets(workbook)
    summary: dict[str, int] = {}
    for sheet_name in ["יום א", "יום ב", "יום ג", "יום ד", "יום ה", "יום ו"]:
        ws = resolved_sheets.get(sheet_name)
        if ws is None:
            summary[sheet_name] = 0
            continue

        header_row, start_col = locate_target_table(ws)
        body_end = detect_table_body_end(ws, header_row, start_col, width=6)
        count = 0
        for row_index in range(header_row + 1, body_end + 1):
            name = clean_text(ws.cell(row=row_index, column=start_col + 1).value)
            shift = clean_text(ws.cell(row=row_index, column=start_col + 2).value)
            if name and shift:
                count += 1
        summary[sheet_name] = count
    return summary


def ensure_output_contains_data(output_path: Path, grouped_records: dict[str, list[ShiftRecord]]) -> None:
    """מוודא שקובץ הפלט אכן מכיל את הנתונים שנכתבו."""
    actual_summary = summarize_output_workbook(output_path)
    expected_summary = {day: len(grouped_records.get(day, [])) for day in ["יום א", "יום ב", "יום ג", "יום ד", "יום ה", "יום ו"]}

    if actual_summary == expected_summary:
        return

    openpyxl = import_or_raise("openpyxl", "pip install openpyxl")
    load_workbook = openpyxl.load_workbook
    workbook = load_workbook(output_path)
    resolved_sheets = resolve_day_worksheets(workbook)
    for sheet_name in ["יום א", "יום ב", "יום ג", "יום ד", "יום ה", "יום ו"]:
        ws = resolved_sheets.get(sheet_name)
        if ws is None:
            continue
        write_day_records_fallback(ws, grouped_records.get(sheet_name, []))
    workbook.save(output_path)

    verified_summary = summarize_output_workbook(output_path)
    if verified_summary != expected_summary:
        raise RuntimeError(
            f"קובץ הפלט נשמר, אבל לא מולא כמו שצריך. ציפיתי ל-{expected_summary}, קיבלתי {verified_summary}."
        )


def fill_template(template_path: Path, output_path: Path, grouped_records: dict[str, list[ShiftRecord]]) -> None:
    """פותח את התבנית, ממלא את הגיליונות ושומר קובץ חדש."""
    openpyxl = import_or_raise("openpyxl", "pip install openpyxl")
    load_workbook = openpyxl.load_workbook

    workbook = load_workbook(template_path)
    resolved_sheets = resolve_day_worksheets(workbook)

    target_sheets = ["יום א", "יום ב", "יום ג", "יום ד", "יום ה", "יום ו"]
    if not resolved_sheets:
        raise RuntimeError(
            "לא נמצאו גיליונות יומיים תואמים בתבנית. "
            f"שמות הגיליונות שנמצאו: {workbook.sheetnames}"
        )

    for sheet_name in target_sheets:
        ws = resolved_sheets.get(sheet_name)
        if ws is None:
            continue
        write_day_records(ws, grouped_records.get(sheet_name, []))

    workbook.save(output_path)
    ensure_output_contains_data(output_path, grouped_records)


def process_attendance_files(
    source_path: Path,
    template_path: Path,
    output_path: Path,
) -> dict[str, list[ShiftRecord]]:
    """מריץ את כל תהליך הקריאה, המיון והמילוי ומחזיר גם סיכום יומי."""
    records = load_source_records(source_path)
    if not records:
        raise RuntimeError("לא נמצאו משמרות חוקיות בקובץ המקור.")

    grouped_records = sort_records(records)
    fill_template(template_path, output_path, grouped_records)
    return grouped_records


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="מילוי קובץ נוכחות מתוך סידור עבודה.")
    parser.add_argument("--source", required=True, help="נתיב לקובץ המקור.")
    parser.add_argument("--template", required=True, help="נתיב לתבנית היעד.")
    parser.add_argument("--output", required=True, help="נתיב לקובץ היעד החדש.")
    return parser


def main() -> int:
    args = build_argument_parser().parse_args()
    source_path = Path(args.source).expanduser().resolve()
    template_path = Path(args.template).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if not source_path.exists():
        raise FileNotFoundError(f"קובץ המקור לא נמצא: {source_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"תבנית היעד לא נמצאה: {template_path}")

    grouped_records = process_attendance_files(source_path, template_path, output_path)

    print(f"נוצר קובץ חדש: {output_path}")
    for sheet_name in ("יום א", "יום ב", "יום ג", "יום ד", "יום ה", "יום ו"):
        count = len(grouped_records.get(sheet_name, []))
        print(f"{sheet_name}: {count} רשומות")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover - זה בלוק CLI שימושי לשגיאות ריצה.
        print(f"שגיאה: {exc}", file=sys.stderr)
        raise SystemExit(1)

#!/usr/bin/env python3
"""אפליקציית Streamlit להכנת קובץ נוכחות מתוך סידור עבודה."""

from __future__ import annotations

import tempfile
from pathlib import Path

from fill_attendance import import_or_raise, process_attendance_files


st = import_or_raise("streamlit", "pip install streamlit")


APP_VERSION = "2026-03-17-divider-detection-v5"
DAY_ORDER = ("יום א", "יום ב", "יום ג", "יום ד", "יום ה", "יום ו")
DOWNLOAD_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def uploaded_files_signature(source_file, template_file) -> tuple[object, ...]:
    """מזהה מצב נוכחי של קבצים שהועלו כדי למנוע הורדה של תוצאה ישנה."""
    return (
        APP_VERSION,
        getattr(source_file, "name", None),
        getattr(source_file, "size", None),
        getattr(template_file, "name", None),
        getattr(template_file, "size", None),
    )


def clear_result_state() -> None:
    """מנקה תוצאה קודמת כאשר אחד הקבצים התחלף."""
    for key in ("prepared_file_bytes", "prepared_file_name", "prepared_summary", "prepared_signature"):
        st.session_state.pop(key, None)


def build_output_filename(template_name: str) -> str:
    """יוצר שם קובץ הגיוני להורדה."""
    template_path = Path(template_name)
    suffix = template_path.suffix or ".xlsx"
    return f"{template_path.stem} - מוכן{suffix}"


def summary_counts(grouped_records: dict[str, list[object]]) -> dict[str, int]:
    """ממיר רשומות יומיות לספירת נציגים קצרה להצגה."""
    return {day: len(grouped_records.get(day, [])) for day in DAY_ORDER}


def summary_text(summary: dict[str, int]) -> str:
    """מייצר משפט קצר להודעת הצלחה."""
    return " | ".join(f"{day}: {count}" for day, count in summary.items())


def process_uploaded_files(source_file, template_file) -> tuple[bytes, str, dict[str, int]]:
    """שומר קבצים זמנית, מריץ את הלוגיקה ומחזיר קובץ מוכן להורדה."""
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        source_path = temp_path / source_file.name
        template_path = temp_path / template_file.name
        output_name = build_output_filename(template_file.name)
        output_path = temp_path / output_name

        source_path.write_bytes(source_file.getvalue())
        template_path.write_bytes(template_file.getvalue())

        grouped_records = process_attendance_files(source_path, template_path, output_path)
        return output_path.read_bytes(), output_name, summary_counts(grouped_records)


def render_summary_metrics(summary: dict[str, int]) -> None:
    """מציג סיכום יומי בצורה נקייה."""
    columns = st.columns(3)
    for index, day in enumerate(DAY_ORDER):
        with columns[index % 3]:
            st.metric(day, summary[day])


def main() -> None:
    st.set_page_config(
        page_title="מחולל קובץ נוכחות",
        page_icon="📄",
        layout="centered",
    )

    st.title("מחולל קובץ נוכחות")
    st.caption("מעלה קובץ סידור עבודה ותבנית נוכחות, ומחזיר קובץ מוכן להורדה תוך שמירה על העיצוב.")

    st.markdown(
        """
        **איך משתמשים**

        1. העלה את קובץ סידור העבודה.
        2. העלה את קובץ תבנית הנוכחות.
        3. לחץ על `הכן קובץ`.
        4. הורד את קובץ האקסל המוכן.
        """
    )

    st.info("העיבוד מתבצע על קבצים זמניים בלבד, והקובץ המוכן נוצר לאחר חישוב משמרות, הפסקות והתאמת שורות בתבנית.")

    upload_col1, upload_col2 = st.columns(2)
    with upload_col1:
        source_file = st.file_uploader(
            "קובץ סידור עבודה",
            type=["xls", "xlsx", "xlsm"],
            help="קובץ המקור שממנו נמשכים המשמרות והסטטוסים.",
        )
    with upload_col2:
        template_file = st.file_uploader(
            "קובץ תבנית נוכחות",
            type=["xlsx", "xlsm"],
            help="קובץ היעד שאליו יודבקו הנתונים.",
        )

    current_signature = uploaded_files_signature(source_file, template_file)
    previous_signature = st.session_state.get("prepared_signature")
    if previous_signature and previous_signature != current_signature:
        clear_result_state()

    if source_file is not None:
        st.caption(f"קובץ מקור: `{source_file.name}`")
    if template_file is not None:
        st.caption(f"קובץ תבנית: `{template_file.name}`")

    prepare_clicked = st.button(
        "הכן קובץ",
        type="primary",
        use_container_width=True,
        disabled=not (source_file and template_file),
    )

    if prepare_clicked:
        progress = st.progress(0, text="מתחיל עיבוד...")
        try:
            with st.spinner("מעבד את קבצי האקסל..."):
                progress.progress(20, text="שומר קבצים זמניים...")
                output_bytes, output_name, summary = process_uploaded_files(source_file, template_file)
                progress.progress(85, text="מכין את הקובץ להורדה...")

            st.session_state["prepared_file_bytes"] = output_bytes
            st.session_state["prepared_file_name"] = output_name
            st.session_state["prepared_summary"] = summary
            st.session_state["prepared_signature"] = current_signature
            progress.progress(100, text="הקובץ מוכן.")
        except Exception as exc:
            progress.empty()
            clear_result_state()
            st.error(f"שגיאה בעיבוד הקבצים: {exc}")
        else:
            progress.empty()

    prepared_bytes = st.session_state.get("prepared_file_bytes")
    prepared_name = st.session_state.get("prepared_file_name")
    prepared_summary = st.session_state.get("prepared_summary")

    if prepared_bytes and prepared_name and prepared_summary:
        st.success(f"הקובץ הוכן בהצלחה. {summary_text(prepared_summary)}")
        render_summary_metrics(prepared_summary)
        st.download_button(
            "הורד את הקובץ המוכן",
            data=prepared_bytes,
            file_name=prepared_name,
            mime=DOWNLOAD_MIME,
            use_container_width=True,
        )


if __name__ == "__main__":
    main()

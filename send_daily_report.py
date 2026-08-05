from __future__ import annotations

import os
import smtplib
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from supabase import create_client


TIMEZONE = ZoneInfo("Europe/Prague")

ACTIVITY_ORDER = [
    "Aperam",
    "Personna",
    "SSI",
    "Zanini",
    "Rebound",
]


def get_secret(name: str) -> str:
    value = os.getenv(name)

    if not value:
        raise RuntimeError(
            f"V GitHub Secrets chybí hodnota: {name}"
        )

    return value


SUPABASE_URL = get_secret("SUPABASE_URL")
SUPABASE_KEY = get_secret("SUPABASE_KEY")
GMAIL_ADDRESS = get_secret("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = get_secret("GMAIL_APP_PASSWORD")
REPORT_RECIPIENT = get_secret("REPORT_RECIPIENT")


def parse_datetime(value: str) -> datetime:
    return datetime.fromisoformat(
        value.replace("Z", "+00:00")
    )


def to_local_time(value: str) -> datetime:
    return parse_datetime(value).astimezone(TIMEZONE)


def format_duration(
    seconds: int | float | None,
) -> str:
    total_seconds = max(
        0,
        int(seconds or 0),
    )

    hours, remainder = divmod(
        total_seconds,
        3600,
    )

    minutes, seconds = divmod(
        remainder,
        60,
    )

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def load_records() -> list[dict]:
    database = create_client(
        SUPABASE_URL,
        SUPABASE_KEY,
    )

    period_end = datetime.now(timezone.utc)
    period_start = period_end - timedelta(hours=24)

    response = (
        database.table("activity_log")
        .select("*")
        .gte("start_time", period_start.isoformat())
        .lt("start_time", period_end.isoformat())
        .order("start_time", desc=False)
        .execute()
    )

    return response.data or []


def get_duration_seconds(
    record: dict,
    current_utc: datetime,
) -> int:
    if record.get("duration_seconds") is not None:
        return max(
            0,
            int(record["duration_seconds"]),
        )

    return max(
        0,
        int(
            (
                current_utc
                - parse_datetime(record["start_time"])
            ).total_seconds()
        ),
    )


def calculate_summary(
    records: list[dict],
) -> list[dict]:
    totals = {
        activity: 0
        for activity in ACTIVITY_ORDER
    }

    current_utc = datetime.now(timezone.utc)

    for record in records:
        activity = record["activity"]

        duration = get_duration_seconds(
            record,
            current_utc,
        )

        totals[activity] = (
            totals.get(activity, 0)
            + duration
        )

    total_seconds = sum(totals.values())

    summary = []

    for activity in ACTIVITY_ORDER:
        seconds = totals.get(activity, 0)

        if total_seconds > 0:
            percentage = (
                seconds / total_seconds
            )
        else:
            percentage = 0

        summary.append(
            {
                "Činnost": activity,
                "Trvání": format_duration(seconds),
                "Minuty": round(seconds / 60, 2),
                "Podíl": percentage,
            }
        )

    return summary


def create_excel(
    records: list[dict],
) -> tuple[bytes, str]:
    detail_rows = []
    current_utc = datetime.now(timezone.utc)

    for record in records:
        start_local = to_local_time(
            record["start_time"]
        )

        end_value = record.get("end_time")

        end_local = (
            to_local_time(end_value)
            if end_value
            else None
        )

        duration_seconds = get_duration_seconds(
            record,
            current_utc,
        )

        detail_rows.append(
            {
                "Datum": start_local.strftime(
                    "%d.%m.%Y"
                ),
                "ID": record["employee_id"],
                "Jméno": record["employee_name"],
                "Činnost": record["activity"],
                "Start": start_local.strftime(
                    "%H:%M:%S"
                ),
                "Konec": (
                    end_local.strftime("%H:%M:%S")
                    if end_local
                    else ""
                ),
                "Trvání": format_duration(
                    duration_seconds
                ),
                "Trvání v minutách": round(
                    duration_seconds / 60,
                    2,
                ),
                "Stav": (
                    "Dokončeno"
                    if end_value
                    else "Probíhá"
                ),
            }
        )

    detail_columns = [
        "Datum",
        "ID",
        "Jméno",
        "Činnost",
        "Start",
        "Konec",
        "Trvání",
        "Trvání v minutách",
        "Stav",
    ]

    detail_dataframe = pd.DataFrame(
        detail_rows,
        columns=detail_columns,
    )

    summary_rows = calculate_summary(records)

    summary_dataframe = pd.DataFrame(
        summary_rows,
        columns=[
            "Činnost",
            "Trvání",
            "Minuty",
            "Podíl",
        ],
    )

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        detail_dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Detail",
        )

        summary_dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Souhrn",
        )

        detail_sheet = writer.sheets["Detail"]
        summary_sheet = writer.sheets["Souhrn"]

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="00529B",
        )

        header_font = Font(
            color="FFFFFF",
            bold=True,
        )

        for worksheet in [
            detail_sheet,
            summary_sheet,
        ]:
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                )

            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

        detail_widths = {
            "A": 13,
            "B": 11,
            "C": 28,
            "D": 17,
            "E": 12,
            "F": 12,
            "G": 14,
            "H": 20,
            "I": 14,
        }

        for column, width in detail_widths.items():
            detail_sheet.column_dimensions[
                column
            ].width = width

        summary_sheet.column_dimensions["A"].width = 18
        summary_sheet.column_dimensions["B"].width = 16
        summary_sheet.column_dimensions["C"].width = 14
        summary_sheet.column_dimensions["D"].width = 14

        for row_number in range(
            2,
            len(summary_rows) + 2,
        ):
            summary_sheet[
                f"D{row_number}"
            ].number_format = "0.0%"

        chart = PieChart()

        chart.title = (
            "Rozdělení času podle činností"
        )

        labels = Reference(
            summary_sheet,
            min_col=1,
            min_row=2,
            max_row=len(summary_rows) + 1,
        )

        data = Reference(
            summary_sheet,
            min_col=3,
            min_row=1,
            max_row=len(summary_rows) + 1,
        )

        chart.add_data(
            data,
            titles_from_data=True,
        )

        chart.set_categories(labels)

        chart.height = 10
        chart.width = 15

        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showLeaderLines = True

        summary_sheet.add_chart(
            chart,
            "F2",
        )

    current_local = datetime.now(TIMEZONE)

    filename = (
        "prehled_cinnosti_"
        + current_local.strftime("%Y-%m-%d")
        + ".xlsx"
    )

    return output.getvalue(), filename


def create_email_summary(
    records: list[dict],
) -> str:
    summary_rows = calculate_summary(records)

    lines = [
        "Souhrn podle činností:",
        "",
    ]

    for row in summary_rows:
        percentage = row["Podíl"] * 100

        lines.append(
            f"{row['Činnost']}: "
            f"{percentage:.1f} % "
            f"({row['Trvání']})"
        )

    return "\n".join(lines)


def send_email(
    excel_data: bytes,
    filename: str,
    records: list[dict],
) -> None:
    now_local = datetime.now(TIMEZONE)
    period_start = now_local - timedelta(hours=24)

    message = EmailMessage()

    message["From"] = GMAIL_ADDRESS
    message["To"] = REPORT_RECIPIENT
    message["Subject"] = (
        "Denní přehled činností UWH – "
        + now_local.strftime("%d.%m.%Y")
    )

    summary = create_email_summary(records)

    message.set_content(
        f"""Dobrý den,

v příloze zasílám automatický přehled činností za posledních 24 hodin.

Období:
{period_start.strftime("%d.%m.%Y %H:%M")}
až
{now_local.strftime("%d.%m.%Y %H:%M")}

Počet záznamů: {len(records)}

{summary}

Excel obsahuje:
- list Detail se všemi záznamy,
- list Souhrn s procenty a grafem.

Tento e-mail byl vytvořen automaticky.
"""
    )

    message.add_attachment(
        excel_data,
        maintype="application",
        subtype=(
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        filename=filename,
    )

    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465,
        timeout=30,
    ) as smtp:
        smtp.login(
            GMAIL_ADDRESS,
            GMAIL_APP_PASSWORD,
        )

        smtp.send_message(message)


def main() -> None:
    records = load_records()

    excel_data, filename = create_excel(
        records
    )

    send_email(
        excel_data=excel_data,
        filename=filename,
        records=records,
    )

    print(
        f"Report byl úspěšně odeslán na "
        f"{REPORT_RECIPIENT}. "
        f"Počet záznamů: {len(records)}"
    )


if __name__ == "__main__":
    main()

from __future__ import annotations

import os
import smtplib
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from supabase import create_client


TIMEZONE = ZoneInfo("Europe/Prague")


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


def format_duration(seconds: int | float | None) -> str:
    total_seconds = max(0, int(seconds or 0))

    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

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


def create_excel(
    records: list[dict],
) -> tuple[bytes, str]:
    export_rows = []
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

        if record.get("duration_seconds") is not None:
            duration_seconds = int(
                record["duration_seconds"]
            )
        else:
            duration_seconds = int(
                (
                    current_utc
                    - parse_datetime(record["start_time"])
                ).total_seconds()
            )

        export_rows.append(
            {
                "Datum": start_local.strftime("%d.%m.%Y"),
                "ID": record["employee_id"],
                "Jméno": record["employee_name"],
                "Činnost": record["activity"],
                "Start": start_local.strftime("%H:%M:%S"),
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

    columns = [
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

    dataframe = pd.DataFrame(
        export_rows,
        columns=columns,
    )

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Denní přehled",
        )

        worksheet = writer.sheets["Denní přehled"]

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="00529B",
        )

        header_font = Font(
            color="FFFFFF",
            bold=True,
        )

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        widths = {
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

        for column, width in widths.items():
            worksheet.column_dimensions[column].width = width

    current_local = datetime.now(TIMEZONE)

    filename = (
        "prehled_cinnosti_"
        + current_local.strftime("%Y-%m-%d")
        + ".xlsx"
    )

    return output.getvalue(), filename


def create_summary(
    records: list[dict],
) -> str:
    activity_totals: dict[str, int] = {}
    current_utc = datetime.now(timezone.utc)

    for record in records:
        activity = record["activity"]

        if record.get("duration_seconds") is not None:
            duration = int(record["duration_seconds"])
        else:
            duration = int(
                (
                    current_utc
                    - parse_datetime(record["start_time"])
                ).total_seconds()
            )

        activity_totals[activity] = (
            activity_totals.get(activity, 0)
            + duration
        )

    if not activity_totals:
        return "Za posledních 24 hodin nebyly nalezeny žádné záznamy."

    lines = ["Souhrn podle činností:"]

    for activity, seconds in sorted(
        activity_totals.items()
    ):
        lines.append(
            f"- {activity}: {format_duration(seconds)}"
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

    summary = create_summary(records)

    message.set_content(
        f"""Dobrý den,

v příloze zasílám automatický přehled činností za posledních 24 hodin.

Období:
{period_start.strftime("%d.%m.%Y %H:%M")}
až
{now_local.strftime("%d.%m.%Y %H:%M")}

Počet záznamů: {len(records)}

{summary}

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

    excel_data, filename = create_excel(records)

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

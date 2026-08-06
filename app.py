from __future__ import annotations

import os
import smtplib
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from io import BytesIO
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from supabase import create_client


# ============================================================
# NASTAVENÍ
# ============================================================

APP_TIMEZONE = ZoneInfo("Europe/Prague")

ACTIVITIES = [
    "Aperam",
    "Personna",
    "SSI",
    "Zanini",
    "Rebound",
]

MACHINES = [
    "F33",
    "F36",
    "F45",
    "F86",
    "F87",
    "F88",
    "F117",
    "F140",
    "FS04",
    "FS07 Lion",
    "Jung1",
    "Jung2",
]

YUSEN_BLUE = "00529B"
YUSEN_DARK_BLUE = "003B70"
YUSEN_ORANGE = "F58220"
WHITE = "FFFFFF"
LIGHT_BLUE = "DDEBF7"
LIGHT_ORANGE = "FCE4D6"
LIGHT_GREY = "E7E6E6"
DARK_TEXT = "1F2937"

THIN_GREY_SIDE = Side(
    style="thin",
    color="B7C4CE",
)

TABLE_BORDER = Border(
    left=THIN_GREY_SIDE,
    right=THIN_GREY_SIDE,
    top=THIN_GREY_SIDE,
    bottom=THIN_GREY_SIDE,
)


# ============================================================
# GITHUB SECRETS
# ============================================================

def get_required_secret(name: str) -> str:
    value = os.getenv(name)

    if not value:
        raise RuntimeError(
            f"V GitHub Secrets chybí hodnota: {name}"
        )

    return value


SUPABASE_URL = get_required_secret("SUPABASE_URL")
SUPABASE_KEY = get_required_secret("SUPABASE_KEY")
GMAIL_ADDRESS = get_required_secret("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = get_required_secret(
    "GMAIL_APP_PASSWORD"
)
REPORT_RECIPIENT = get_required_secret(
    "REPORT_RECIPIENT"
)


# ============================================================
# DATUM A ČAS
# ============================================================

def parse_datetime(value: str) -> datetime:
    return datetime.fromisoformat(
        value.replace("Z", "+00:00")
    )


def to_local_datetime(value: str) -> datetime:
    return parse_datetime(value).astimezone(
        APP_TIMEZONE
    )


def format_duration(
    seconds: int | float | None,
) -> str:
    total_seconds = max(0, int(seconds or 0))

    hours, remainder = divmod(
        total_seconds,
        3600,
    )

    minutes, seconds = divmod(
        remainder,
        60,
    )

    return (
        f"{hours:02d}:"
        f"{minutes:02d}:"
        f"{seconds:02d}"
    )


def get_record_duration(
    record: dict,
    current_utc: datetime,
) -> int:
    saved_duration = record.get(
        "duration_seconds"
    )

    if saved_duration is not None:
        return max(
            0,
            int(saved_duration),
        )

    start_time = parse_datetime(
        record["start_time"]
    )

    return max(
        0,
        int(
            (
                current_utc - start_time
            ).total_seconds()
        ),
    )


# ============================================================
# NAČTENÍ DAT
# ============================================================

def load_records() -> tuple[
    list[dict],
    datetime,
    datetime,
]:
    database = create_client(
        SUPABASE_URL,
        SUPABASE_KEY,
    )

    period_end_utc = datetime.now(
        timezone.utc
    )

    period_start_utc = (
        period_end_utc
        - timedelta(hours=24)
    )

    response = (
        database.table("activity_log")
        .select("*")
        .gte(
            "start_time",
            period_start_utc.isoformat(),
        )
        .lt(
            "start_time",
            period_end_utc.isoformat(),
        )
        .order(
            "start_time",
            desc=False,
        )
        .execute()
    )

    return (
        response.data or [],
        period_start_utc,
        period_end_utc,
    )


# ============================================================
# VÝPOČTY
# ============================================================

def calculate_activity_totals(
    records: list[dict],
    current_utc: datetime,
) -> dict[str, int]:
    totals = {
        activity: 0
        for activity in ACTIVITIES
    }

    for record in records:
        activity = record.get(
            "activity",
            "Neznámá",
        )

        duration = get_record_duration(
            record,
            current_utc,
        )

        totals[activity] = (
            totals.get(activity, 0)
            + duration
        )

    return totals


def get_machine_name(record: dict) -> str:
    machine = str(record.get("machine") or "").strip()
    return machine if machine else "Neuveden"


def calculate_machine_activity_totals(
    records: list[dict],
    current_utc: datetime,
) -> dict[tuple[str, str], int]:
    totals: dict[tuple[str, str], int] = {}

    for record in records:
        machine = get_machine_name(record)
        activity = str(record.get("activity") or "Neznámá")
        duration = get_record_duration(record, current_utc)
        key = (machine, activity)
        totals[key] = totals.get(key, 0) + duration

    return totals


def calculate_worker_totals(
    records: list[dict],
    current_utc: datetime,
) -> dict[tuple[str, str], int]:
    totals: dict[
        tuple[str, str],
        int,
    ] = {}

    for record in records:
        employee_id = str(
            record.get("employee_id", "")
        )

        employee_name = str(
            record.get("employee_name", "")
        )

        key = (
            employee_id,
            employee_name,
        )

        duration = get_record_duration(
            record,
            current_utc,
        )

        totals[key] = (
            totals.get(key, 0)
            + duration
        )

    return totals


def find_longest_record(
    records: list[dict],
    current_utc: datetime,
) -> tuple[dict | None, int]:
    longest_record = None
    longest_duration = 0

    for record in records:
        duration = get_record_duration(
            record,
            current_utc,
        )

        if duration > longest_duration:
            longest_record = record
            longest_duration = duration

    return (
        longest_record,
        longest_duration,
    )


# ============================================================
# FORMÁTOVÁNÍ EXCELU
# ============================================================

def style_header_row(
    worksheet,
    row_number: int = 1,
) -> None:
    fill = PatternFill(
        fill_type="solid",
        fgColor=YUSEN_BLUE,
    )

    font = Font(
        color=WHITE,
        bold=True,
    )

    for cell in worksheet[row_number]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )
        cell.border = TABLE_BORDER


def style_table_area(
    worksheet,
    min_row: int,
    max_row: int,
    min_column: int,
    max_column: int,
) -> None:
    for row in worksheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_column,
        max_col=max_column,
    ):
        for cell in row:
            cell.border = TABLE_BORDER
            cell.alignment = Alignment(
                vertical="center",
            )


def set_column_widths(
    worksheet,
    widths: dict[int, float],
) -> None:
    for column_number, width in widths.items():
        column_letter = get_column_letter(
            column_number
        )

        worksheet.column_dimensions[
            column_letter
        ].width = width


def add_section_title(
    worksheet,
    cell_reference: str,
    text: str,
) -> None:
    cell = worksheet[cell_reference]

    cell.value = text

    cell.fill = PatternFill(
        fill_type="solid",
        fgColor=YUSEN_ORANGE,
    )

    cell.font = Font(
        color=WHITE,
        bold=True,
        size=13,
    )

    cell.alignment = Alignment(
        horizontal="left",
        vertical="center",
    )


# ============================================================
# VYTVOŘENÍ EXCELU
# ============================================================

def create_excel(
    records: list[dict],
    period_start_utc: datetime,
    period_end_utc: datetime,
) -> tuple[bytes, str]:
    workbook = Workbook()

    detail_sheet = workbook.active
    detail_sheet.title = "Detail"

    summary_sheet = workbook.create_sheet(
        "Souhrn"
    )

    worker_sheet = workbook.create_sheet(
        "Pracovníci"
    )

    # ========================================================
    # LIST DETAIL
    # ========================================================

    detail_headers = [
        "Datum",
        "ID",
        "Jméno",
        "Stroj",
        "Činnost",
        "Start",
        "Konec",
        "Trvání",
        "Trvání v minutách",
        "Stav",
    ]

    detail_sheet.append(
        detail_headers
    )

    for record in records:
        start_local = to_local_datetime(
            record["start_time"]
        )

        end_value = record.get(
            "end_time"
        )

        end_local = (
            to_local_datetime(end_value)
            if end_value
            else None
        )

        duration_seconds = get_record_duration(
            record,
            period_end_utc,
        )

        detail_sheet.append(
            [
                start_local.strftime(
                    "%d.%m.%Y"
                ),
                record.get(
                    "employee_id",
                    "",
                ),
                record.get(
                    "employee_name",
                    "",
                ),
                get_machine_name(record),
                record.get(
                    "activity",
                    "",
                ),
                start_local.strftime(
                    "%H:%M:%S"
                ),
                (
                    end_local.strftime(
                        "%H:%M:%S"
                    )
                    if end_local
                    else ""
                ),
                format_duration(
                    duration_seconds
                ),
                round(
                    duration_seconds / 60,
                    2,
                ),
                (
                    "Dokončeno"
                    if end_value
                    else "Probíhá"
                ),
            ]
        )

    style_header_row(
        detail_sheet
    )

    if detail_sheet.max_row >= 2:
        style_table_area(
            detail_sheet,
            min_row=2,
            max_row=detail_sheet.max_row,
            min_column=1,
            max_column=10,
        )

    detail_sheet.freeze_panes = "A2"
    detail_sheet.auto_filter.ref = (
        detail_sheet.dimensions
    )

    set_column_widths(
        detail_sheet,
        {
            1: 14,
            2: 12,
            3: 28,
            4: 16,
            5: 18,
            6: 12,
            7: 12,
            8: 14,
            9: 21,
            10: 14,
        },
    )

    # ========================================================
    # LIST SOUHRN
    # ========================================================

    period_start_local = (
        period_start_utc.astimezone(
            APP_TIMEZONE
        )
    )

    period_end_local = (
        period_end_utc.astimezone(
            APP_TIMEZONE
        )
    )

    summary_sheet["A1"] = (
        "DENNÍ PŘEHLED ČINNOSTÍ UWH"
    )

    summary_sheet["A1"].font = Font(
        bold=True,
        size=18,
        color=YUSEN_DARK_BLUE,
    )

    summary_sheet["A2"] = (
        "Období:"
    )

    summary_sheet["B2"] = (
        f"{period_start_local.strftime('%d.%m.%Y %H:%M')} "
        f"až "
        f"{period_end_local.strftime('%d.%m.%Y %H:%M')}"
    )

    summary_sheet["A3"] = (
        "Počet záznamů:"
    )

    summary_sheet["B3"] = len(
        records
    )

    activity_totals = (
        calculate_activity_totals(
            records,
            period_end_utc,
        )
    )

    total_seconds = sum(
        activity_totals.values()
    )

    summary_sheet["A4"] = (
        "Celkový zaznamenaný čas:"
    )

    summary_sheet["B4"] = (
        format_duration(total_seconds)
    )

    add_section_title(
        summary_sheet,
        "A6",
        "Rozdělení času podle činností",
    )

    summary_sheet["A7"] = "Činnost"
    summary_sheet["B7"] = "Trvání"
    summary_sheet["C7"] = "Minuty"
    summary_sheet["D7"] = "Podíl"

    style_header_row(
        summary_sheet,
        row_number=7,
    )

    activity_rows = []

    for activity in ACTIVITIES:
        seconds = activity_totals.get(
            activity,
            0,
        )

        percentage = (
            seconds / total_seconds
            if total_seconds > 0
            else 0
        )

        activity_rows.append(
            (
                activity,
                seconds,
                percentage,
            )
        )

    for activity, seconds in (
        activity_totals.items()
    ):
        if activity in ACTIVITIES:
            continue

        percentage = (
            seconds / total_seconds
            if total_seconds > 0
            else 0
        )

        activity_rows.append(
            (
                activity,
                seconds,
                percentage,
            )
        )

    current_row = 8

    for (
        activity,
        seconds,
        percentage,
    ) in activity_rows:
        summary_sheet.append(
            [
                activity,
                format_duration(seconds),
                round(seconds / 60, 2),
                percentage,
            ]
        )

        summary_sheet[
            f"D{current_row}"
        ].number_format = "0.0%"

        current_row += 1

    last_activity_row = (
        current_row - 1
    )

    style_table_area(
        summary_sheet,
        min_row=8,
        max_row=last_activity_row,
        min_column=1,
        max_column=4,
    )

    total_row = (
        last_activity_row + 1
    )

    summary_sheet[
        f"A{total_row}"
    ] = "CELKEM"

    summary_sheet[
        f"B{total_row}"
    ] = format_duration(
        total_seconds
    )

    summary_sheet[
        f"C{total_row}"
    ] = round(
        total_seconds / 60,
        2,
    )

    summary_sheet[
        f"D{total_row}"
    ] = (
        1 if total_seconds > 0 else 0
    )

    summary_sheet[
        f"D{total_row}"
    ].number_format = "0.0%"

    for cell in summary_sheet[
        total_row
    ]:
        cell.font = Font(
            bold=True,
            color=YUSEN_DARK_BLUE,
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=LIGHT_BLUE,
        )

        cell.border = TABLE_BORDER

    # --------------------------------------------------------
    # NEJDELŠÍ JEDNOTLIVÁ ČINNOST
    # --------------------------------------------------------

    longest_record, longest_duration = (
        find_longest_record(
            records,
            period_end_utc,
        )
    )

    longest_title_row = (
        total_row + 3
    )

    add_section_title(
        summary_sheet,
        f"A{longest_title_row}",
        "Nejdelší jednotlivý záznam",
    )

    if longest_record:
        longest_start = to_local_datetime(
            longest_record["start_time"]
        )

        longest_end_value = (
            longest_record.get(
                "end_time"
            )
        )

        longest_end = (
            to_local_datetime(
                longest_end_value
            )
            if longest_end_value
            else None
        )

        summary_sheet[
            f"A{longest_title_row + 1}"
        ] = "Pracovník"

        summary_sheet[
            f"B{longest_title_row + 1}"
        ] = longest_record.get(
            "employee_name",
            "",
        )

        summary_sheet[
            f"A{longest_title_row + 2}"
        ] = "Stroj"

        summary_sheet[
            f"B{longest_title_row + 2}"
        ] = get_machine_name(longest_record)

        summary_sheet[
            f"A{longest_title_row + 3}"
        ] = "Činnost"

        summary_sheet[
            f"B{longest_title_row + 3}"
        ] = longest_record.get(
            "activity",
            "",
        )

        summary_sheet[
            f"A{longest_title_row + 4}"
        ] = "Trvání"

        summary_sheet[
            f"B{longest_title_row + 4}"
        ] = format_duration(
            longest_duration
        )

        summary_sheet[
            f"A{longest_title_row + 5}"
        ] = "Čas"

        summary_sheet[
            f"B{longest_title_row + 5}"
        ] = (
            f"{longest_start.strftime('%H:%M:%S')} – "
            + (
                longest_end.strftime(
                    "%H:%M:%S"
                )
                if longest_end
                else "stále probíhá"
            )
        )

        style_table_area(
            summary_sheet,
            min_row=longest_title_row + 1,
            max_row=longest_title_row + 5,
            min_column=1,
            max_column=2,
        )

    else:
        summary_sheet[
            f"A{longest_title_row + 1}"
        ] = (
            "Za dané období nejsou záznamy."
        )

    # --------------------------------------------------------
    # SOUHRN PODLE STROJE A ČINNOSTI
    # --------------------------------------------------------

    machine_activity_totals = calculate_machine_activity_totals(
        records,
        period_end_utc,
    )

    machine_title_row = longest_title_row + 8

    add_section_title(
        summary_sheet,
        f"A{machine_title_row}",
        "Rozdělení času podle stroje a činnosti",
    )

    header_row = machine_title_row + 1
    summary_sheet[f"A{header_row}"] = "Stroj"
    summary_sheet[f"B{header_row}"] = "Činnost"
    summary_sheet[f"C{header_row}"] = "Trvání"
    summary_sheet[f"D{header_row}"] = "Minuty"
    style_header_row(summary_sheet, row_number=header_row)

    machine_order = {name: index for index, name in enumerate(MACHINES)}
    activity_order = {name: index for index, name in enumerate(ACTIVITIES)}

    sorted_machine_rows = sorted(
        machine_activity_totals.items(),
        key=lambda item: (
            machine_order.get(item[0][0], len(MACHINES)),
            item[0][0],
            activity_order.get(item[0][1], len(ACTIVITIES)),
            item[0][1],
        ),
    )

    machine_data_start = header_row + 1

    if sorted_machine_rows:
        for (machine, activity), seconds in sorted_machine_rows:
            summary_sheet.append(
                [
                    machine,
                    activity,
                    format_duration(seconds),
                    round(seconds / 60, 2),
                ]
            )

        style_table_area(
            summary_sheet,
            min_row=machine_data_start,
            max_row=summary_sheet.max_row,
            min_column=1,
            max_column=4,
        )
    else:
        summary_sheet[f"A{machine_data_start}"] = (
            "Za dané období nejsou záznamy."
        )

    set_column_widths(
        summary_sheet,
        {
            1: 26,
            2: 26,
            3: 16,
            4: 14,
            5: 3,
            6: 18,
            7: 18,
            8: 18,
            9: 18,
        },
    )

    # --------------------------------------------------------
    # KOLÁČOVÝ GRAF
    # --------------------------------------------------------

    pie_chart = PieChart()

    pie_chart.title = (
        "Podíl činností z celkového času"
    )

    pie_chart.height = 10
    pie_chart.width = 15

    pie_labels = Reference(
        summary_sheet,
        min_col=1,
        min_row=8,
        max_row=last_activity_row,
    )

    pie_values = Reference(
        summary_sheet,
        min_col=3,
        min_row=7,
        max_row=last_activity_row,
    )

    pie_chart.add_data(
        pie_values,
        titles_from_data=True,
    )

    pie_chart.set_categories(
        pie_labels
    )

    pie_chart.legend.position = "r"

    pie_chart.dataLabels = (
        DataLabelList()
    )

    # Na grafu budou pouze procenta.
    pie_chart.dataLabels.showPercent = True
    pie_chart.dataLabels.showVal = False
    pie_chart.dataLabels.showCatName = False
    pie_chart.dataLabels.showSerName = False
    pie_chart.dataLabels.showLegendKey = False
    pie_chart.dataLabels.showLeaderLines = True

    summary_sheet.add_chart(
        pie_chart,
        "F2",
    )

    # --------------------------------------------------------
    # SLOUPCOVÝ GRAF
    # --------------------------------------------------------

    bar_chart = BarChart()

    bar_chart.type = "bar"
    bar_chart.style = 10

    bar_chart.title = (
        "Čas jednotlivých činností v minutách"
    )

    bar_chart.y_axis.title = (
        "Činnost"
    )

    bar_chart.x_axis.title = (
        "Minuty"
    )

    bar_chart.height = 8
    bar_chart.width = 15

    bar_values = Reference(
        summary_sheet,
        min_col=3,
        min_row=7,
        max_row=last_activity_row,
    )

    bar_labels = Reference(
        summary_sheet,
        min_col=1,
        min_row=8,
        max_row=last_activity_row,
    )

    bar_chart.add_data(
        bar_values,
        titles_from_data=True,
    )

    bar_chart.set_categories(
        bar_labels
    )

    bar_chart.legend = None

    summary_sheet.add_chart(
        bar_chart,
        "F22",
    )

    summary_sheet.freeze_panes = "A7"

    # ========================================================
    # LIST PRACOVNÍCI
    # ========================================================

    worker_totals = (
        calculate_worker_totals(
            records,
            period_end_utc,
        )
    )

    worker_sheet.append(
        [
            "ID",
            "Pracovník",
            "Celkový čas",
            "Minuty",
            "Podíl",
            "Počet záznamů",
        ]
    )

    style_header_row(
        worker_sheet
    )

    worker_record_counts: dict[
        tuple[str, str],
        int,
    ] = {}

    for record in records:
        key = (
            str(
                record.get(
                    "employee_id",
                    "",
                )
            ),
            str(
                record.get(
                    "employee_name",
                    "",
                )
            ),
        )

        worker_record_counts[key] = (
            worker_record_counts.get(
                key,
                0,
            )
            + 1
        )

    sorted_workers = sorted(
        worker_totals.items(),
        key=lambda item: item[1],
        reverse=True,
    )

    worker_row = 2

    for (
        employee_id,
        employee_name,
    ), seconds in sorted_workers:
        percentage = (
            seconds / total_seconds
            if total_seconds > 0
            else 0
        )

        worker_sheet.append(
            [
                employee_id,
                employee_name,
                format_duration(seconds),
                round(seconds / 60, 2),
                percentage,
                worker_record_counts.get(
                    (
                        employee_id,
                        employee_name,
                    ),
                    0,
                ),
            ]
        )

        worker_sheet[
            f"E{worker_row}"
        ].number_format = "0.0%"

        worker_row += 1

    if worker_sheet.max_row >= 2:
        style_table_area(
            worker_sheet,
            min_row=2,
            max_row=worker_sheet.max_row,
            min_column=1,
            max_column=6,
        )

    worker_sheet.freeze_panes = "A2"
    worker_sheet.auto_filter.ref = (
        worker_sheet.dimensions
    )

    set_column_widths(
        worker_sheet,
        {
            1: 12,
            2: 30,
            3: 16,
            4: 14,
            5: 14,
            6: 16,
        },
    )

    # ========================================================
    # ULOŽENÍ
    # ========================================================

    output = BytesIO()

    workbook.save(
        output
    )

    filename = (
        "prehled_cinnosti_"
        + period_end_local.strftime(
            "%Y-%m-%d_%H-%M"
        )
        + ".xlsx"
    )

    return (
        output.getvalue(),
        filename,
    )


# ============================================================
# TEXT DO E-MAILU
# ============================================================

def create_email_summary(
    records: list[dict],
    period_end_utc: datetime,
) -> str:
    activity_totals = calculate_activity_totals(
        records,
        period_end_utc,
    )
    machine_activity_totals = calculate_machine_activity_totals(
        records,
        period_end_utc,
    )
    total_seconds = sum(activity_totals.values())
    longest_record, longest_duration = find_longest_record(
        records,
        period_end_utc,
    )

    machine_order = {name: index for index, name in enumerate(MACHINES)}
    activity_order = {name: index for index, name in enumerate(ACTIVITIES)}
    sorted_machine_rows = sorted(
        machine_activity_totals.items(),
        key=lambda item: (
            machine_order.get(item[0][0], len(MACHINES)),
            item[0][0],
            activity_order.get(item[0][1], len(ACTIVITIES)),
            item[0][1],
        ),
    )

    lines = [
        "Souhrn podle stroje a činnosti:",
        "",
    ]

    if sorted_machine_rows:
        for (machine, activity), seconds in sorted_machine_rows:
            lines.append(
                f"{machine} → {activity}: {format_duration(seconds)}"
            )
    else:
        lines.append("Za dané období nejsou záznamy.")

    lines.extend(
        [
            "",
            "Souhrn podle činností:",
            "",
        ]
    )

    for activity in ACTIVITIES:
        seconds = activity_totals.get(activity, 0)
        percentage = (
            seconds / total_seconds * 100
            if total_seconds > 0
            else 0
        )
        lines.append(
            f"{activity}: {percentage:.1f} % "
            f"({format_duration(seconds)})"
        )

    lines.extend(
        [
            "",
            (
                "Celkový zaznamenaný čas: "
                f"{format_duration(total_seconds)}"
            ),
        ]
    )

    if longest_record:
        lines.extend(
            [
                "",
                "Nejdelší jednotlivý záznam:",
                (
                    f"{longest_record.get('employee_name', '')} – "
                    f"{get_machine_name(longest_record)} → "
                    f"{longest_record.get('activity', '')} – "
                    f"{format_duration(longest_duration)}"
                ),
            ]
        )

    return "\n".join(lines)


# ============================================================
# ODESLÁNÍ E-MAILU
# ============================================================

def send_email(
    excel_data: bytes,
    filename: str,
    records: list[dict],
    period_start_utc: datetime,
    period_end_utc: datetime,
) -> None:
    period_start_local = (
        period_start_utc.astimezone(
            APP_TIMEZONE
        )
    )

    period_end_local = (
        period_end_utc.astimezone(
            APP_TIMEZONE
        )
    )

    summary_text = (
        create_email_summary(
            records,
            period_end_utc,
        )
    )

    message = EmailMessage()

    message["From"] = (
        GMAIL_ADDRESS
    )

    message["To"] = (
        REPORT_RECIPIENT
    )

    message["Subject"] = (
        "Denní přehled činností UWH – "
        + period_end_local.strftime(
            "%d.%m.%Y"
        )
    )

    message.set_content(
        f"""Dobrý den,

v příloze zasílám automatický přehled činností za posledních 24 hodin.

Období:
{period_start_local.strftime("%d.%m.%Y %H:%M")}
až
{period_end_local.strftime("%d.%m.%Y %H:%M")}

Počet záznamů: {len(records)}

{summary_text}

Excel obsahuje:
- Detail – všechny jednotlivé záznamy,
- Souhrn – procenta, grafy, nejdelší záznam a přehled Stroj → Činnost → Čas,
- Pracovníci – součet času podle pracovníků.

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

        smtp.send_message(
            message
        )


# ============================================================
# SPUŠTĚNÍ
# ============================================================

def main() -> None:
    print(
        "Spouštím nový report: "
        "Detail, Souhrn a Pracovníci."
    )

    (
        records,
        period_start_utc,
        period_end_utc,
    ) = load_records()

    excel_data, filename = (
        create_excel(
            records,
            period_start_utc,
            period_end_utc,
        )
    )

    send_email(
        excel_data=excel_data,
        filename=filename,
        records=records,
        period_start_utc=period_start_utc,
        period_end_utc=period_end_utc,
    )

    print(
        f"Report byl odeslán na "
        f"{REPORT_RECIPIENT}."
    )

    print(
        f"Název přílohy: "
        f"{filename}"
    )

    print(
        f"Počet záznamů: "
        f"{len(records)}"
    )


if __name__ == "__main__":
    main()

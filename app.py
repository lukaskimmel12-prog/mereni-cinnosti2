from __future__ import annotations

from datetime import datetime, timedelta, timezone
from html import escape
from io import BytesIO
from textwrap import dedent
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from supabase import Client, create_client


# ============================================================
# NASTAVENÍ
# ============================================================

APP_TZ = ZoneInfo("Europe/Prague")

YUSEN_ORANGE = "#F58220"
YUSEN_ORANGE_DARK = "#D96E13"
YUSEN_BLUE = "#00529B"
YUSEN_DARK_BLUE = "#003B70"
YUSEN_LIGHT_BLUE = "#E8F2FA"

BACKGROUND = "#EEF3F7"
WHITE = "#FFFFFF"
DARK_TEXT = "#172A3A"
GREY_TEXT = "#526574"
GREEN = "#14804A"
LIGHT_GREEN = "#E5F6ED"

PRACOVNICI = {
    "11122": "Běloubek František",
    "11138": "Popelka Filip",
    "11063": "Cieplik Lukáš",
    "10607": "Drapák Patrik",
    "11073": "Fiala Vladislav",
    "1661": "Herold Ladislav",
    "10680": "Horáček Josef",
    "11064": "Houžvička Lukáš",
    "10342": "Hyšpler Jan",
    "1477": "Janeček Václav",
    "1424": "Jeřábek Karel",
    "10904": "Jeřábek Viktor",
    "1423": "Kimmel Lukáš",
    "10457": "Leksa Václav",
    "10891": "Matíscsák Michal",
    "10484": "Mayerhofer Ladislav",
    "10846": "Mikšík Filip",
    "10009": "Mokoš Michal",
    "10501": "Pelikán Petr",
    "11040": "Pleticha Rostislav",
    "10898": "Svoboda Martin",
    "10932": "Valský Pavel",
    "10203": "Vitásek Jan",
    "11182": "Kellner Karel",
    "11483": "Kvasnička Tomáš",
    "11485": "Žemlová Veronika",
    "11486": "Liehmová Hana",
}

CINNOSTI = [
    "Aperam",
    "Personna",
    "SSI",
    "Zanini",
    "Rebound",
]

STROJE = [
    "F33",
    "F36",
    "F45",
    "F86",
    "F87",
    "F88",
    "F117",
    "F140",
    "F205",
    "F206",
    "FS04",
    "FS07 Lion",
    "FP88-LION",
]


# ============================================================
# STRÁNKA
# ============================================================

st.set_page_config(
    page_title="UWH Activity Tracker",
    page_icon="⏱️",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def render_html(html: str) -> None:
    st.html(dedent(html).strip())


# ============================================================
# CSS
# ============================================================

render_html(
    f"""
    <style>
        #MainMenu, footer, header {{
            visibility: hidden;
        }}

        .stApp {{
            background:
                radial-gradient(
                    circle at top right,
                    rgba(0, 82, 155, 0.09),
                    transparent 32%
                ),
                {BACKGROUND};
        }}

        .block-container {{
            max-width: 1380px;
            padding-top: 0.8rem;
            padding-bottom: 2.5rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }}

        .app-header {{
            position: relative;
            overflow: hidden;
            background:
                linear-gradient(
                    135deg,
                    {YUSEN_DARK_BLUE},
                    {YUSEN_BLUE}
                );
            border-radius: 24px;
            padding: 24px 24px 20px;
            margin-bottom: 14px;
            box-shadow:
                0 12px 28px rgba(0, 59, 112, 0.22);
        }}

        .app-header::after {{
            content: "";
            position: absolute;
            width: 190px;
            height: 190px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.07);
            top: -95px;
            right: -45px;
        }}

        .header-accent {{
            width: 72px;
            height: 7px;
            border-radius: 10px;
            background: {YUSEN_ORANGE};
            margin-bottom: 12px;
        }}

        .app-title {{
            position: relative;
            z-index: 2;
            color: white;
            font-size: 1.95rem;
            line-height: 1.05;
            font-weight: 950;
        }}

        .app-subtitle {{
            position: relative;
            z-index: 2;
            color: rgba(255, 255, 255, 0.86);
            font-size: 0.95rem;
            margin-top: 7px;
        }}

        .app-date {{
            position: relative;
            z-index: 2;
            display: inline-block;
            margin-top: 13px;
            padding: 6px 11px;
            border-radius: 30px;
            color: white;
            background: rgba(255, 255, 255, 0.13);
            font-size: 0.82rem;
            font-weight: 800;
        }}

        .employee-card {{
            display: flex;
            align-items: center;
            gap: 15px;
            background: white;
            border-radius: 20px;
            padding: 16px 17px;
            margin-bottom: 16px;
            border: 1px solid rgba(0, 82, 155, 0.09);
            box-shadow:
                0 7px 22px rgba(0, 59, 112, 0.10);
        }}

        .employee-avatar {{
            width: 56px;
            height: 56px;
            border-radius: 18px;
            display: flex;
            align-items: center;
            justify-content: center;
            background:
                linear-gradient(
                    135deg,
                    {YUSEN_ORANGE},
                    {YUSEN_ORANGE_DARK}
                );
            color: white;
            font-size: 1.65rem;
        }}

        .employee-info {{
            flex: 1;
        }}

        .employee-label {{
            color: {GREY_TEXT};
            font-size: 0.75rem;
            font-weight: 800;
            text-transform: uppercase;
        }}

        .employee-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.25rem;
            font-weight: 950;
            margin-top: 2px;
        }}

        .employee-id {{
            color: {GREY_TEXT};
            font-size: 0.88rem;
            margin-top: 3px;
        }}

        .online-chip {{
            color: {GREEN};
            background: {LIGHT_GREEN};
            border-radius: 30px;
            padding: 7px 11px;
            font-size: 0.74rem;
            font-weight: 900;
        }}

        .status-card {{
            background: white;
            border-radius: 24px;
            padding: 24px 18px;
            text-align: center;
            margin-bottom: 17px;
            border: 1px solid rgba(0, 82, 155, 0.09);
            box-shadow:
                0 9px 26px rgba(0, 59, 112, 0.11);
        }}

        .status-running {{
            border-top: 8px solid {YUSEN_ORANGE};
        }}

        .status-idle {{
            border-top: 8px solid {YUSEN_BLUE};
        }}

        .status-caption {{
            color: {GREY_TEXT};
            font-size: 0.78rem;
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: 1.1px;
        }}

        .status-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 2.25rem;
            font-weight: 950;
            margin-top: 8px;
        }}

        .status-time {{
            color: {YUSEN_ORANGE};
            font-size: 3.25rem;
            font-weight: 950;
            letter-spacing: 2px;
            margin-top: 14px;
        }}

        .status-start {{
            display: inline-block;
            color: {GREY_TEXT};
            background: #F1F5F8;
            border-radius: 30px;
            padding: 7px 12px;
            font-size: 0.84rem;
            font-weight: 750;
            margin-top: 14px;
        }}

        .idle-icon {{
            width: 68px;
            height: 68px;
            border-radius: 22px;
            background: {YUSEN_LIGHT_BLUE};
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 12px;
            color: {YUSEN_BLUE};
            font-size: 2rem;
            font-weight: 900;
        }}

        .section-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.18rem;
            font-weight: 950;
            margin: 19px 0 8px;
        }}

        .section-subtitle {{
            color: {GREY_TEXT};
            font-size: 0.88rem;
            margin-bottom: 11px;
        }}

        .selected-activity {{
            background: {YUSEN_LIGHT_BLUE};
            border: 2px solid #BED8EB;
            border-left: 8px solid {YUSEN_ORANGE};
            border-radius: 16px;
            padding: 14px 16px;
            margin: 12px 0 15px;
        }}

        .selected-label {{
            color: {GREY_TEXT};
            font-size: 0.76rem;
            font-weight: 850;
            text-transform: uppercase;
        }}

        .selected-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.22rem;
            font-weight: 950;
            margin-top: 2px;
        }}

        .machine-warning {{
            background: #FFF4E8;
            border: 2px solid #F5B36D;
            border-left: 8px solid #F58220;
            border-radius: 17px;
            padding: 15px 16px;
            margin: 14px 0 18px;
            color: #172A3A;
            box-shadow: 0 5px 16px rgba(217, 110, 19, 0.12);
        }}

        .machine-warning-title {{
            color: #B45309;
            font-size: 1rem;
            font-weight: 950;
            margin-bottom: 8px;
        }}

        .machine-warning-row {{
            color: #374151;
            font-size: 0.9rem;
            font-weight: 750;
            line-height: 1.45;
            margin-top: 4px;
        }}

        .metric-card {{
            min-height: 125px;
            height: 100%;
            background: white;
            border-radius: 20px;
            padding: 17px 18px;
            border: 1px solid rgba(0, 82, 155, 0.09);
            box-shadow:
                0 7px 20px rgba(0, 59, 112, 0.09);
        }}

        .metric-orange {{
            border-top: 7px solid {YUSEN_ORANGE};
        }}

        .metric-blue {{
            border-top: 7px solid {YUSEN_BLUE};
        }}

        .metric-green {{
            border-top: 7px solid {GREEN};
        }}

        .metric-label {{
            color: {GREY_TEXT};
            font-size: 0.75rem;
            font-weight: 850;
            text-transform: uppercase;
        }}

        .metric-value {{
            color: {YUSEN_DARK_BLUE};
            font-size: 2.15rem;
            font-weight: 950;
            margin-top: 11px;
        }}

        .metric-note {{
            color: {GREY_TEXT};
            font-size: 0.84rem;
            margin-top: 7px;
        }}

        .dashboard-card {{
            background: white;
            border-radius: 22px;
            padding: 19px;
            border: 1px solid rgba(0, 82, 155, 0.09);
            box-shadow:
                0 8px 24px rgba(0, 59, 112, 0.09);
            margin-top: 17px;
        }}

        .dashboard-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.22rem;
            font-weight: 950;
        }}

        .dashboard-description {{
            color: {GREY_TEXT};
            font-size: 0.86rem;
            margin-top: 3px;
            margin-bottom: 14px;
        }}

        .zone-card {{
            background: #F8FBFD;
            border: 1px solid #D6E2EA;
            border-left: 8px solid {YUSEN_BLUE};
            border-radius: 17px;
            padding: 14px 15px;
            margin-bottom: 11px;
        }}

        .zone-card-active {{
            border-left-color: {YUSEN_ORANGE};
        }}

        .zone-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .zone-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.08rem;
            font-weight: 950;
        }}

        .zone-count {{
            color: white;
            background: {YUSEN_BLUE};
            border-radius: 30px;
            padding: 5px 10px;
            font-size: 0.76rem;
            font-weight: 900;
        }}

        .zone-count-active {{
            background: {YUSEN_ORANGE};
        }}

        .worker-list {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 11px;
        }}

        .worker-chip {{
            display: inline-flex;
            align-items: center;
            gap: 7px;
            color: {YUSEN_DARK_BLUE};
            background: {LIGHT_GREEN};
            border: 1px solid #BDE4CC;
            border-radius: 30px;
            padding: 7px 11px;
            font-size: 0.83rem;
            font-weight: 850;
        }}

        .worker-dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: {GREEN};
        }}

        .empty-zone {{
            color: #7A8995;
            background: #F0F3F5;
            border-radius: 12px;
            padding: 9px 11px;
            font-size: 0.83rem;
            margin-top: 10px;
        }}

        .progress-row {{
            margin-bottom: 14px;
        }}

        .progress-top {{
            display: flex;
            justify-content: space-between;
            margin-bottom: 6px;
        }}

        .progress-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 0.92rem;
            font-weight: 900;
        }}

        .progress-value {{
            color: {YUSEN_ORANGE_DARK};
            font-size: 0.88rem;
            font-weight: 950;
        }}

        .progress-bg {{
            width: 100%;
            height: 13px;
            background: #E3EAF0;
            border-radius: 20px;
            overflow: hidden;
        }}

        .progress-fill {{
            height: 100%;
            border-radius: 20px;
            background:
                linear-gradient(
                    90deg,
                    {YUSEN_BLUE},
                    {YUSEN_ORANGE}
                );
        }}

        .active-row {{
            display: grid;
            grid-template-columns: 1.5fr 1fr 0.7fr;
            gap: 12px;
            align-items: center;
            background: #F8FAFC;
            border: 1px solid #DDE6ED;
            border-radius: 14px;
            padding: 11px 13px;
            margin-bottom: 8px;
        }}

        .active-name {{
            color: {YUSEN_DARK_BLUE};
            font-weight: 900;
        }}

        .active-activity {{
            color: {YUSEN_ORANGE_DARK};
            font-weight: 900;
        }}

        .active-time {{
            color: {GREEN};
            font-weight: 950;
            text-align: right;
        }}

        .history-card {{
            background: white;
            border-radius: 15px;
            padding: 12px 13px;
            border: 1px solid #DCE5EC;
            margin-bottom: 8px;
        }}

        .history-top {{
            display: flex;
            justify-content: space-between;
        }}

        .history-activity {{
            color: {YUSEN_DARK_BLUE};
            font-weight: 950;
        }}

        .history-duration {{
            color: {YUSEN_ORANGE_DARK};
            font-weight: 950;
        }}

        .history-time {{
            color: {GREY_TEXT};
            font-size: 0.81rem;
            margin-top: 4px;
        }}

        div[data-testid="stSelectbox"] label {{
            color: {YUSEN_DARK_BLUE} !important;
            font-weight: 900 !important;
        }}

        div[data-testid="stSelectbox"]
        div[data-baseweb="select"] > div {{
            min-height: 58px !important;
            background: white !important;
            border: 2px solid #C9D8E3 !important;
            border-radius: 15px !important;
        }}

        div[data-testid="stSelectbox"] span {{
            color: {YUSEN_DARK_BLUE} !important;
            font-weight: 800 !important;
            opacity: 1 !important;
        }}

        div[role="listbox"],
        div[role="option"] {{
            background: white !important;
            color: {YUSEN_DARK_BLUE} !important;
        }}

        div[role="option"] * {{
            color: {YUSEN_DARK_BLUE} !important;
        }}

        div.stButton > button {{
            width: 100%;
            min-height: 62px;
            border-radius: 16px;
            font-size: 1rem;
            font-weight: 950;
            border: none;
        }}

        div.stButton > button[kind="primary"] {{
            background:
                linear-gradient(
                    135deg,
                    {YUSEN_ORANGE},
                    {YUSEN_ORANGE_DARK}
                ) !important;
            color: white !important;
        }}

        div.stButton > button[kind="secondary"] {{
            background:
                linear-gradient(
                    135deg,
                    {YUSEN_BLUE},
                    {YUSEN_DARK_BLUE}
                ) !important;
            color: white !important;
        }}

        div.stButton > button p,
        div.stButton > button span {{
            color: white !important;
        }}

        div.stButton > button:disabled {{
            background: #A8B6C1 !important;
            opacity: 0.7;
        }}

        div[data-testid="stDownloadButton"] > button {{
            width: 100%;
            min-height: 62px;
            border-radius: 16px;
            border: none;
            background: {YUSEN_BLUE} !important;
            color: white !important;
            font-weight: 950;
        }}

        div[data-testid="stDownloadButton"] > button * {{
            color: white !important;
        }}

        div[data-testid="stExpander"] {{
            background: white !important;
            border: 1px solid #D4E0E8 !important;
            border-radius: 17px !important;
            overflow: hidden;
        }}

        div[data-testid="stExpander"] summary p,
        div[data-testid="stExpander"] summary span {{
            color: {YUSEN_DARK_BLUE} !important;
            font-weight: 900 !important;
        }}

        div[data-testid="stAlert"] p,
        div[data-testid="stAlert"] strong {{
            color: {DARK_TEXT} !important;
        }}

        @media (max-width: 800px) {{
            .block-container {{
                padding-left: 0.7rem;
                padding-right: 0.7rem;
            }}

            .app-title {{
                font-size: 1.55rem;
            }}

            .status-time {{
                font-size: 2.55rem;
            }}

            .active-row {{
                grid-template-columns: 1fr;
                gap: 4px;
            }}

            .active-time {{
                text-align: left;
            }}

            .online-chip {{
                display: none;
            }}
        }}
    </style>
    """
)


# ============================================================
# SESSION STATE
# ============================================================

if "page" not in st.session_state:
    st.session_state.page = (
        st.query_params.get("page")
        if st.query_params.get("page") in ["evidence", "dashboard"]
        else "evidence"
    )

employee_from_url = st.query_params.get("employee")

if "logged_employee_id" not in st.session_state:
    if employee_from_url in PRACOVNICI:
        st.session_state.logged_employee_id = employee_from_url
    else:
        st.session_state.logged_employee_id = None

if "selected_machine" not in st.session_state:
    st.session_state.selected_machine = None

if "selected_activity" not in st.session_state:
    st.session_state.selected_activity = None


# ============================================================
# SUPABASE
# ============================================================

@st.cache_resource
def get_supabase() -> Client:
    try:
        return create_client(
            st.secrets["supabase"]["url"],
            st.secrets["supabase"]["key"],
        )
    except Exception:
        st.error(
            "Chybí nebo je chybně nastavené připojení k Supabase."
        )
        st.stop()


db = get_supabase()


# ============================================================
# FUNKCE
# ============================================================

def parse_dt(value: str) -> datetime:
    return datetime.fromisoformat(
        value.replace("Z", "+00:00")
    )


def local_dt(value: str) -> datetime:
    return parse_dt(value).astimezone(APP_TZ)


def format_duration(seconds: int | float | None) -> str:
    total = max(0, int(seconds or 0))

    hours, remainder = divmod(total, 3600)
    minutes, seconds = divmod(remainder, 60)

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def get_active_record(
    database: Client,
    employee_id: str,
) -> dict | None:
    response = (
        database.table("activity_log")
        .select("*")
        .eq("employee_id", employee_id)
        .is_("end_time", "null")
        .order("start_time", desc=True)
        .limit(1)
        .execute()
    )

    return response.data[0] if response.data else None


def load_all_active_records(
    database: Client,
) -> list[dict]:
    response = (
        database.table("activity_log")
        .select("*")
        .is_("end_time", "null")
        .order("start_time", desc=False)
        .execute()
    )

    return response.data or []


def load_active_machine_records(
    database: Client,
    machine: str,
) -> list[dict]:
    response = (
        database.table("activity_log")
        .select("*")
        .eq("machine", machine)
        .is_("end_time", "null")
        .order("start_time", desc=False)
        .execute()
    )

    return response.data or []


def start_activity(
    database: Client,
    employee_id: str,
    employee_name: str,
    machine: str,
    activity: str,
) -> None:
    database.table("activity_log").insert(
        {
            "employee_id": employee_id,
            "employee_name": employee_name,
            "machine": machine,
            "activity": activity,
            "start_time": datetime.now(
                timezone.utc
            ).isoformat(),
        }
    ).execute()


def end_activity(
    database: Client,
    record: dict,
) -> int:
    end_time = datetime.now(timezone.utc)
    start_time = parse_dt(record["start_time"])

    duration_seconds = max(
        0,
        int((end_time - start_time).total_seconds()),
    )

    (
        database.table("activity_log")
        .update(
            {
                "end_time": end_time.isoformat(),
                "duration_seconds": duration_seconds,
            }
        )
        .eq("id", record["id"])
        .is_("end_time", "null")
        .execute()
    )

    return duration_seconds


def load_employee_history(
    database: Client,
    employee_id: str,
    limit: int = 8,
) -> list[dict]:
    response = (
        database.table("activity_log")
        .select("*")
        .eq("employee_id", employee_id)
        .order("start_time", desc=True)
        .limit(limit)
        .execute()
    )

    return response.data or []


def load_last_24_hours(
    database: Client,
) -> list[dict]:
    since = datetime.now(
        timezone.utc
    ) - timedelta(hours=24)

    response = (
        database.table("activity_log")
        .select("*")
        .gte("start_time", since.isoformat())
        .order("start_time", desc=False)
        .execute()
    )

    return response.data or []


def make_excel(rows: list[dict]) -> bytes:
    output_rows = []
    now_utc = datetime.now(timezone.utc)

    for row in rows:
        start_local = local_dt(row["start_time"])
        end_value = row.get("end_time")

        end_local = (
            local_dt(end_value)
            if end_value
            else None
        )

        if row.get("duration_seconds") is not None:
            duration_seconds = int(
                row["duration_seconds"]
            )
        else:
            duration_seconds = int(
                (
                    now_utc
                    - parse_dt(row["start_time"])
                ).total_seconds()
            )

        output_rows.append(
            {
                "Datum": start_local.strftime("%d.%m.%Y"),
                "ID": row["employee_id"],
                "Jméno": row["employee_name"],
                "Stroj": row.get("machine") or "Neuveden",
                "Činnost": row["activity"],
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
        "Stroj",
        "Činnost",
        "Start",
        "Konec",
        "Trvání",
        "Trvání v minutách",
        "Stav",
    ]

    dataframe = pd.DataFrame(
        output_rows,
        columns=columns,
    )

    buffer = BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl",
    ) as writer:
        dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Posledních 24 hodin",
        )

        worksheet = writer.sheets[
            "Posledních 24 hodin"
        ]

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
            "A": 14,
            "B": 12,
            "C": 28,
            "D": 16,
            "E": 17,
            "F": 12,
            "G": 12,
            "H": 15,
            "I": 21,
            "J": 14,
        }

        for column, width in widths.items():
            worksheet.column_dimensions[
                column
            ].width = width

    return buffer.getvalue()


# ============================================================
# HLAVIČKA
# ============================================================

now_local = datetime.now(APP_TZ)

day_translation = {
    "Monday": "Pondělí",
    "Tuesday": "Úterý",
    "Wednesday": "Středa",
    "Thursday": "Čtvrtek",
    "Friday": "Pátek",
    "Saturday": "Sobota",
    "Sunday": "Neděle",
}

today_text = now_local.strftime("%A %d.%m.%Y")

for english_day, czech_day in day_translation.items():
    today_text = today_text.replace(
        english_day,
        czech_day,
    )

render_html(
    f"""
    <div class="app-header">
        <div class="header-accent"></div>
        <div class="app-title">
            UWH ACTIVITY TRACKER
        </div>
        <div class="app-subtitle">
            Evidence pracovních činností a živý dashboard
        </div>
        <div class="app-date">
            {today_text}
        </div>
    </div>
    """
)


# ============================================================
# VLASTNÍ MENU
# ============================================================

menu_left, menu_right = st.columns(2)

with menu_left:
    if st.button(
        "🏠 EVIDENCE ČINNOSTÍ",
        type=(
            "primary"
            if st.session_state.page == "evidence"
            else "secondary"
        ),
        use_container_width=True,
    ):
        st.session_state.page = "evidence"
        st.query_params["page"] = "evidence"
        st.rerun()

with menu_right:
    if st.button(
        "📊 LIVE DASHBOARD",
        type=(
            "primary"
            if st.session_state.page == "dashboard"
            else "secondary"
        ),
        use_container_width=True,
    ):
        st.session_state.page = "dashboard"
        st.query_params["page"] = "dashboard"
        st.rerun()


# ============================================================
# LIVE DASHBOARD
# ============================================================

if st.session_state.page == "dashboard":

    @st.fragment(run_every="5s")
    def render_dashboard() -> None:
        try:
            records = load_all_active_records(db)

        except Exception as error:
            st.error(
                f"Nepodařilo se načíst dashboard: {error}"
            )
            return

        current_utc = datetime.now(timezone.utc)
        current_local = current_utc.astimezone(APP_TZ)

        grouped = {
            activity: []
            for activity in CINNOSTI
        }

        for record in records:
            activity = str(
                record.get("activity", "Neznámá")
            )

            if activity not in grouped:
                grouped[activity] = []

            grouped[activity].append(record)

        worker_count = len(records)

        occupied_count = sum(
            1
            for activity_records in grouped.values()
            if activity_records
        )

        busiest_activity = "Žádná"

        if records:
            busiest_activity = max(
                grouped,
                key=lambda activity: len(
                    grouped[activity]
                ),
            )

        longest_seconds = 0

        if records:
            oldest_start = min(
                parse_dt(record["start_time"])
                for record in records
            )

            longest_seconds = int(
                (
                    current_utc - oldest_start
                ).total_seconds()
            )

        render_html(
            f"""
            <div class="dashboard-card">
                <div class="dashboard-title">
                    Live přehled provozu
                </div>
                <div class="dashboard-description">
                    Poslední aktualizace:
                    {current_local.strftime("%d.%m.%Y %H:%M:%S")}
                    · obnovuje se každých 5 sekund
                </div>
            </div>
            """
        )

        metric_1, metric_2, metric_3, metric_4 = (
            st.columns(4)
        )

        with metric_1:
            render_html(
                f"""
                <div class="metric-card metric-green">
                    <div class="metric-label">
                        Právě pracuje
                    </div>
                    <div class="metric-value">
                        {worker_count}
                    </div>
                    <div class="metric-note">
                        aktivních pracovníků
                    </div>
                </div>
                """
            )

        with metric_2:
            render_html(
                f"""
                <div class="metric-card metric-blue">
                    <div class="metric-label">
                        Obsazené činnosti
                    </div>
                    <div class="metric-value">
                        {occupied_count}
                    </div>
                    <div class="metric-note">
                        z celkem {len(CINNOSTI)}
                    </div>
                </div>
                """
            )

        with metric_3:
            render_html(
                f"""
                <div class="metric-card metric-orange">
                    <div class="metric-label">
                        Nejvíce pracovníků
                    </div>
                    <div
                        class="metric-value"
                        style="font-size:1.45rem;"
                    >
                        {escape(busiest_activity)}
                    </div>
                    <div class="metric-note">
                        {
                            len(
                                grouped.get(
                                    busiest_activity,
                                    [],
                                )
                            )
                            if records
                            else 0
                        }
                        pracovníků
                    </div>
                </div>
                """
            )

        with metric_4:
            render_html(
                f"""
                <div class="metric-card metric-blue">
                    <div class="metric-label">
                        Nejdelší aktivita
                    </div>
                    <div
                        class="metric-value"
                        style="font-size:1.65rem;"
                    >
                        {format_duration(longest_seconds)}
                    </div>
                    <div class="metric-note">
                        aktuálně běžící záznam
                    </div>
                </div>
                """
            )

        if not records:
            st.info(
                "Momentálně není spuštěná žádná činnost."
            )
            return

        map_column, graph_column = st.columns(
            [1.25, 0.75]
        )

        with map_column:
            render_html(
                """
                <div class="dashboard-card">
                    <div class="dashboard-title">
                        Mapa skladu
                    </div>
                    <div class="dashboard-description">
                        Pracovníci podle právě spuštěné činnosti
                    </div>
                </div>
                """
            )

            for activity in CINNOSTI:
                workers = grouped.get(activity, [])

                zone_class = (
                    "zone-card zone-card-active"
                    if workers
                    else "zone-card"
                )

                count_class = (
                    "zone-count zone-count-active"
                    if workers
                    else "zone-count"
                )

                if workers:
                    chips = ""

                    for worker in workers:
                        worker_name = escape(
                            str(
                                worker.get(
                                    "employee_name",
                                    "",
                                )
                            )
                        )

                        chips += (
                            '<div class="worker-chip">'
                            '<span class="worker-dot"></span>'
                            f"{worker_name}"
                            "</div>"
                        )

                    zone_body = (
                        '<div class="worker-list">'
                        f"{chips}"
                        "</div>"
                    )

                else:
                    zone_body = (
                        '<div class="empty-zone">'
                        "Momentálně zde nikdo nepracuje."
                        "</div>"
                    )

                render_html(
                    f"""
                    <div class="{zone_class}">
                        <div class="zone-header">
                            <div class="zone-name">
                                {escape(activity.upper())}
                            </div>
                            <div class="{count_class}">
                                {len(workers)} pracovníků
                            </div>
                        </div>
                        {zone_body}
                    </div>
                    """
                )

        with graph_column:
            render_html(
                """
                <div class="dashboard-card">
                    <div class="dashboard-title">
                        Aktuální rozdělení
                    </div>
                    <div class="dashboard-description">
                        Podíl aktivních pracovníků podle činnosti
                    </div>
                </div>
                """
            )

            total = max(1, worker_count)

            for activity in CINNOSTI:
                count = len(
                    grouped.get(activity, [])
                )

                percentage = (
                    count / total * 100
                )

                render_html(
                    f"""
                    <div class="progress-row">
                        <div class="progress-top">
                            <div class="progress-name">
                                {escape(activity)}
                            </div>
                            <div class="progress-value">
                                {percentage:.1f} % · {count}
                            </div>
                        </div>
                        <div class="progress-bg">
                            <div
                                class="progress-fill"
                                style="width:{percentage:.2f}%"
                            ></div>
                        </div>
                    </div>
                    """
                )

            chart_data = pd.DataFrame(
                {
                    "Činnost": CINNOSTI,
                    "Pracovníci": [
                        len(grouped.get(activity, []))
                        for activity in CINNOSTI
                    ],
                }
            ).set_index("Činnost")

            st.bar_chart(
                chart_data,
                use_container_width=True,
            )

        render_html(
            """
            <div class="dashboard-card">
                <div class="dashboard-title">
                    Aktivní pracovníci
                </div>
                <div class="dashboard-description">
                    Přehled všech právě probíhajících záznamů
                </div>
            </div>
            """
        )

        for record in sorted(
            records,
            key=lambda item: parse_dt(
                item["start_time"]
            ),
        ):
            employee_name = escape(
                str(
                    record.get(
                        "employee_name",
                        "",
                    )
                )
            )

            machine = escape(
                str(
                    record.get(
                        "machine",
                        "",
                    )
                    or "Neuveden"
                )
            )

            activity = escape(
                str(
                    record.get(
                        "activity",
                        "",
                    )
                )
            )

            elapsed = int(
                (
                    current_utc
                    - parse_dt(record["start_time"])
                ).total_seconds()
            )

            render_html(
                f"""
                <div class="active-row">
                    <div class="active-name">
                        ● {employee_name}
                    </div>
                    <div class="active-activity">
                        {machine} · {activity}
                    </div>
                    <div class="active-time">
                        {format_duration(elapsed)}
                    </div>
                </div>
                """
            )

    render_dashboard()
    st.stop()


# ============================================================
# EVIDENCE – PŘIHLÁŠENÍ
# ============================================================

if not st.session_state.logged_employee_id:
    render_html(
        """
        <div class="dashboard-card">
            <div class="dashboard-title">
                👤 Přihlášení pracovníka
            </div>
            <div class="dashboard-description">
                Klepni na své jméno. Přihlášení proběhne ihned
                a klávesnice se na skeneru neotevře.
            </div>
        </div>
        """
    )

    excluded_employee_ids: set[str] = set()

    login_employees = sorted(
        [
            (employee_id, name)
            for employee_id, name in PRACOVNICI.items()
            if employee_id not in excluded_employee_ids
        ],
        key=lambda item: item[1].casefold(),
    )

    for row_start in range(0, len(login_employees), 2):
        employee_columns = st.columns(2)
        row_employees = login_employees[row_start:row_start + 2]

        for column_index, (employee_id_option, employee_name_option) in enumerate(
            row_employees
        ):
            with employee_columns[column_index]:
                if st.button(
                    f"{employee_name_option}\n\nID {employee_id_option}",
                    key=f"login_employee_{employee_id_option}",
                    type="secondary",
                    use_container_width=True,
                ):
                    st.session_state.logged_employee_id = employee_id_option
                    st.session_state.selected_machine = None
                    st.session_state.selected_activity = None
                    st.query_params["employee"] = employee_id_option
                    st.rerun()

    st.stop()


# ============================================================
# EVIDENCE – PRACOVNÍK
# ============================================================

employee_id = st.session_state.logged_employee_id

if employee_id not in PRACOVNICI:
    st.session_state.logged_employee_id = None
    st.query_params.pop("employee", None)
    st.rerun()

employee_name = PRACOVNICI[employee_id]

render_html(
    f"""
    <div class="employee-card">
        <div class="employee-avatar">
            👤
        </div>
        <div class="employee-info">
            <div class="employee-label">
                Přihlášený pracovník
            </div>
            <div class="employee-name">
                {employee_name}
            </div>
            <div class="employee-id">
                Osobní ID: {employee_id}
            </div>
        </div>
        <div class="online-chip">
            ● PŘIHLÁŠEN
        </div>
    </div>
    """
)

try:
    active = get_active_record(
        db,
        employee_id,
    )

except Exception as error:
    st.error(
        f"Nepodařilo se načíst data: {error}"
    )
    st.stop()


# ============================================================
# EVIDENCE – AKTIVNÍ ČINNOST
# ============================================================

if active:
    started_local = local_dt(
        active["start_time"]
    )

    @st.fragment(run_every="1s")
    def live_timer() -> None:
        elapsed = int(
            (
                datetime.now(timezone.utc)
                - parse_dt(active["start_time"])
            ).total_seconds()
        )

        render_html(
            f"""
            <div class="status-card status-running">
                <div class="status-caption">
                    Aktuálně probíhá
                </div>
                <div class="status-name">
                    {(active.get("machine") or "Neuveden").upper()}
                </div>
                <div class="status-caption" style="margin-top: 8px;">
                    {active["activity"].upper()}
                </div>
                <div class="status-time">
                    {format_duration(elapsed)}
                </div>
                <div class="status-start">
                    Start:
                    {started_local.strftime("%d.%m.%Y %H:%M:%S")}
                </div>
            </div>
            """
        )

    live_timer()

    if st.button(
        "■ UKONČIT ČINNOST",
        type="primary",
        use_container_width=True,
    ):
        duration = end_activity(
            db,
            active,
        )

        st.session_state.selected_machine = None
        st.session_state.selected_activity = None

        st.success(
            f"Činnost {active['activity']} byla ukončena. "
            f"Trvání: {format_duration(duration)}"
        )

        st.rerun()


# ============================================================
# EVIDENCE – VÝBĚR ČINNOSTI
# ============================================================

else:
    render_html(
        """
        <div class="status-card status-idle">
            <div class="idle-icon">
                Ⅱ
            </div>
            <div class="status-caption">
                Aktuální stav
            </div>
            <div class="status-name">
                ŽÁDNÁ ČINNOST
            </div>
        </div>
        """
    )

    render_html(
        """
        <div class="section-title">
            Vyber stroj
        </div>
        <div class="section-subtitle">
            Klepni na stroj, na kterém budeš pracovat.
        </div>
        """
    )

    # Stroje vykreslujeme po dvojicích v jednotlivých řádcích.
    # Na úzkém displeji skeneru tak zůstane zachované přesné pořadí
    # F33, F36, F45, F86... místo seskupení celého levého sloupce.
    for row_start in range(0, len(STROJE), 2):
        row_columns = st.columns(2)
        row_machines = STROJE[row_start:row_start + 2]

        for column, machine in zip(row_columns, row_machines):
            with column:
                selected = (
                    st.session_state.selected_machine
                    == machine
                )

                button_label = (
                    f"✓ {machine.upper()}"
                    if selected
                    else machine.upper()
                )

                if st.button(
                    button_label,
                    key=f"machine_{machine}",
                    type=(
                        "primary"
                        if selected
                        else "secondary"
                    ),
                    use_container_width=True,
                ):
                    st.session_state.selected_machine = machine
                    st.rerun()

    if st.session_state.selected_machine:
        try:
            machine_users = load_active_machine_records(
                db,
                st.session_state.selected_machine,
            )
        except Exception as error:
            machine_users = []
            st.warning(
                f"Nepodařilo se ověřit obsazenost stroje: {error}"
            )

        if machine_users:
            warning_rows = ""

            for machine_user in machine_users:
                worker_name = escape(
                    str(
                        machine_user.get(
                            "employee_name",
                            "Neznámý pracovník",
                        )
                    )
                )

                worker_activity = escape(
                    str(
                        machine_user.get(
                            "activity",
                            "Neznámá činnost",
                        )
                    )
                )

                start_value = machine_user.get("start_time")

                if start_value:
                    start_text = local_dt(
                        start_value
                    ).strftime("%H:%M")
                else:
                    start_text = "neuvedeno"

                warning_rows += (
                    '<div class="machine-warning-row">'
                    f"👤 {worker_name} · "
                    f"{worker_activity} · od {start_text}"
                    "</div>"
                )

            render_html(
                f"""
                <div class="machine-warning">
                    <div class="machine-warning-title">
                        ⚠️ Stroj {
                            escape(
                                st.session_state.selected_machine
                            )
                        } je aktuálně veden jako používaný
                    </div>
                    {warning_rows}
                    <div class="machine-warning-row"
                         style="margin-top:9px;font-weight:650;">
                        Činnost můžeš i přesto normálně zahájit.
                    </div>
                </div>
                """
            )

    render_html(
        """
        <div class="section-title">
            Vyber činnost
        </div>
        <div class="section-subtitle">
            Klepni na činnost, kterou chceš zahájit.
        </div>
        """
    )

    activity_columns = st.columns(2)

    for index, activity in enumerate(CINNOSTI):
        target_column = activity_columns[index % 2]

        with target_column:
            selected = (
                st.session_state.selected_activity
                == activity
            )

            button_label = (
                f"✓ {activity.upper()}"
                if selected
                else activity.upper()
            )

            if st.button(
                button_label,
                key=f"activity_{activity}",
                type=(
                    "primary"
                    if selected
                    else "secondary"
                ),
                use_container_width=True,
            ):
                st.session_state.selected_activity = activity
                st.rerun()

    if (
        st.session_state.selected_machine
        and st.session_state.selected_activity
    ):
        render_html(
            f"""
            <div class="selected-activity">
                <div class="selected-label">
                    Vybráno
                </div>
                <div class="selected-name">
                    {
                        st.session_state
                        .selected_machine
                        .upper()
                    }
                    ·
                    {
                        st.session_state
                        .selected_activity
                        .upper()
                    }
                </div>
            </div>
            """
        )
    else:
        missing_parts = []

        if not st.session_state.selected_machine:
            missing_parts.append("stroj")

        if not st.session_state.selected_activity:
            missing_parts.append("činnost")

        st.info(
            "Vyber "
            + " a ".join(missing_parts)
            + "."
        )

    if st.button(
        "▶ ZAHÁJIT ČINNOST",
        type="primary",
        use_container_width=True,
        disabled=not bool(
            st.session_state.selected_machine
            and st.session_state.selected_activity
        ),
    ):
        selected_machine = (
            st.session_state.selected_machine
        )

        selected_activity = (
            st.session_state.selected_activity
        )

        start_activity(
            db,
            employee_id,
            employee_name,
            selected_machine,
            selected_activity,
        )

        st.session_state.selected_machine = None
        st.session_state.selected_activity = None
        st.rerun()


# ============================================================
# ODHLÁŠENÍ
# ============================================================

st.write("")

if active:
    st.caption(
        "Před odhlášením je potřeba ukončit "
        "aktuální činnost."
    )

if st.button(
    "ODHLÁSIT PRACOVNÍKA",
    type="secondary",
    use_container_width=True,
    disabled=bool(active),
):
    st.session_state.logged_employee_id = None
    st.session_state.selected_machine = None
    st.session_state.selected_activity = None

    st.query_params.pop("employee", None)

    st.rerun()


# ============================================================
# HISTORIE
# ============================================================

with st.expander(
    "📋 Poslední činnosti pracovníka"
):
    history = load_employee_history(
        db,
        employee_id,
        limit=8,
    )

    if not history:
        st.info(
            "Zatím nejsou uložené žádné záznamy."
        )

    for record in history:
        start_local = local_dt(
            record["start_time"]
        )

        end_value = record.get("end_time")

        if end_value:
            end_local = local_dt(end_value)

            end_text = end_local.strftime(
                "%H:%M:%S"
            )

            duration_text = format_duration(
                record.get("duration_seconds")
            )

            time_text = (
                f"{start_local.strftime('%d.%m.%Y')} · "
                f"{start_local.strftime('%H:%M:%S')} "
                f"→ {end_text}"
            )

        else:
            elapsed = int(
                (
                    datetime.now(timezone.utc)
                    - parse_dt(record["start_time"])
                ).total_seconds()
            )

            duration_text = format_duration(
                elapsed
            )

            time_text = (
                f"{start_local.strftime('%d.%m.%Y')} · "
                f"{start_local.strftime('%H:%M:%S')} "
                "→ stále probíhá"
            )

        render_html(
            f"""
            <div class="history-card">
                <div class="history-top">
                    <div class="history-activity">
                        {escape((record.get("machine") or "Neuveden").upper())}
                        ·
                        {escape(record["activity"].upper())}
                    </div>
                    <div class="history-duration">
                        {duration_text}
                    </div>
                </div>
                <div class="history-time">
                    {time_text}
                </div>
            </div>
            """
        )


# ============================================================
# EXPORT
# ============================================================

with st.expander(
    "📊 Export záznamů"
):
    export_rows = load_last_24_hours(db)

    st.write(
        "Excel bude obsahovat záznamy "
        "za posledních 24 hodin."
    )

    st.caption(
        f"Počet nalezených záznamů: "
        f"{len(export_rows)}"
    )

    excel_data = make_excel(
        export_rows
    )

    filename = (
        "cinnosti_poslednich_24h_"
        + datetime.now(APP_TZ).strftime(
            "%Y-%m-%d_%H-%M"
        )
        + ".xlsx"
    )

    st.download_button(
        "📥 STÁHNOUT EXCEL",
        data=excel_data,
        file_name=filename,
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

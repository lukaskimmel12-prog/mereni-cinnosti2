from __future__ import annotations

from datetime import datetime, timedelta, timezone
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
RED = "#C62828"

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
    "1617": "Chárová Zdena",
    "1758": "Lechmanová Kateřina",
    "11196": "Štefková Klára",
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


# ============================================================
# NASTAVENÍ STRÁNKY
# ============================================================

st.set_page_config(
    page_title="UWH Activity Tracker",
    page_icon="⏱️",
    layout="centered",
    initial_sidebar_state="collapsed",
)


# ============================================================
# HTML
# ============================================================

def render_html(html: str) -> None:
    st.markdown(
        dedent(html).strip(),
        unsafe_allow_html=True,
    )


# ============================================================
# CSS
# ============================================================

render_html(
    f"""
    <style>
        #MainMenu {{
            visibility: hidden;
        }}

        footer {{
            visibility: hidden;
        }}

        header {{
            visibility: hidden;
        }}

        html {{
            scroll-behavior: smooth;
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
            max-width: 760px;
            padding-top: 0.8rem;
            padding-bottom: 2.5rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }}

        /* HLAVIČKA */

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
            padding: 26px 22px 22px;
            margin-bottom: 18px;
            box-shadow:
                0 12px 28px rgba(0, 59, 112, 0.22);
        }}

        .app-header::after {{
            content: "";
            position: absolute;
            width: 170px;
            height: 170px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.07);
            top: -80px;
            right: -45px;
        }}

        .header-accent {{
            width: 72px;
            height: 7px;
            border-radius: 10px;
            background: {YUSEN_ORANGE};
            margin-bottom: 13px;
        }}

        .app-title {{
            position: relative;
            z-index: 2;
            color: white;
            font-size: 2rem;
            line-height: 1.05;
            font-weight: 950;
            letter-spacing: 0.4px;
        }}

        .app-subtitle {{
            position: relative;
            z-index: 2;
            color: rgba(255, 255, 255, 0.86);
            font-size: 0.96rem;
            margin-top: 7px;
        }}

        .app-date {{
            position: relative;
            z-index: 2;
            display: inline-block;
            margin-top: 14px;
            padding: 6px 11px;
            border-radius: 30px;
            color: white;
            background: rgba(255, 255, 255, 0.13);
            font-size: 0.82rem;
            font-weight: 750;
        }}

        /* PŘIHLÁŠENÍ */

        .login-card {{
            background: white;
            border-radius: 22px;
            padding: 22px;
            box-shadow:
                0 8px 24px rgba(0, 59, 112, 0.10);
            border: 1px solid rgba(0, 82, 155, 0.09);
            margin-top: 8px;
        }}

        .login-icon {{
            width: 58px;
            height: 58px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 18px;
            background: {YUSEN_LIGHT_BLUE};
            font-size: 1.8rem;
            margin-bottom: 13px;
        }}

        .login-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.35rem;
            font-weight: 950;
            margin-bottom: 3px;
        }}

        .login-description {{
            color: {GREY_TEXT};
            font-size: 0.93rem;
            margin-bottom: 12px;
        }}

        /* PRACOVNÍK */

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
            flex: 0 0 auto;
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
            box-shadow:
                0 6px 14px rgba(245, 130, 32, 0.25);
        }}

        .employee-info {{
            flex: 1;
            min-width: 0;
        }}

        .employee-label {{
            color: {GREY_TEXT};
            font-size: 0.75rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.6px;
        }}

        .employee-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.25rem;
            font-weight: 950;
            margin-top: 1px;
        }}

        .employee-id {{
            color: {GREY_TEXT};
            font-size: 0.88rem;
            margin-top: 2px;
        }}

        .online-chip {{
            flex: 0 0 auto;
            color: {GREEN};
            background: #E5F6ED;
            border-radius: 30px;
            padding: 6px 10px;
            font-size: 0.74rem;
            font-weight: 900;
        }}

        /* STAV */

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
            line-height: 1.05;
            font-weight: 950;
            margin-top: 8px;
        }}

        .status-time {{
            color: {YUSEN_ORANGE};
            font-size: 3.25rem;
            line-height: 1;
            font-weight: 950;
            letter-spacing: 2px;
            margin-top: 16px;
        }}

        .status-start {{
            display: inline-block;
            color: {GREY_TEXT};
            background: #F1F5F8;
            border-radius: 30px;
            padding: 7px 12px;
            font-size: 0.84rem;
            font-weight: 750;
            margin-top: 15px;
        }}

        .idle-icon {{
            width: 70px;
            height: 70px;
            border-radius: 22px;
            background: {YUSEN_LIGHT_BLUE};
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 12px;
            font-size: 2rem;
        }}

        /* NADPISY */

        .section-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.15rem;
            font-weight: 950;
            margin: 20px 0 10px;
        }}

        .section-subtitle {{
            color: {GREY_TEXT};
            font-size: 0.88rem;
            margin-top: -5px;
            margin-bottom: 11px;
        }}

        /* VYBRANÁ ČINNOST */

        .selected-activity {{
            background:
                linear-gradient(
                    135deg,
                    {YUSEN_LIGHT_BLUE},
                    #F5FAFE
                );
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

        /* HISTORIE */

        .history-card {{
            background: white;
            border-radius: 16px;
            padding: 13px 14px;
            border: 1px solid #DCE5EC;
            margin-bottom: 9px;
            box-shadow:
                0 3px 10px rgba(0, 59, 112, 0.06);
        }}

        .history-top {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
        }}

        .history-activity {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.02rem;
            font-weight: 950;
        }}

        .history-duration {{
            color: {YUSEN_ORANGE_DARK};
            font-size: 1rem;
            font-weight: 950;
        }}

        .history-time {{
            color: {GREY_TEXT};
            font-size: 0.82rem;
            margin-top: 5px;
        }}

        .history-running {{
            color: {GREEN};
            font-weight: 850;
        }}

        /* SELECTBOX */

        div[data-testid="stSelectbox"] label {{
            color: {YUSEN_DARK_BLUE} !important;
            font-size: 1rem !important;
            font-weight: 900 !important;
        }}

        div[data-testid="stSelectbox"]
        div[data-baseweb="select"] > div {{
            min-height: 60px !important;
            background: white !important;
            border: 2px solid #C9D8E3 !important;
            border-radius: 15px !important;
            color: {YUSEN_DARK_BLUE} !important;
        }}

        div[data-testid="stSelectbox"]
        div[data-baseweb="select"] > div:focus-within {{
            border-color: {YUSEN_ORANGE} !important;
            box-shadow:
                0 0 0 2px rgba(245, 130, 32, 0.15) !important;
        }}

        div[data-testid="stSelectbox"]
        div[data-baseweb="select"] span {{
            color: {YUSEN_DARK_BLUE} !important;
            font-weight: 800 !important;
            opacity: 1 !important;
        }}

        div[data-testid="stSelectbox"]
        div[data-baseweb="select"] svg {{
            color: {YUSEN_DARK_BLUE} !important;
            fill: {YUSEN_DARK_BLUE} !important;
        }}

        div[role="listbox"] {{
            background: white !important;
            border: 1px solid #B7CBD9 !important;
            border-radius: 13px !important;
        }}

        div[role="option"] {{
            background: white !important;
            color: {YUSEN_DARK_BLUE} !important;
            font-weight: 750 !important;
        }}

        div[role="option"] * {{
            color: {YUSEN_DARK_BLUE} !important;
        }}

        div[role="option"]:hover {{
            background: {YUSEN_LIGHT_BLUE} !important;
        }}

        div[role="option"][aria-selected="true"] {{
            background: #D8E9F6 !important;
        }}

        /* TLAČÍTKA */

        div.stButton > button {{
            width: 100%;
            min-height: 66px;
            border-radius: 16px;
            font-size: 1.05rem;
            font-weight: 950;
            border: none;
            box-shadow:
                0 5px 14px rgba(0, 59, 112, 0.13);
            transition:
                transform 0.12s ease,
                box-shadow 0.12s ease;
        }}

        div.stButton > button:hover {{
            transform: translateY(-1px);
            box-shadow:
                0 7px 17px rgba(0, 59, 112, 0.18);
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
            color: white !important;
            opacity: 0.7;
            transform: none;
        }}

        /* DOWNLOAD */

        div[data-testid="stDownloadButton"] > button {{
            width: 100%;
            min-height: 64px;
            border-radius: 16px;
            border: none !important;
            background:
                linear-gradient(
                    135deg,
                    {YUSEN_BLUE},
                    {YUSEN_DARK_BLUE}
                ) !important;
            color: white !important;
            font-size: 1.02rem;
            font-weight: 950;
            box-shadow:
                0 5px 14px rgba(0, 59, 112, 0.17);
        }}

        div[data-testid="stDownloadButton"] > button * {{
            color: white !important;
            opacity: 1 !important;
        }}

        /* EXPANDERY */

        div[data-testid="stExpander"] {{
            background: white !important;
            border: 1px solid #D4E0E8 !important;
            border-radius: 17px !important;
            overflow: hidden;
            box-shadow:
                0 4px 14px rgba(0, 59, 112, 0.06);
        }}

        div[data-testid="stExpander"] summary {{
            min-height: 58px;
            background: white !important;
        }}

        div[data-testid="stExpander"] summary p,
        div[data-testid="stExpander"] summary span {{
            color: {YUSEN_DARK_BLUE} !important;
            font-weight: 900 !important;
            opacity: 1 !important;
        }}

        div[data-testid="stExpander"] summary svg {{
            color: {YUSEN_DARK_BLUE} !important;
            fill: {YUSEN_DARK_BLUE} !important;
        }}

        /* HLÁŠKY */

        div[data-testid="stAlert"] {{
            border-radius: 15px;
        }}

        div[data-testid="stAlert"] p,
        div[data-testid="stAlert"] strong {{
            color: {DARK_TEXT} !important;
            font-weight: 750 !important;
        }}

        .stCaption p {{
            color: {GREY_TEXT} !important;
        }}

        @media (max-width: 600px) {{
            .block-container {{
                padding-left: 0.75rem;
                padding-right: 0.75rem;
            }}

            .app-header {{
                border-radius: 20px;
                padding: 22px 18px;
            }}

            .app-title {{
                font-size: 1.65rem;
            }}

            .status-name {{
                font-size: 1.9rem;
            }}

            .status-time {{
                font-size: 2.65rem;
            }}

            .employee-card {{
                padding: 14px;
            }}

            .employee-name {{
                font-size: 1.08rem;
            }}

            .online-chip {{
                display: none;
            }}

            div.stButton > button {{
                min-height: 61px;
                font-size: 0.95rem;
            }}
        }}
    </style>
    """
)


# ============================================================
# SESSION STATE
# ============================================================

employee_from_url = st.query_params.get("employee")

if "logged_employee_id" not in st.session_state:
    if employee_from_url in PRACOVNICI:
        st.session_state.logged_employee_id = employee_from_url
    else:
        st.session_state.logged_employee_id = None

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
# ČAS
# ============================================================

def parse_dt(value: str) -> datetime:
    return datetime.fromisoformat(
        value.replace("Z", "+00:00")
    )


def local_dt(value: str) -> datetime:
    return parse_dt(value).astimezone(APP_TZ)


def format_duration(
    seconds: int | float | None,
) -> str:
    total = max(0, int(seconds or 0))

    hours, remainder = divmod(total, 3600)
    minutes, seconds = divmod(remainder, 60)

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


# ============================================================
# DATABÁZE
# ============================================================

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


def start_activity(
    database: Client,
    employee_id: str,
    employee_name: str,
    activity: str,
) -> None:
    database.table("activity_log").insert(
        {
            "employee_id": employee_id,
            "employee_name": employee_name,
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


def load_last_24_hours(
    database: Client,
) -> list[dict]:
    since = (
        datetime.now(timezone.utc)
        - timedelta(hours=24)
    )

    response = (
        database.table("activity_log")
        .select("*")
        .gte("start_time", since.isoformat())
        .order("start_time", desc=False)
        .execute()
    )

    return response.data or []


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


# ============================================================
# EXCEL
# ============================================================

def make_excel(
    rows: list[dict],
) -> bytes:
    output_rows = []
    now_utc = datetime.now(timezone.utc)

    for row in rows:
        start_local = local_dt(
            row["start_time"]
        )

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
                "Datum": start_local.strftime(
                    "%d.%m.%Y"
                ),
                "ID": row["employee_id"],
                "Jméno": row["employee_name"],
                "Činnost": row["activity"],
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
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

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
            "D": 17,
            "E": 12,
            "F": 12,
            "G": 15,
            "H": 21,
            "I": 14,
        }

        for column, width in widths.items():
            worksheet.column_dimensions[
                column
            ].width = width

    return buffer.getvalue()


# ============================================================
# HLAVIČKA
# ============================================================

today_text = datetime.now(APP_TZ).strftime(
    "%A %d.%m.%Y"
)

day_translation = {
    "Monday": "Pondělí",
    "Tuesday": "Úterý",
    "Wednesday": "Středa",
    "Thursday": "Čtvrtek",
    "Friday": "Pátek",
    "Saturday": "Sobota",
    "Sunday": "Neděle",
}

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
            Evidence pracovních činností
        </div>
        <div class="app-date">
            {today_text}
        </div>
    </div>
    """
)


# ============================================================
# PŘIHLÁŠENÍ
# ============================================================

if not st.session_state.logged_employee_id:
    render_html(
        """
        <div class="login-card">
            <div class="login-icon">👤</div>
            <div class="login-title">
                Přihlášení pracovníka
            </div>
            <div class="login-description">
                Vyber své jméno ze seznamu a pokračuj
                tlačítkem Přihlásit.
            </div>
        </div>
        """
    )

    employee_options = {
        f"{name} – ID {employee_id}": employee_id
        for employee_id, name in PRACOVNICI.items()
    }

    selected_employee = st.selectbox(
        "Pracovník",
        options=list(
            employee_options.keys()
        ),
        index=None,
        placeholder="Vyber své jméno",
    )

    if st.button(
        "PŘIHLÁSIT SE",
        type="primary",
        use_container_width=True,
        disabled=not bool(selected_employee),
    ):
        selected_employee_id = (
            employee_options[
                selected_employee
            ]
        )

        st.session_state.logged_employee_id = (
            selected_employee_id
        )

        st.session_state.selected_activity = None

        st.query_params["employee"] = (
            selected_employee_id
        )

        st.rerun()

    st.stop()


# ============================================================
# PŘIHLÁŠENÝ PRACOVNÍK
# ============================================================

employee_id = (
    st.session_state.logged_employee_id
)

if employee_id not in PRACOVNICI:
    st.session_state.logged_employee_id = None
    st.query_params.clear()
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


# ============================================================
# AKTUÁLNÍ ČINNOST
# ============================================================

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
# BĚŽÍCÍ ČINNOST
# ============================================================

if active:
    started_local = local_dt(
        active["start_time"]
    )

    @st.fragment(run_every="1s")
    def live_timer() -> None:
        elapsed_seconds = int(
            (
                datetime.now(timezone.utc)
                - parse_dt(
                    active["start_time"]
                )
            ).total_seconds()
        )

        render_html(
            f"""
            <div class="status-card status-running">
                <div class="status-caption">
                    Aktuálně probíhá
                </div>

                <div class="status-name">
                    {active["activity"].upper()}
                </div>

                <div class="status-time">
                    {format_duration(elapsed_seconds)}
                </div>

                <div class="status-start">
                    Start:
                    {
                        started_local.strftime(
                            "%d.%m.%Y %H:%M:%S"
                        )
                    }
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
        try:
            duration = end_activity(
                db,
                active,
            )

            st.session_state.selected_activity = None

            st.success(
                f"Činnost {active['activity']} "
                f"byla ukončena. "
                f"Trvání: {format_duration(duration)}"
            )

            st.rerun()

        except Exception as error:
            st.error(
                f"Činnost se nepodařilo "
                f"ukončit: {error}"
            )


# ============================================================
# VÝBĚR ČINNOSTI
# ============================================================

else:
    render_html(
        """
        <div class="status-card status-idle">
            <div class="idle-icon">
                ⏸
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
            Vyber činnost
        </div>

        <div class="section-subtitle">
            Klepni na činnost, kterou chceš zahájit.
        </div>
        """
    )

    left_column, right_column = st.columns(2)

    for index, activity in enumerate(
        CINNOSTI
    ):
        target_column = (
            left_column
            if index % 2 == 0
            else right_column
        )

        with target_column:
            is_selected = (
                st.session_state.selected_activity
                == activity
            )

            button_label = (
                f"✓ {activity.upper()}"
                if is_selected
                else activity.upper()
            )

            button_type = (
                "primary"
                if is_selected
                else "secondary"
            )

            if st.button(
                button_label,
                key=f"activity_{activity}",
                type=button_type,
                use_container_width=True,
            ):
                st.session_state.selected_activity = (
                    activity
                )

                st.rerun()

    if st.session_state.selected_activity:
        render_html(
            f"""
            <div class="selected-activity">
                <div class="selected-label">
                    Vybraná činnost
                </div>

                <div class="selected-name">
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
        st.info(
            "Vyber jednu z nabízených činností."
        )

    if st.button(
        "▶ ZAHÁJIT ČINNOST",
        type="primary",
        use_container_width=True,
        disabled=not bool(
            st.session_state.selected_activity
        ),
    ):
        try:
            selected_activity = (
                st.session_state.selected_activity
            )

            start_activity(
                db,
                employee_id,
                employee_name,
                selected_activity,
            )

            st.session_state.selected_activity = None

            st.rerun()

        except Exception as error:
            error_text = str(error).lower()

            if (
                "duplicate" in error_text
                or "one_active_activity" in error_text
            ):
                st.warning(
                    "Tento pracovník už má "
                    "spuštěnou činnost."
                )
            else:
                st.error(
                    f"Činnost se nepodařilo "
                    f"spustit: {error}"
                )


# ============================================================
# ODHLÁŠENÍ
# ============================================================

st.markdown("<br>", unsafe_allow_html=True)

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
    st.session_state.selected_activity = None
    st.query_params.clear()

    st.rerun()


# ============================================================
# HISTORIE
# ============================================================

with st.expander(
    "📋 Poslední činnosti pracovníka"
):
    try:
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

            end_value = record.get(
                "end_time"
            )

            if end_value:
                end_local = local_dt(
                    end_value
                )

                end_text = end_local.strftime(
                    "%H:%M:%S"
                )

                duration_text = format_duration(
                    record.get(
                        "duration_seconds"
                    )
                )

                time_text = (
                    f"{start_local.strftime('%d.%m.%Y')} "
                    f"• "
                    f"{start_local.strftime('%H:%M:%S')} "
                    f"→ {end_text}"
                )

                duration_class = (
                    "history-duration"
                )

            else:
                elapsed_seconds = int(
                    (
                        datetime.now(timezone.utc)
                        - parse_dt(
                            record["start_time"]
                        )
                    ).total_seconds()
                )

                duration_text = (
                    format_duration(
                        elapsed_seconds
                    )
                )

                time_text = (
                    f"{start_local.strftime('%d.%m.%Y')} "
                    f"• "
                    f"{start_local.strftime('%H:%M:%S')} "
                    f"→ stále probíhá"
                )

                duration_class = (
                    "history-duration "
                    "history-running"
                )

            render_html(
                f"""
                <div class="history-card">
                    <div class="history-top">
                        <div class="history-activity">
                            {
                                record["activity"]
                                .upper()
                            }
                        </div>

                        <div class="{duration_class}">
                            {duration_text}
                        </div>
                    </div>

                    <div class="history-time">
                        {time_text}
                    </div>
                </div>
                """
            )

    except Exception as error:
        st.error(
            f"Historii se nepodařilo "
            f"načíst: {error}"
        )


# ============================================================
# EXPORT
# ============================================================

with st.expander(
    "📊 Export záznamů"
):
    try:
        export_rows = load_last_24_hours(
            db
        )

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

    except Exception as error:
        st.error(
            f"Export se nepodařilo "
            f"připravit: {error}"
        )

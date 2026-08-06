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

FORKLIFT_ICON_DATA = "data:image/png;base64,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"

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

        .tv-dashboard-header {{
            background: linear-gradient(135deg, {YUSEN_DARK_BLUE}, {YUSEN_BLUE});
            color: white;
            border-radius: 22px;
            padding: 10px 24px;
            margin-top: 16px;
            margin-bottom: 16px;
            display: grid;
            grid-template-columns: minmax(95px, 1fr) auto minmax(95px, 1fr);
            align-items: center;
            gap: 18px;
            font-size: 1.8rem;
            font-weight: 950;
            letter-spacing: 1px;
            box-shadow: 0 10px 26px rgba(0, 59, 112, 0.18);
            overflow: hidden;
        }}

        .tv-dashboard-title {{
            text-align: center;
            white-space: nowrap;
        }}

        .tv-forklift-wrap {{
            display: flex;
            align-items: center;
            min-width: 0;
        }}

        .tv-forklift-wrap-left {{
            justify-content: flex-start;
        }}

        .tv-forklift-wrap-right {{
            justify-content: flex-end;
        }}

        .tv-forklift-image {{
            display: block;
            width: 112px;
            height: 62px;
            object-fit: contain;
            flex: 0 0 auto;
            filter: drop-shadow(0 4px 5px rgba(0, 0, 0, 0.22));
        }}

        .tv-forklift-left {{
            transform: none;
        }}

        .tv-forklift-right {{
            transform: scaleX(-1);
        }}

        @media (max-width: 800px) {{
            .tv-dashboard-header {{
                grid-template-columns: 68px 1fr 68px;
                gap: 7px;
                padding: 9px 10px;
                font-size: 1.1rem;
            }}

            .tv-forklift-image {{
                width: 66px;
                height: 42px;
            }}
        }}

        .machine-panel {{
            background: white;
            border-radius: 22px;
            padding: 18px;
            min-height: 245px;
            border: 1px solid rgba(0, 82, 155, 0.10);
            box-shadow: 0 8px 22px rgba(0, 59, 112, 0.09);
            margin-top: 16px;
        }}

        .machine-panel-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.25rem;
            font-weight: 950;
            margin-bottom: 14px;
        }}

        .machine-grid {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 10px;
        }}

        .free-machine-card {{
            background: {LIGHT_GREEN};
            border: 1px solid #BDE4CC;
            border-left: 7px solid {GREEN};
            border-radius: 14px;
            padding: 12px 14px;
            color: {YUSEN_DARK_BLUE};
            font-size: 1.08rem;
            font-weight: 950;
        }}

        .occupied-machine-card {{
            background: #FFF2E8;
            border: 1px solid #F5C9A5;
            border-left: 7px solid {YUSEN_ORANGE};
            border-radius: 14px;
            padding: 11px 13px;
        }}

        .occupied-machine-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.05rem;
            font-weight: 950;
        }}

        .occupied-machine-worker {{
            color: {YUSEN_ORANGE_DARK};
            font-size: 0.93rem;
            font-weight: 900;
            margin-top: 3px;
        }}

        .tv-section-card {{
            background: white;
            border-radius: 22px;
            padding: 18px;
            margin-top: 16px;
            border: 1px solid rgba(0, 82, 155, 0.10);
            box-shadow: 0 8px 22px rgba(0, 59, 112, 0.09);
        }}

        .tv-table-header,
        .tv-activity-row {{
            display: grid;
            grid-template-columns: 1.3fr 0.9fr 1fr 0.8fr;
            gap: 12px;
            align-items: center;
        }}

        .tv-table-header {{
            color: white;
            background: {YUSEN_DARK_BLUE};
            border-radius: 13px;
            padding: 10px 14px;
            font-size: 0.82rem;
            font-weight: 900;
            text-transform: uppercase;
        }}

        .tv-activity-row {{
            background: #F8FAFC;
            border: 1px solid #DDE6ED;
            border-left: 7px solid {YUSEN_ORANGE};
            border-radius: 14px;
            padding: 12px 14px;
            margin-top: 9px;
        }}

        .tv-worker {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1rem;
            font-weight: 950;
        }}

        .tv-machine {{
            color: {YUSEN_BLUE};
            font-weight: 950;
        }}

        .tv-activity {{
            color: {YUSEN_ORANGE_DARK};
            font-weight: 950;
        }}

        .tv-time {{
            color: {GREEN};
            font-weight: 950;
            text-align: right;
            font-size: 1.02rem;
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
        if st.query_params.get("page") in ["evidence", "dashboard", "tv"]
        else "evidence"
    )

# Samostatný režim pro televizi: bez menu, bez přihlášení a bez exportu.
is_tv_mode = st.session_state.page == "tv"

if is_tv_mode:
    render_html(
        """
        <style>
            .block-container {
                max-width: 1600px !important;
                padding-top: 0.45rem !important;
                padding-bottom: 0.6rem !important;
            }
            html, body, [data-testid="stAppViewContainer"] {
                overflow: hidden !important;
                cursor: none !important;
            }
            * {
                cursor: none !important;
            }
        </style>
        """
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

if not is_tv_mode:
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

if st.session_state.page in ["dashboard", "tv"]:

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

        def surname_from_full_name(full_name: str) -> str:
            parts = str(full_name or "").strip().split()
            return parts[0] if parts else "Neuveden"

        occupied_by_machine: dict[str, list[dict]] = {}

        for record in records:
            machine = str(record.get("machine") or "Neuveden")
            occupied_by_machine.setdefault(machine, []).append(record)

        occupied_known_machines = {
            machine
            for machine in occupied_by_machine
            if machine in STROJE
        }

        free_machines = [
            machine
            for machine in STROJE
            if machine not in occupied_known_machines
        ]

        worker_count = len(records)
        occupied_machine_count = len(occupied_known_machines)

        render_html(
            f"""
            <div class="tv-dashboard-header">
                <div class="tv-forklift-wrap tv-forklift-wrap-left">
                    <img
                        class="tv-forklift-image tv-forklift-left"
                        src="{FORKLIFT_ICON_DATA}"
                        alt="Vysokozdvižný vozík"
                    >
                </div>
                <div class="tv-dashboard-title">UWH LIVE DASHBOARD</div>
                <div class="tv-forklift-wrap tv-forklift-wrap-right">
                    <img
                        class="tv-forklift-image tv-forklift-right"
                        src="{FORKLIFT_ICON_DATA}"
                        alt="Vysokozdvižný vozík"
                    >
                </div>
            </div>
            """
        )

        metric_1, metric_2, metric_3 = st.columns(3)

        with metric_1:
            render_html(
                f"""
                <div class="metric-card metric-green">
                    <div class="metric-label">Aktivních lidí</div>
                    <div class="metric-value">{worker_count}</div>
                    <div class="metric-note">
                        právě probíhajících záznamů
                    </div>
                </div>
                """
            )

        with metric_2:
            render_html(
                f"""
                <div class="metric-card metric-orange">
                    <div class="metric-label">Obsazené stroje</div>
                    <div class="metric-value">
                        {occupied_machine_count} / {len(STROJE)}
                    </div>
                    <div class="metric-note">
                        aktivně používaných strojů
                    </div>
                </div>
                """
            )

        with metric_3:
            render_html(
                f"""
                <div class="metric-card metric-blue">
                    <div class="metric-label">Volné stroje</div>
                    <div class="metric-value">{len(free_machines)}</div>
                    <div class="metric-note">
                        aktualizováno {current_local.strftime('%H:%M:%S')}
                    </div>
                </div>
                """
            )

        free_column, occupied_column = st.columns(2)

        with free_column:
            free_cards = "".join(
                f'<div class="free-machine-card">🟢 {escape(machine)}</div>'
                for machine in free_machines
            )

            if not free_cards:
                free_cards = (
                    '<div class="empty-zone">'
                    'Momentálně není volný žádný stroj.'
                    '</div>'
                )

            render_html(
                f"""
                <div class="machine-panel">
                    <div class="machine-panel-title">
                        🟢 VOLNÉ STROJE
                    </div>
                    <div class="machine-grid">
                        {free_cards}
                    </div>
                </div>
                """
            )

        with occupied_column:
            occupied_cards = ""

            for machine in STROJE:
                machine_records = occupied_by_machine.get(machine, [])

                for record in machine_records:
                    surname = escape(
                        surname_from_full_name(
                            str(record.get("employee_name", ""))
                        )
                    )
                    activity = escape(
                        str(record.get("activity", ""))
                    )

                    occupied_cards += f"""
                    <div class="occupied-machine-card">
                        <div class="occupied-machine-name">
                            🟠 {escape(machine)} – {surname}
                        </div>
                        <div class="occupied-machine-worker">
                            {activity}
                        </div>
                    </div>
                    """

            if not occupied_cards:
                occupied_cards = (
                    '<div class="empty-zone">'
                    'Momentálně není obsazený žádný stroj.'
                    '</div>'
                )

            render_html(
                f"""
                <div class="machine-panel">
                    <div class="machine-panel-title">
                        🔴 OBSAZENÉ STROJE
                    </div>
                    <div class="machine-grid">
                        {occupied_cards}
                    </div>
                </div>
                """
            )

        # Na TV zobrazujeme pouze hlavičku, metriky a přehled strojů.
        if is_tv_mode:
            return

        render_html(
            """
            <div class="tv-section-card">
                <div class="machine-panel-title">
                    AKTUÁLNÍ ČINNOSTI
                </div>
                <div class="tv-table-header">
                    <div>Příjmení</div>
                    <div>Stroj</div>
                    <div>Činnost</div>
                    <div style="text-align:right;">Čas</div>
                </div>
            </div>
            """
        )

        if not records:
            st.info("Momentálně není spuštěná žádná činnost.")
            return

        for record in sorted(
            records,
            key=lambda item: parse_dt(item["start_time"]),
        ):
            surname = escape(
                surname_from_full_name(
                    str(record.get("employee_name", ""))
                )
            )
            machine = escape(
                str(record.get("machine") or "Neuveden")
            )
            activity = escape(
                str(record.get("activity", ""))
            )
            elapsed = int(
                (
                    current_utc
                    - parse_dt(record["start_time"])
                ).total_seconds()
            )

            render_html(
                f"""
                <div class="tv-activity-row">
                    <div class="tv-worker">👤 {surname}</div>
                    <div class="tv-machine">{machine}</div>
                    <div class="tv-activity">{activity}</div>
                    <div class="tv-time">
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

    # Před vykreslením zjistíme, které stroje jsou právě používané.
    # Zelená tečka = volný stroj, oranžová tečka = aktivní záznam.
    try:
        all_active_records = load_all_active_records(db)
        occupied_machines = {
            str(record.get("machine", "")).strip()
            for record in all_active_records
            if str(record.get("machine", "")).strip()
        }
    except Exception as error:
        occupied_machines = set()
        st.warning(
            f"Nepodařilo se načíst stav strojů: {error}"
        )

    # Stroje vykreslujeme po dvojicích v jednotlivých řádcích.
    # Na úzkém displeji skeneru tak zůstane zachované přesné pořadí.
    for row_start in range(0, len(STROJE), 2):
        row_columns = st.columns(2)
        row_machines = STROJE[row_start:row_start + 2]

        for column, machine in zip(row_columns, row_machines):
            with column:
                selected = (
                    st.session_state.selected_machine
                    == machine
                )

                status_dot = (
                    "🟠"
                    if machine in occupied_machines
                    else "🟢"
                )

                button_label = (
                    f"✓ {status_dot} {machine.upper()}"
                    if selected
                    else f"{status_dot} {machine.upper()}"
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

# TV VZV compatibility override
render_html(
    '''
    <style>
        .tv-forklift-image,
        .tv-forklift-wrap {
            animation: none !important;
            transition: none !important;
        }
    </style>
    '''
)

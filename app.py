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
            position: relative;
            width: 100%;
            min-width: 120px;
            height: 70px;
            overflow: hidden;
        }}

        .tv-forklift-vehicle {{
            position: absolute;
            top: 1px;
            width: 112px;
            height: 68px;
            object-fit: contain;
            filter: drop-shadow(0 4px 5px rgba(0, 0, 0, 0.22));
            will-change: transform;
        }}

        .tv-forklift-left {{
            left: 0;
            animation: forklift-left-drive 6s ease-in-out infinite;
        }}

        .tv-forklift-right {{
            right: 0;
            animation: forklift-right-drive 6s ease-in-out infinite;
        }}

        .tv-forklift-right img {{
            transform: scaleX(-1);
        }}

        @keyframes forklift-left-drive {{
            0%, 12% {{ transform: translateX(0); }}
            45%, 55% {{ transform: translateX(calc(100% - 112px)); }}
            88%, 100% {{ transform: translateX(0); }}
        }}

        @keyframes forklift-right-drive {{
            0%, 12% {{ transform: translateX(0); }}
            45%, 55% {{ transform: translateX(calc(-100% + 112px)); }}
            88%, 100% {{ transform: translateX(0); }}
        }}

        @media (max-width: 800px) {{
            .tv-dashboard-header {{
                grid-template-columns: 72px 1fr 72px;
                gap: 8px;
                padding: 10px 10px;
                font-size: 1.2rem;
            }}

            .tv-forklift-wrap {{
                min-width: 74px;
                height: 48px;
            }}

            .tv-forklift-vehicle {{
                width: 72px;
                height: 46px;
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
                <div class="tv-forklift-wrap">
                    <div class="tv-forklift-vehicle tv-forklift-left">
                        <img src="data:image/png;base64,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" alt="Vysokozdvižný vozík" style="width:100%;height:100%;object-fit:contain;" />
                    </div>
                </div>
                <div class="tv-dashboard-title">UWH LIVE DASHBOARD</div>
                <div class="tv-forklift-wrap">
                    <div class="tv-forklift-vehicle tv-forklift-right">
                        <img src="data:image/png;base64,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" alt="Vysokozdvižný vozík" style="width:100%;height:100%;object-fit:contain;transform:scaleX(-1);" />
                    </div>
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

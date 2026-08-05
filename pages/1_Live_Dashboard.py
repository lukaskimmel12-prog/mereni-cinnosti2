from __future__ import annotations

from datetime import datetime, timezone
from html import escape
from textwrap import dedent
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
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
LIGHT_GREY = "#F2F5F7"

CINNOSTI = [
    "Aperam",
    "Personna",
    "SSI",
    "Zanini",
    "Rebound",
]


# ============================================================
# STRÁNKA
# ============================================================

st.set_page_config(
    page_title="Live Dashboard UWH",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def render_html(html: str) -> None:
    st.html(
        dedent(html).strip()
    )


# ============================================================
# VZHLED
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

        .stApp {{
            background:
                radial-gradient(
                    circle at top right,
                    rgba(0, 82, 155, 0.10),
                    transparent 34%
                ),
                {BACKGROUND};
        }}

        .block-container {{
            max-width: 1450px;
            padding-top: 0.8rem;
            padding-bottom: 2rem;
            padding-left: 1.2rem;
            padding-right: 1.2rem;
        }}

        .dashboard-header {{
            position: relative;
            overflow: hidden;
            background:
                linear-gradient(
                    135deg,
                    {YUSEN_DARK_BLUE},
                    {YUSEN_BLUE}
                );
            border-radius: 24px;
            padding: 25px 26px;
            margin-bottom: 18px;
            box-shadow:
                0 12px 28px rgba(0, 59, 112, 0.22);
        }}

        .dashboard-header::after {{
            content: "";
            position: absolute;
            width: 220px;
            height: 220px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.07);
            top: -120px;
            right: -50px;
        }}

        .header-line {{
            width: 74px;
            height: 7px;
            border-radius: 10px;
            background: {YUSEN_ORANGE};
            margin-bottom: 12px;
        }}

        .dashboard-title {{
            position: relative;
            z-index: 2;
            color: white;
            font-size: 2rem;
            font-weight: 950;
            line-height: 1.05;
        }}

        .dashboard-subtitle {{
            position: relative;
            z-index: 2;
            color: rgba(255, 255, 255, 0.86);
            font-size: 0.95rem;
            margin-top: 7px;
        }}

        .update-chip {{
            position: relative;
            z-index: 2;
            display: inline-block;
            color: white;
            background: rgba(255, 255, 255, 0.14);
            border-radius: 30px;
            padding: 7px 12px;
            font-size: 0.82rem;
            font-weight: 800;
            margin-top: 13px;
        }}

        .metric-card {{
            height: 100%;
            min-height: 135px;
            background: white;
            border-radius: 20px;
            padding: 18px 19px;
            border: 1px solid rgba(0, 82, 155, 0.09);
            box-shadow:
                0 7px 20px rgba(0, 59, 112, 0.09);
        }}

        .metric-accent-orange {{
            border-top: 7px solid {YUSEN_ORANGE};
        }}

        .metric-accent-blue {{
            border-top: 7px solid {YUSEN_BLUE};
        }}

        .metric-accent-green {{
            border-top: 7px solid {GREEN};
        }}

        .metric-label {{
            color: {GREY_TEXT};
            font-size: 0.76rem;
            font-weight: 850;
            text-transform: uppercase;
            letter-spacing: 0.7px;
        }}

        .metric-value {{
            color: {YUSEN_DARK_BLUE};
            font-size: 2.25rem;
            font-weight: 950;
            line-height: 1;
            margin-top: 12px;
        }}

        .metric-text {{
            color: {GREY_TEXT};
            font-size: 0.86rem;
            margin-top: 9px;
        }}

        .section-card {{
            background: white;
            border-radius: 22px;
            padding: 20px;
            border: 1px solid rgba(0, 82, 155, 0.09);
            box-shadow:
                0 8px 24px rgba(0, 59, 112, 0.09);
            margin-top: 18px;
        }}

        .section-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.25rem;
            font-weight: 950;
            margin-bottom: 4px;
        }}

        .section-description {{
            color: {GREY_TEXT};
            font-size: 0.88rem;
            margin-bottom: 15px;
        }}

        .activity-zone {{
            background:
                linear-gradient(
                    135deg,
                    #FFFFFF,
                    #F8FBFD
                );
            border: 1px solid #D6E2EA;
            border-left: 8px solid {YUSEN_BLUE};
            border-radius: 17px;
            padding: 15px 16px;
            margin-bottom: 12px;
        }}

        .activity-zone-active {{
            border-left-color: {YUSEN_ORANGE};
        }}

        .activity-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 10px;
        }}

        .activity-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.14rem;
            font-weight: 950;
        }}

        .activity-count {{
            color: white;
            background: {YUSEN_BLUE};
            border-radius: 30px;
            padding: 5px 10px;
            font-size: 0.78rem;
            font-weight: 900;
        }}

        .activity-count-active {{
            background: {YUSEN_ORANGE};
        }}

        .worker-list {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 12px;
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
            font-size: 0.84rem;
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
            background: {LIGHT_GREY};
            border-radius: 12px;
            padding: 10px 12px;
            font-size: 0.85rem;
            margin-top: 11px;
        }}

        .progress-row {{
            margin-bottom: 15px;
        }}

        .progress-top {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 12px;
            margin-bottom: 7px;
        }}

        .progress-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 0.94rem;
            font-weight: 900;
        }}

        .progress-value {{
            color: {YUSEN_ORANGE_DARK};
            font-size: 0.9rem;
            font-weight: 950;
        }}

        .progress-background {{
            width: 100%;
            height: 14px;
            overflow: hidden;
            background: #E3EAF0;
            border-radius: 20px;
        }}

        .progress-fill {{
            height: 100%;
            min-width: 0;
            border-radius: 20px;
            background:
                linear-gradient(
                    90deg,
                    {YUSEN_BLUE},
                    {YUSEN_ORANGE}
                );
        }}

        .active-worker-row {{
            display: grid;
            grid-template-columns:
                minmax(170px, 1.5fr)
                minmax(120px, 1fr)
                minmax(100px, 0.7fr);
            align-items: center;
            gap: 12px;
            background: #F8FAFC;
            border: 1px solid #DDE6ED;
            border-radius: 14px;
            padding: 11px 13px;
            margin-bottom: 8px;
        }}

        .active-worker-name {{
            color: {YUSEN_DARK_BLUE};
            font-weight: 900;
        }}

        .active-worker-activity {{
            color: {YUSEN_ORANGE_DARK};
            font-weight: 900;
        }}

        .active-worker-time {{
            color: {GREEN};
            font-weight: 950;
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}

        .no-data {{
            background: white;
            border-radius: 20px;
            padding: 30px;
            text-align: center;
            border: 1px solid #D9E3EA;
            color: {GREY_TEXT};
        }}

        div[data-testid="stAlert"] {{
            border-radius: 15px;
        }}

        div[data-testid="stAlert"] p {{
            color: {DARK_TEXT} !important;
            font-weight: 750 !important;
        }}

        @media (max-width: 800px) {{
            .dashboard-title {{
                font-size: 1.6rem;
            }}

            .metric-value {{
                font-size: 1.9rem;
            }}

            .active-worker-row {{
                grid-template-columns: 1fr;
                gap: 4px;
            }}

            .active-worker-time {{
                text-align: left;
            }}
        }}
    </style>
    """
)


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
            "Chybí nebo je chybně nastavené "
            "připojení k Supabase."
        )
        st.stop()


db = get_supabase()


# ============================================================
# POMOCNÉ FUNKCE
# ============================================================

def parse_dt(value: str) -> datetime:
    return datetime.fromisoformat(
        value.replace("Z", "+00:00")
    )


def format_duration(
    seconds: int | float | None,
) -> str:
    total = max(
        0,
        int(seconds or 0),
    )

    hours, remainder = divmod(
        total,
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


def load_active_records(
    database: Client,
) -> list[dict]:
    response = (
        database.table("activity_log")
        .select("*")
        .is_("end_time", "null")
        .order(
            "start_time",
            desc=False,
        )
        .execute()
    )

    return response.data or []


def group_by_activity(
    records: list[dict],
) -> dict[str, list[dict]]:
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

    return grouped


# ============================================================
# HLAVIČKA
# ============================================================

render_html(
    """
    <div class="dashboard-header">
        <div class="header-line"></div>
        <div class="dashboard-title">
            LIVE DASHBOARD UWH
        </div>
        <div class="dashboard-subtitle">
            Aktuální rozložení pracovníků podle činností
        </div>
        <div class="update-chip">
            Automatická aktualizace každých 5 sekund
        </div>
    </div>
    """
)


# ============================================================
# ŽIVÝ DASHBOARD
# ============================================================

@st.fragment(run_every="5s")
def live_dashboard() -> None:
    try:
        active_records = load_active_records(db)

    except Exception as error:
        st.error(
            f"Nepodařilo se načíst živá data: {error}"
        )
        return

    now_utc = datetime.now(timezone.utc)
    now_local = now_utc.astimezone(APP_TZ)

    grouped = group_by_activity(
        active_records
    )

    active_worker_count = len(
        active_records
    )

    occupied_activities = sum(
        1
        for records in grouped.values()
        if records
    )

    busiest_activity = "Žádná"

    if active_records:
        busiest_activity = max(
            grouped,
            key=lambda activity: len(
                grouped[activity]
            ),
        )

    oldest_seconds = 0

    if active_records:
        oldest_start = min(
            parse_dt(record["start_time"])
            for record in active_records
        )

        oldest_seconds = int(
            (
                now_utc - oldest_start
            ).total_seconds()
        )

    render_html(
        f"""
        <div class="update-chip"
             style="
                color: {YUSEN_DARK_BLUE};
                background: white;
                margin-top: 0;
                margin-bottom: 12px;
                box-shadow:
                    0 3px 10px
                    rgba(0, 59, 112, 0.08);
             ">
            Poslední aktualizace:
            {now_local.strftime("%d.%m.%Y %H:%M:%S")}
        </div>
        """
    )

    metric_1, metric_2, metric_3, metric_4 = (
        st.columns(4)
    )

    with metric_1:
        render_html(
            f"""
            <div class="
                metric-card
                metric-accent-green
            ">
                <div class="metric-label">
                    Právě pracuje
                </div>
                <div class="metric-value">
                    {active_worker_count}
                </div>
                <div class="metric-text">
                    pracovníků s aktivní činností
                </div>
            </div>
            """
        )

    with metric_2:
        render_html(
            f"""
            <div class="
                metric-card
                metric-accent-blue
            ">
                <div class="metric-label">
                    Obsazené činnosti
                </div>
                <div class="metric-value">
                    {occupied_activities}
                </div>
                <div class="metric-text">
                    z celkem {len(CINNOSTI)}
                </div>
            </div>
            """
        )

    with metric_3:
        render_html(
            f"""
            <div class="
                metric-card
                metric-accent-orange
            ">
                <div class="metric-label">
                    Nejvíce lidí
                </div>
                <div class="metric-value"
                     style="font-size: 1.55rem;">
                    {escape(busiest_activity)}
                </div>
                <div class="metric-text">
                    {
                        len(grouped.get(
                            busiest_activity,
                            [],
                        ))
                        if active_records
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
            <div class="
                metric-card
                metric-accent-blue
            ">
                <div class="metric-label">
                    Nejdelší aktivita
                </div>
                <div class="metric-value"
                     style="font-size: 1.7rem;">
                    {format_duration(oldest_seconds)}
                </div>
                <div class="metric-text">
                    aktuálně běžící záznam
                </div>
            </div>
            """
        )

    if not active_records:
        render_html(
            """
            <div class="no-data"
                 style="margin-top: 18px;">
                <div style="
                    font-size: 2rem;
                    margin-bottom: 8px;
                ">
                    ⏸
                </div>
                Momentálně není spuštěná žádná činnost.
            </div>
            """
        )
        return

    map_column, graph_column = st.columns(
        [1.25, 0.75]
    )

    with map_column:
        render_html(
            """
            <div class="section-card">
                <div class="section-title">
                    Mapa činností
                </div>
                <div class="section-description">
                    Pracovníci jsou rozděleni podle
                    právě spuštěné činnosti.
                </div>
            </div>
            """
        )

        for activity in CINNOSTI:
            workers = grouped.get(
                activity,
                [],
            )

            zone_class = (
                "activity-zone activity-zone-active"
                if workers
                else "activity-zone"
            )

            count_class = (
                "activity-count activity-count-active"
                if workers
                else "activity-count"
            )

            if workers:
                worker_chips = ""

                for worker in workers:
                    worker_name = escape(
                        str(
                            worker.get(
                                "employee_name",
                                "",
                            )
                        )
                    )

                    worker_chips += (
                        '<div class="worker-chip">'
                        '<span class="worker-dot"></span>'
                        f"{worker_name}"
                        "</div>"
                    )

                body = (
                    '<div class="worker-list">'
                    f"{worker_chips}"
                    "</div>"
                )

            else:
                body = (
                    '<div class="empty-zone">'
                    "Momentálně zde nikdo nepracuje."
                    "</div>"
                )

            render_html(
                f"""
                <div class="{zone_class}">
                    <div class="activity-header">
                        <div class="activity-name">
                            {escape(activity.upper())}
                        </div>
                        <div class="{count_class}">
                            {len(workers)}
                            pracovníků
                        </div>
                    </div>
                    {body}
                </div>
                """
            )

    with graph_column:
        render_html(
            """
            <div class="section-card">
                <div class="section-title">
                    Aktuální rozdělení
                </div>
                <div class="section-description">
                    Podíl pracovníků podle právě
                    spuštěné činnosti.
                </div>
            </div>
            """
        )

        total_workers = max(
            1,
            active_worker_count,
        )

        for activity in CINNOSTI:
            count = len(
                grouped.get(activity, [])
            )

            percentage = (
                count
                / total_workers
                * 100
            )

            render_html(
                f"""
                <div class="progress-row">
                    <div class="progress-top">
                        <div class="progress-name">
                            {escape(activity)}
                        </div>
                        <div class="progress-value">
                            {percentage:.1f} %
                            · {count}
                        </div>
                    </div>
                    <div class="progress-background">
                        <div
                            class="progress-fill"
                            style="
                                width:
                                {percentage:.2f}%;
                            ">
                        </div>
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
        <div class="section-card">
            <div class="section-title">
                Aktivní pracovníci
            </div>
            <div class="section-description">
                Přehled všech právě probíhajících
                záznamů a jejich aktuální délky.
            </div>
        </div>
        """
    )

    sorted_records = sorted(
        active_records,
        key=lambda record: parse_dt(
            record["start_time"]
        ),
    )

    for record in sorted_records:
        employee_name = escape(
            str(
                record.get(
                    "employee_name",
                    "",
                )
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

        elapsed_seconds = int(
            (
                now_utc
                - parse_dt(
                    record["start_time"]
                )
            ).total_seconds()
        )

        render_html(
            f"""
            <div class="active-worker-row">
                <div class="active-worker-name">
                    ● {employee_name}
                </div>
                <div class="active-worker-activity">
                    {activity}
                </div>
                <div class="active-worker-time">
                    {format_duration(elapsed_seconds)}
                </div>
            </div>
            """
        )


live_dashboard()

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from supabase import Client, create_client


APP_TZ = ZoneInfo("Europe/Prague")

YUSEN_ORANGE = "#F58220"
YUSEN_BLUE = "#00529B"
YUSEN_DARK_BLUE = "#003B70"
BACKGROUND = "#F4F7FA"

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

CINNOSTI = ["Aperam", "Personna", "SSI", "Zanini", "Rebound"]


st.set_page_config(
    page_title="Měření činností",
    page_icon="⏱️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    f"""
    <style>
        #MainMenu, footer, header {{
            visibility: hidden;
        }}

        .stApp {{
            background: {BACKGROUND};
        }}

        .block-container {{
            max-width: 720px;
            padding-top: 1rem;
            padding-bottom: 2rem;
        }}

        .app-header {{
            background: linear-gradient(135deg, {YUSEN_DARK_BLUE}, {YUSEN_BLUE});
            border-bottom: 8px solid {YUSEN_ORANGE};
            padding: 20px 18px;
            border-radius: 18px;
            margin-bottom: 20px;
            text-align: center;
            box-shadow: 0 5px 16px rgba(0, 59, 112, 0.20);
        }}

        .app-header h1 {{
            color: white;
            margin: 0;
            font-size: 2rem;
            font-weight: 900;
        }}

        .app-header p {{
            color: white;
            margin: 6px 0 0;
            font-size: 1rem;
        }}

        .user-box {{
            background: white;
            border-left: 8px solid {YUSEN_ORANGE};
            border-radius: 14px;
            padding: 15px 17px;
            margin-bottom: 16px;
            box-shadow: 0 3px 12px rgba(0, 82, 155, 0.12);
        }}

        .user-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.35rem;
            font-weight: 900;
        }}

        .user-id {{
            color: #555;
            font-size: 0.95rem;
            margin-top: 3px;
        }}

        .status-running,
        .status-idle {{
            background: white;
            border-radius: 16px;
            padding: 18px;
            margin: 12px 0 18px;
            text-align: center;
        }}

        .status-running {{
            border: 3px solid {YUSEN_ORANGE};
        }}

        .status-idle {{
            border: 3px solid {YUSEN_BLUE};
        }}

        .status-label {{
            color: #666;
            font-size: 0.95rem;
            font-weight: 700;
            text-transform: uppercase;
        }}

        .status-activity {{
            color: {YUSEN_DARK_BLUE};
            font-size: 2rem;
            font-weight: 900;
            margin-top: 4px;
        }}

        .status-time {{
            color: {YUSEN_ORANGE};
            font-size: 1.6rem;
            font-weight: 900;
            margin-top: 7px;
        }}

        .section-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.25rem;
            font-weight: 900;
            margin: 15px 0 8px;
        }}

        div[data-testid="stSelectbox"] label {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.08rem;
            font-weight: 900;
        }}

        div.stButton > button {{
            width: 100%;
            min-height: 62px;
            border-radius: 14px;
            font-size: 1.18rem;
            font-weight: 900;
            border: none;
        }}

        div.stButton > button[kind="primary"] {{
            background: {YUSEN_ORANGE};
            color: white;
        }}

        div.stButton > button[kind="secondary"] {{
            background: {YUSEN_BLUE};
            color: white;
        }}

        div[data-testid="stDownloadButton"] > button {{
            min-height: 58px;
            border-radius: 14px;
            background: {YUSEN_BLUE};
            color: white;
            font-size: 1.1rem;
            font-weight: 900;
        }}

        .small-note {{
            color: #5E6872;
            text-align: center;
            font-size: 0.9rem;
            margin-top: 8px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

if "logged_employee_id" not in st.session_state:
    st.session_state.logged_employee_id = None

if "selected_activity" not in st.session_state:
    st.session_state.selected_activity = None


@st.cache_resource
def get_supabase() -> Client:
    try:
        return create_client(
            st.secrets["supabase"]["url"],
            st.secrets["supabase"]["key"],
        )
    except Exception:
        st.error("Chybí nebo je chybně nastavené připojení k Supabase.")
        st.stop()


def parse_dt(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def local_dt(value: str) -> datetime:
    return parse_dt(value).astimezone(APP_TZ)


def format_duration(seconds: int | float | None) -> str:
    total = max(0, int(seconds or 0))
    hours, remainder = divmod(total, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def get_active_record(db: Client, employee_id: str) -> dict | None:
    response = (
        db.table("activity_log")
        .select("*")
        .eq("employee_id", employee_id)
        .is_("end_time", "null")
        .order("start_time", desc=True)
        .limit(1)
        .execute()
    )
    return response.data[0] if response.data else None


def start_activity(
    db: Client,
    employee_id: str,
    employee_name: str,
    activity: str,
) -> None:
    db.table("activity_log").insert(
        {
            "employee_id": employee_id,
            "employee_name": employee_name,
            "activity": activity,
            "start_time": datetime.now(timezone.utc).isoformat(),
        }
    ).execute()


def end_activity(db: Client, record: dict) -> int:
    end_time = datetime.now(timezone.utc)
    start_time = parse_dt(record["start_time"])
    duration_seconds = max(0, int((end_time - start_time).total_seconds()))

    (
        db.table("activity_log")
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


def load_last_24_hours(db: Client) -> list[dict]:
    since = datetime.now(timezone.utc) - timedelta(hours=24)
    response = (
        db.table("activity_log")
        .select("*")
        .gte("start_time", since.isoformat())
        .order("start_time", desc=False)
        .execute()
    )
    return response.data or []


def make_excel(rows: list[dict]) -> bytes:
    output_rows = []

    for row in rows:
        start_local = local_dt(row["start_time"])
        end_value = row.get("end_time")
        end_local = local_dt(end_value) if end_value else None

        if row.get("duration_seconds") is not None:
            duration_seconds = int(row["duration_seconds"])
        else:
            duration_seconds = int(
                (
                    datetime.now(timezone.utc)
                    - parse_dt(row["start_time"])
                ).total_seconds()
            )

        output_rows.append(
            {
                "Datum": start_local.strftime("%d.%m.%Y"),
                "ID": row["employee_id"],
                "Jméno": row["employee_name"],
                "Činnost": row["activity"],
                "Start": start_local.strftime("%H:%M:%S"),
                "Konec": end_local.strftime("%H:%M:%S") if end_local else "",
                "Trvání": format_duration(duration_seconds),
                "Trvání v minutách": round(duration_seconds / 60, 2),
                "Stav": "Dokončeno" if end_value else "Probíhá",
            }
        )

    df = pd.DataFrame(output_rows)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Posledních 24 hodin")
        ws = writer.sheets["Posledních 24 hodin"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    return buffer.getvalue()


st.markdown(
    """
    <div class="app-header">
        <h1>MĚŘENÍ ČINNOSTÍ</h1>
        <p>UWH • pracovní evidence</p>
    </div>
    """,
    unsafe_allow_html=True,
)

db = get_supabase()


if not st.session_state.logged_employee_id:
    st.markdown(
        '<div class="section-title">Přihlášení pracovníka</div>',
        unsafe_allow_html=True,
    )

    options = {
        f"{name} – ID {employee_id}": employee_id
        for employee_id, name in PRACOVNICI.items()
    }

    selected = st.selectbox(
        "Vyber pracovníka",
        options=list(options.keys()),
        index=None,
        placeholder="Klikni a vyber své jméno",
    )

    if st.button(
        "PŘIHLÁSIT",
        type="primary",
        use_container_width=True,
        disabled=not bool(selected),
    ):
        st.session_state.logged_employee_id = options[selected]
        st.session_state.selected_activity = None
        st.rerun()

    st.stop()


employee_id = st.session_state.logged_employee_id
employee_name = PRACOVNICI[employee_id]

st.markdown(
    f"""
    <div class="user-box">
        <div class="user-name">👤 {employee_name}</div>
        <div class="user-id">Osobní ID: {employee_id}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

active = get_active_record(db, employee_id)

if active:
    started_local = local_dt(active["start_time"])
    elapsed = int(
        (
            datetime.now(timezone.utc)
            - parse_dt(active["start_time"])
        ).total_seconds()
    )

    st.markdown(
        f"""
        <div class="status-running">
            <div class="status-label">Aktuálně probíhá</div>
            <div class="status-activity">{active["activity"].upper()}</div>
            <div class="status-time">{format_duration(elapsed)}</div>
            <div class="small-note">
                Začátek: {started_local.strftime("%d.%m.%Y %H:%M:%S")}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button(
        "🔴 END – UKONČIT ČINNOST",
        type="primary",
        use_container_width=True,
    ):
        duration = end_activity(db, active)
        st.success(
            f"Činnost {active['activity']} ukončena. "
            f"Trvání: {format_duration(duration)}"
        )
        st.rerun()

else:
    st.markdown(
        """
        <div class="status-idle">
            <div class="status-label">Aktuální stav</div>
            <div class="status-activity">ŽÁDNÁ ČINNOST</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-title">1. Vyber činnost</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    for index, activity in enumerate(CINNOSTI):
        with col1 if index % 2 == 0 else col2:
            selected_now = st.session_state.selected_activity == activity

            if st.button(
                f"✅ {activity.upper()}" if selected_now else activity.upper(),
                key=f"activity_{activity}",
                type="primary" if selected_now else "secondary",
                use_container_width=True,
            ):
                st.session_state.selected_activity = activity
                st.rerun()

    if st.session_state.selected_activity:
        st.success(
            f"Vybraná činnost: **{st.session_state.selected_activity}**"
        )

    st.markdown(
        '<div class="section-title">2. Zahaj měření</div>',
        unsafe_allow_html=True,
    )

    if st.button(
        "🟢 START",
        type="primary",
        use_container_width=True,
        disabled=not bool(st.session_state.selected_activity),
    ):
        start_activity(
            db,
            employee_id,
            employee_name,
            st.session_state.selected_activity,
        )
        st.session_state.selected_activity = None
        st.rerun()


st.divider()

if st.button(
    "🚪 ODHLÁSIT PRACOVNÍKA",
    type="secondary",
    use_container_width=True,
    disabled=bool(active),
):
    st.session_state.logged_employee_id = None
    st.session_state.selected_activity = None
    st.rerun()

if active:
    st.caption("Nejdříve ukonči aktuální činnost tlačítkem END.")


with st.expander("📥 Export do Excelu – posledních 24 hodin"):
    rows = load_last_24_hours(db)
    excel_data = make_excel(rows)

    filename = (
        "cinnosti_poslednich_24h_"
        + datetime.now(APP_TZ).strftime("%Y-%m-%d_%H-%M")
        + ".xlsx"
    )

    st.download_button(
        "STÁHNOUT EXCEL",
        data=excel_data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

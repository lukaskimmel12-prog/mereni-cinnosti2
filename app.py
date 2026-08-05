from __future__ import annotations

from datetime import datetime, timedelta, timezone
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from supabase import Client, create_client


# ============================================================
# ZÁKLADNÍ NASTAVENÍ
# ============================================================

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
    page_title="Měření činností",
    page_icon="⏱️",
    layout="centered",
    initial_sidebar_state="collapsed",
)


# ============================================================
# VZHLED APLIKACE
# ============================================================

st.markdown(
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
            background: {BACKGROUND};
        }}

        .block-container {{
            max-width: 760px;
            padding-top: 1rem;
            padding-bottom: 2rem;
        }}

        .app-header {{
            background: linear-gradient(
                135deg,
                {YUSEN_DARK_BLUE},
                {YUSEN_BLUE}
            );
            border-bottom: 8px solid {YUSEN_ORANGE};
            padding: 24px 18px;
            border-radius: 20px;
            margin-bottom: 20px;
            text-align: center;
            box-shadow: 0 5px 16px rgba(0, 59, 112, 0.20);
        }}

        .app-header h1 {{
            color: white;
            margin: 0;
            font-size: 2.15rem;
            font-weight: 900;
        }}

        .app-header p {{
            color: white;
            margin: 7px 0 0;
            font-size: 1rem;
        }}

        .user-box {{
            background: white;
            border-left: 9px solid {YUSEN_ORANGE};
            border-radius: 15px;
            padding: 16px 18px;
            margin-bottom: 18px;
            box-shadow: 0 3px 12px rgba(0, 82, 155, 0.12);
        }}

        .user-name {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.4rem;
            font-weight: 900;
        }}

        .user-id {{
            color: #555;
            font-size: 0.95rem;
            margin-top: 4px;
        }}

        .status-running {{
            background: white;
            border: 3px solid {YUSEN_ORANGE};
            border-radius: 18px;
            padding: 22px 18px;
            margin: 14px 0 20px;
            text-align: center;
            box-shadow: 0 4px 14px rgba(245, 130, 32, 0.15);
        }}

        .status-idle {{
            background: white;
            border: 3px solid {YUSEN_BLUE};
            border-radius: 18px;
            padding: 22px 18px;
            margin: 14px 0 20px;
            text-align: center;
            box-shadow: 0 4px 14px rgba(0, 82, 155, 0.12);
        }}

        .status-label {{
            color: #666;
            font-size: 0.95rem;
            font-weight: 800;
            text-transform: uppercase;
        }}

        .status-activity {{
            color: {YUSEN_DARK_BLUE};
            font-size: 2.2rem;
            font-weight: 900;
            margin-top: 6px;
        }}

        .status-time {{
            color: {YUSEN_ORANGE};
            font-size: 3rem;
            font-weight: 900;
            margin-top: 10px;
            letter-spacing: 2px;
        }}

        .status-start {{
            color: #5E6872;
            font-size: 0.95rem;
            margin-top: 10px;
        }}

        .section-title {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.25rem;
            font-weight: 900;
            margin: 17px 0 10px;
        }}

        .small-note {{
            color: #5E6872;
            text-align: center;
            font-size: 0.9rem;
            margin-top: 8px;
        }}

        div[data-testid="stSelectbox"] label {{
            color: {YUSEN_DARK_BLUE};
            font-size: 1.08rem;
            font-weight: 900;
        }}

        div[data-testid="stSelectbox"] > div > div {{
            min-height: 58px;
            border-radius: 13px;
            font-size: 1.05rem;
        }}

        div.stButton > button {{
            width: 100%;
            min-height: 66px;
            border-radius: 15px;
            font-size: 1.18rem;
            font-weight: 900;
            border: none;
            box-shadow: 0 3px 10px rgba(0, 0, 0, 0.10);
        }}

        div.stButton > button[kind="primary"] {{
            background: {YUSEN_ORANGE};
            color: white;
        }}

        div.stButton > button[kind="primary"]:hover {{
            background: #D96E13;
            color: white;
        }}

        div.stButton > button[kind="secondary"] {{
            background: {YUSEN_BLUE};
            color: white;
        }}

        div.stButton > button[kind="secondary"]:hover {{
            background: {YUSEN_DARK_BLUE};
            color: white;
        }}

        div.stButton > button:disabled {{
            background: #AAB4BE;
            color: white;
            opacity: 0.75;
        }}

        div[data-testid="stDownloadButton"] > button {{
            min-height: 60px;
            border-radius: 15px;
            background: {YUSEN_BLUE};
            color: white;
            font-size: 1.1rem;
            font-weight: 900;
        }}

        div[data-testid="stExpander"] {{
            background: white;
            border-radius: 14px;
            border: 1px solid #D5DCE3;
        }}

        @media (max-width: 600px) {{
            .block-container {{
                padding-left: 0.8rem;
                padding-right: 0.8rem;
            }}

            .app-header h1 {{
                font-size: 1.7rem;
            }}

            .status-time {{
                font-size: 2.5rem;
            }}

            .status-activity {{
                font-size: 1.9rem;
            }}
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

if "logged_employee_id" not in st.session_state:
    st.session_state.logged_employee_id = None

if "selected_activity" not in st.session_state:
    st.session_state.selected_activity = None


# ============================================================
# PŘIPOJENÍ K SUPABASE
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
# POMOCNÉ FUNKCE
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

    if response.data:
        return response.data[0]

    return None


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
        int(
            (
                end_time - start_time
            ).total_seconds()
        ),
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


def load_employee_history(
    database: Client,
    employee_id: str,
    limit: int = 10,
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
# EXCEL EXPORT
# ============================================================

def make_excel(
    rows: list[dict],
) -> bytes:
    output_rows = []

    for row in rows:
        start_local = local_dt(
            row["start_time"]
        )

        end_value = row.get("end_time")

        if end_value:
            end_local = local_dt(end_value)
        else:
            end_local = None

        if row.get("duration_seconds") is not None:
            duration_seconds = int(
                row["duration_seconds"]
            )
        else:
            duration_seconds = int(
                (
                    datetime.now(timezone.utc)
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

        widths = {
            "A": 13,
            "B": 11,
            "C": 27,
            "D": 16,
            "E": 12,
            "F": 12,
            "G": 14,
            "H": 20,
            "I": 13,
        }

        for column, width in widths.items():
            worksheet.column_dimensions[
                column
            ].width = width

    return buffer.getvalue()


# ============================================================
# HLAVIČKA
# ============================================================

st.markdown(
    """
    <div class="app-header">
        <h1>MĚŘENÍ ČINNOSTÍ</h1>
        <p>UWH • pracovní evidence</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PŘIHLÁŠENÍ
# ============================================================

if not st.session_state.logged_employee_id:
    st.markdown(
        """
        <div class="section-title">
            Přihlášení pracovníka
        </div>
        """,
        unsafe_allow_html=True,
    )

    employee_options = {
        f"{name} – ID {employee_id}": employee_id
        for employee_id, name in PRACOVNICI.items()
    }

    selected_employee = st.selectbox(
        "Vyber pracovníka",
        options=list(
            employee_options.keys()
        ),
        index=None,
        placeholder="Klikni a vyber své jméno",
    )

    login_clicked = st.button(
        "PŘIHLÁSIT",
        type="primary",
        use_container_width=True,
        disabled=not bool(selected_employee),
    )

    if login_clicked:
        st.session_state.logged_employee_id = (
            employee_options[selected_employee]
        )

        st.session_state.selected_activity = None

        st.rerun()

    st.markdown(
        """
        <div class="small-note">
            Vyber své jméno a klikni na PŘIHLÁSIT.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.stop()


# ============================================================
# PŘIHLÁŠENÝ PRACOVNÍK
# ============================================================

employee_id = (
    st.session_state.logged_employee_id
)

employee_name = PRACOVNICI[employee_id]

st.markdown(
    f"""
    <div class="user-box">
        <div class="user-name">
            👤 {employee_name}
        </div>

        <div class="user-id">
            Osobní ID: {employee_id}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# NAČTENÍ AKTUÁLNÍ ČINNOSTI
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
# ŽIVÉ STOPKY
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

        st.markdown(
            f"""
            <div class="status-running">
                <div class="status-label">
                    Aktuálně probíhá
                </div>

                <div class="status-activity">
                    {active["activity"].upper()}
                </div>

                <div class="status-time">
                    {format_duration(elapsed_seconds)}
                </div>

                <div class="status-start">
                    Začátek:
                    {
                        started_local.strftime(
                            "%d.%m.%Y %H:%M:%S"
                        )
                    }
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    live_timer()

    end_clicked = st.button(
        "🔴 END – UKONČIT ČINNOST",
        type="primary",
        use_container_width=True,
    )

    if end_clicked:
        try:
            duration = end_activity(
                db,
                active,
            )

            st.session_state.selected_activity = None

            st.success(
                f"Činnost {active['activity']} "
                f"byla ukončena. "
                f"Trvání: "
                f"{format_duration(duration)}"
            )

            st.rerun()

        except Exception as error:
            st.error(
                f"Činnost se nepodařilo "
                f"ukončit: {error}"
            )


# ============================================================
# ŽÁDNÁ ČINNOST NEBĚŽÍ
# ============================================================

else:
    st.markdown(
        """
        <div class="status-idle">
            <div class="status-label">
                Aktuální stav
            </div>

            <div class="status-activity">
                ŽÁDNÁ ČINNOST
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="section-title">
            1. Vyber činnost
        </div>
        """,
        unsafe_allow_html=True,
    )

    column_left, column_right = st.columns(2)

    for index, activity in enumerate(
        CINNOSTI
    ):
        if index % 2 == 0:
            target_column = column_left
        else:
            target_column = column_right

        with target_column:
            is_selected = (
                st.session_state.selected_activity
                == activity
            )

            if is_selected:
                button_text = (
                    f"✅ {activity.upper()}"
                )
                button_type = "primary"
            else:
                button_text = activity.upper()
                button_type = "secondary"

            activity_clicked = st.button(
                button_text,
                key=f"activity_{activity}",
                type=button_type,
                use_container_width=True,
            )

            if activity_clicked:
                st.session_state.selected_activity = (
                    activity
                )

                st.rerun()

    if st.session_state.selected_activity:
        st.success(
            "Vybraná činnost: "
            f"**{st.session_state.selected_activity}**"
        )
    else:
        st.warning(
            "Nejdříve vyber jednu činnost."
        )

    st.markdown(
        """
        <div class="section-title">
            2. Zahaj měření
        </div>
        """,
        unsafe_allow_html=True,
    )

    start_clicked = st.button(
        "🟢 START – ZAHÁJIT ČINNOST",
        type="primary",
        use_container_width=True,
        disabled=not bool(
            st.session_state.selected_activity
        ),
    )

    if start_clicked:
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

            st.success(
                f"Činnost {selected_activity} "
                f"byla spuštěna."
            )

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

st.divider()

if active:
    st.caption(
        "Pracovníka lze odhlásit až po "
        "ukončení aktuální činnosti."
    )

logout_clicked = st.button(
    "🚪 ODHLÁSIT PRACOVNÍKA",
    type="secondary",
    use_container_width=True,
    disabled=bool(active),
)

if logout_clicked:
    st.session_state.logged_employee_id = None
    st.session_state.selected_activity = None

    st.rerun()


# ============================================================
# HISTORIE PRACOVNÍKA
# ============================================================

with st.expander(
    "📋 Moje poslední záznamy"
):
    try:
        history = load_employee_history(
            db,
            employee_id,
            limit=10,
        )

        history_rows = []

        for row in history:
            start_local = local_dt(
                row["start_time"]
            )

            end_value = row.get("end_time")

            if end_value:
                end_local = local_dt(
                    end_value
                )

                end_text = end_local.strftime(
                    "%H:%M:%S"
                )

                duration_text = format_duration(
                    row.get("duration_seconds")
                )

                status_text = "Dokončeno"

            else:
                end_text = ""

                duration_text = format_duration(
                    (
                        datetime.now(timezone.utc)
                        - parse_dt(
                            row["start_time"]
                        )
                    ).total_seconds()
                )

                status_text = "Probíhá"

            history_rows.append(
                {
                    "Datum": start_local.strftime(
                        "%d.%m.%Y"
                    ),
                    "Činnost": row["activity"],
                    "Start": start_local.strftime(
                        "%H:%M:%S"
                    ),
                    "Konec": end_text,
                    "Trvání": duration_text,
                    "Stav": status_text,
                }
            )

        history_dataframe = pd.DataFrame(
            history_rows
        )

        st.dataframe(
            history_dataframe,
            use_container_width=True,
            hide_index=True,
        )

    except Exception as error:
        st.error(
            f"Historii se nepodařilo "
            f"načíst: {error}"
        )


# ============================================================
# EXPORT DO EXCELU
# ============================================================

with st.expander(
    "📥 Export do Excelu – posledních 24 hodin"
):
    try:
        export_rows = load_last_24_hours(
            db
        )

        st.write(
            f"Počet záznamů: "
            f"**{len(export_rows)}**"
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
            "STÁHNOUT EXCEL",
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

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from supabase import Client, create_client

APP_TZ = ZoneInfo("Europe/Prague")

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
    """
    <style>
        .block-container {max-width: 720px; padding-top: 1rem;}
        div.stButton > button {
            min-height: 4.2rem;
            font-size: 1.35rem;
            font-weight: 800;
            border-radius: 14px;
        }
        div[data-testid="stSelectbox"] label,
        div[data-testid="stTextInput"] label {
            font-size: 1.05rem;
            font-weight: 700;
        }
        .status-box {
            padding: 18px;
            border-radius: 14px;
            border: 2px solid #f0a500;
            background: rgba(240,165,0,.08);
            margin: 12px 0 18px 0;
        }
        .status-title {font-size: 1.5rem; font-weight: 800;}
        .status-line {font-size: 1.1rem; margin-top: 5px;}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def get_supabase() -> Client:
    try:
        return create_client(
            st.secrets["supabase"]["url"],
            st.secrets["supabase"]["key"],
        )
    except Exception as exc:
        st.error("Chybí nebo je chybně nastavené připojení k Supabase.")
        st.code(
            '[supabase]\nurl = "https://TVUJ-PROJEKT.supabase.co"\n'
            'key = "TVUJ_ANON_KEY"'
        )
        st.stop()
        raise exc


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


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
    db: Client, employee_id: str, employee_name: str, activity: str
) -> None:
    db.table("activity_log").insert(
        {
            "employee_id": employee_id,
            "employee_name": employee_name,
            "activity": activity,
            "start_time": utc_now_iso(),
        }
    ).execute()


def end_activity(db: Client, record: dict) -> None:
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

    output_rows = []
    for row in rows:
        start_local = local_dt(row["start_time"])
        end_value = row.get("end_time")
        end_local = local_dt(end_value) if end_value else None

        if row.get("duration_seconds") is not None:
            duration_seconds = int(row["duration_seconds"])
        elif not end_value:
            duration_seconds = int(
                (datetime.now(timezone.utc) - parse_dt(row["start_time"])).total_seconds()
            )
        else:
            duration_seconds = 0

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

    df = pd.DataFrame(output_rows, columns=columns)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Posledních 24 hodin")
        ws = writer.sheets["Posledních 24 hodin"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        widths = {
            "A": 13, "B": 11, "C": 27, "D": 16, "E": 12,
            "F": 12, "G": 13, "H": 20, "I": 13
        }
        for column, width in widths.items():
            ws.column_dimensions[column].width = width

    return buffer.getvalue()


db = get_supabase()

st.title("⏱️ Měření činností")
st.caption("Vyber pracovníka a činnost. Potom stiskni START nebo END.")

employee_options = {
    f"{name} – {employee_id}": employee_id
    for employee_id, name in PRACOVNICI.items()
}

selected_label = st.selectbox(
    "Pracovník",
    options=list(employee_options.keys()),
    index=None,
    placeholder="Vyber jméno nebo ID",
)
selected_activity = st.selectbox(
    "Činnost",
    options=CINNOSTI,
    index=None,
    placeholder="Vyber činnost",
)

if not selected_label:
    st.info("Nejdříve vyber pracovníka.")
else:
    employee_id = employee_options[selected_label]
    employee_name = PRACOVNICI[employee_id]

    try:
        active = get_active_record(db, employee_id)
    except Exception as exc:
        st.error(f"Nepodařilo se načíst data: {exc}")
        st.stop()

    if active:
        started = local_dt(active["start_time"])
        elapsed = int(
            (datetime.now(timezone.utc) - parse_dt(active["start_time"])).total_seconds()
        )

        st.markdown(
            f"""
            <div class="status-box">
                <div class="status-title">Probíhá: {active["activity"]}</div>
                <div class="status-line"><b>{active["employee_name"]}</b> · ID {active["employee_id"]}</div>
                <div class="status-line">Začátek: {started.strftime("%d.%m.%Y %H:%M:%S")}</div>
                <div class="status-line">Aktuální doba: <b>{format_duration(elapsed)}</b></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("🔴 END – ukončit činnost", use_container_width=True, type="primary"):
            try:
                end_activity(db, active)
                st.success(
                    f"Činnost {active['activity']} byla ukončena. "
                    f"Trvání: {format_duration(elapsed)}"
                )
                st.rerun()
            except Exception as exc:
                st.error(f"Činnost se nepodařilo ukončit: {exc}")
    else:
        if selected_activity:
            st.success(
                f"Připraveno: **{employee_name}** · **{selected_activity}**"
            )

        if st.button(
            "🟢 START – zahájit činnost",
            use_container_width=True,
            type="primary",
            disabled=not bool(selected_activity),
        ):
            try:
                start_activity(db, employee_id, employee_name, selected_activity)
                st.success(f"Spuštěno: {selected_activity}")
                st.rerun()
            except Exception as exc:
                if "one_active_activity_per_employee" in str(exc).lower() or "duplicate" in str(exc).lower():
                    st.warning("Tento pracovník už má rozpracovanou činnost.")
                else:
                    st.error(f"Činnost se nepodařilo spustit: {exc}")

st.divider()
with st.expander("📥 Export do Excelu – posledních 24 hodin"):
    try:
        export_rows = load_last_24_hours(db)
        st.write(f"Počet záznamů: **{len(export_rows)}**")

        excel_data = make_excel(export_rows)
        filename = (
            "cinnosti_poslednich_24h_"
            + datetime.now(APP_TZ).strftime("%Y-%m-%d_%H-%M")
            + ".xlsx"
        )

        st.download_button(
            "Stáhnout Excel",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as exc:
        st.error(f"Export se nepodařilo připravit: {exc}")

with st.expander("Poslední záznamy"):
    try:
        recent = (
            db.table("activity_log")
            .select("*")
            .order("start_time", desc=True)
            .limit(20)
            .execute()
            .data
            or []
        )

        display_rows = []
        for row in recent:
            start_local = local_dt(row["start_time"])
            if row.get("end_time"):
                end_local = local_dt(row["end_time"])
                duration = format_duration(row.get("duration_seconds"))
                status = "Dokončeno"
                end_text = end_local.strftime("%H:%M:%S")
            else:
                duration = format_duration(
                    (datetime.now(timezone.utc) - parse_dt(row["start_time"])).total_seconds()
                )
                status = "Probíhá"
                end_text = ""

            display_rows.append(
                {
                    "Datum": start_local.strftime("%d.%m.%Y"),
                    "ID": row["employee_id"],
                    "Jméno": row["employee_name"],
                    "Činnost": row["activity"],
                    "Start": start_local.strftime("%H:%M:%S"),
                    "Konec": end_text,
                    "Trvání": duration,
                    "Stav": status,
                }
            )

        st.dataframe(
            pd.DataFrame(display_rows),
            use_container_width=True,
            hide_index=True,
        )
    except Exception as exc:
        st.error(f"Historii se nepodařilo načíst: {exc}")

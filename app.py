import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import unicodedata

# ==================================================
# CONFIG
# ==================================================

st.set_page_config(
    page_title="Time Control Platform",
    page_icon="⏱️",
    layout="wide"
)

MINIMUM_DAILY_HOURS = 8

# ==================================================
# FUNCTIONS
# ==================================================

def normalize_name(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().upper()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = " ".join(text.split())
    return text


def build_user_mapping(df_names):
    mappings = []
    source_columns = ["NAME TG", "NAME CL", "NAME SB"]

    for _, row in df_names.iterrows():
        correct_name = row.get("NAME CORRECT")
        if pd.isna(correct_name):
            continue
        for col in source_columns:
            source_name = row.get(col)
            if pd.notna(source_name):
                mappings.append({
                    "SOURCE_NAME": normalize_name(source_name),
                    "NAME_CORRECT": correct_name
                })

    if len(mappings) == 0:
        return {}

    map_df = pd.DataFrame(mappings)
    return dict(zip(map_df["SOURCE_NAME"], map_df["NAME_CORRECT"]))


def convert_duration_to_hours(duration):
    try:
        td = pd.to_timedelta(str(duration))
        return round(td.total_seconds() / 3600, 4)
    except Exception:
        return 0


def get_novelty_status(user, target_date, df_novelties):
    result = df_novelties[
        (df_novelties["Persona"] == user) &
        (df_novelties["Fecha Inicio"] <= target_date) &
        (df_novelties["Fecha Fin"] >= target_date)
    ]
    if len(result) > 0:
        return result.iloc[0]["Tipo de Novedad"]
    return None


@st.cache_data
def process_files(toggl_file, camplegal_file, smokeball_file, resources_file,
                  novelties_file, clg_novelties_file, start_date, end_date):

    # =========================================
    # READ FILES
    # =========================================

    df_toggl = pd.read_excel(toggl_file, sheet_name="DataBaseToggl", engine="openpyxl")
    df_camplegal = pd.read_excel(camplegal_file, sheet_name="Time entries", engine="openpyxl")
    df_smokeball = pd.read_excel(smokeball_file, sheet_name="Entries", engine="openpyxl")
    df_names = pd.read_excel(resources_file, sheet_name="Names", engine="openpyxl")
    df_novelties = pd.read_excel(novelties_file, sheet_name="Novedades", engine="openpyxl")
    df_clg_novelties = pd.read_excel(clg_novelties_file, sheet_name="Novedades", engine="openpyxl")
    df_special_days = pd.read_excel(novelties_file, sheet_name="Novedades 2", engine="openpyxl")
    df_clg_special_days = pd.read_excel(clg_novelties_file, sheet_name="Novedades 2", engine="openpyxl")

    # =========================================
    # PARSE NOVELTIES
    # =========================================

    df_novelties["Fecha Inicio"] = pd.to_datetime(df_novelties["Fecha Inicio"], errors="coerce")
    df_novelties["Fecha Fin"] = pd.to_datetime(df_novelties["Fecha Fin"], errors="coerce")
    df_clg_novelties["Fecha Inicio"] = pd.to_datetime(df_clg_novelties["Fecha Inicio"], errors="coerce")
    df_clg_novelties["Fecha Fin"] = pd.to_datetime(df_clg_novelties["Fecha Fin"], errors="coerce")

    # =========================================
    # PARSE SPECIAL DAYS
    # =========================================

    df_special_days["Fecha"] = pd.to_datetime(df_special_days["Fecha"], errors="coerce")
    df_clg_special_days["Fecha"] = pd.to_datetime(df_clg_special_days["Fecha"], errors="coerce")

    # Normalizar columnas para días especiales
    df_special_days["PERSONA_NORMALIZADA"] = df_special_days["Persona"].astype(str).apply(normalize_name)
    df_special_days["TIPO_NORMALIZADO"] = df_special_days["Tipo de Novedad"].astype(str).str.upper().str.strip()

    df_clg_special_days["PERSONA_NORMALIZADA"] = df_clg_special_days["Persona"].astype(str).apply(normalize_name)
    df_clg_special_days["TIPO_NORMALIZADO"] = df_clg_special_days["Tipo de Novedad"].astype(str).str.upper().str.strip()

    # Normalizar columnas para novedades CLG
    df_clg_novelties["PERSONA_NORMALIZADA"] = df_clg_novelties["Persona"].astype(str).apply(normalize_name)

    # =========================================
    # USER MAP
    # =========================================

    user_map = build_user_mapping(df_names)

    # =========================================
    # NORMALIZE USERS
    # =========================================

    df_toggl["NORMALIZED_MEMBER"] = df_toggl["Member"].astype(str).apply(normalize_name)
    df_toggl["USER_CORRECT"] = df_toggl["NORMALIZED_MEMBER"].map(user_map).fillna(df_toggl["Member"])

    df_camplegal["NORMALIZED_MEMBER"] = df_camplegal["Staff Name"].astype(str).apply(normalize_name)
    df_camplegal["USER_CORRECT"] = df_camplegal["NORMALIZED_MEMBER"].map(user_map).fillna(df_camplegal["Staff Name"])

    df_smokeball["NORMALIZED_MEMBER"] = df_smokeball["Name"].astype(str).apply(normalize_name)
    df_smokeball["USER_CORRECT"] = df_smokeball["NORMALIZED_MEMBER"].map(user_map).fillna(df_smokeball["Name"])

    # =========================================
    # DATE
    # =========================================

    df_toggl["Date1"] = pd.to_datetime(df_toggl["Date1"], errors="coerce")
    df_camplegal["Date1"] = pd.to_datetime(df_camplegal["Date"], errors="coerce")
    df_smokeball["Date1"] = pd.to_datetime(df_smokeball["Date"], errors="coerce")

    # =========================================
    # FILTER DATES
    # =========================================

    start_dt = pd.to_datetime(start_date)
    end_dt = pd.to_datetime(end_date)

    df_toggl = df_toggl[
        (df_toggl["Date1"] >= start_dt) &
        (df_toggl["Date1"] <= end_dt)
    ]

    df_camplegal = df_camplegal[
        (df_camplegal["Date1"] >= start_dt) &
        (df_camplegal["Date1"] <= end_dt)
    ]

    df_smokeball = df_smokeball[
        (df_smokeball["Date1"] >= start_dt) &
        (df_smokeball["Date1"] <= end_dt)
    ]

    # =========================================
    # HOURS
    # =========================================

    df_toggl["Hours"] = df_toggl["Duration"].apply(convert_duration_to_hours)
    df_camplegal["Hours"] = pd.to_numeric(df_camplegal["Hours Spent"], errors="coerce").fillna(0)
    df_smokeball["Hours"] = pd.to_numeric(df_smokeball["Hours"], errors="coerce").fillna(0)

    # =========================================
    # STANDARDIZE COLUMNS
    # =========================================

    df_toggl["Activity"] = df_toggl["Project (Activity)"]
    df_toggl["Source"] = "Toggl"

    df_camplegal["Activity"] = df_camplegal["Activity"].astype(str)
    df_camplegal["Source"] = "Camp Legal"

    df_smokeball["Activity"] = df_smokeball["Subject"].astype(str)
    df_smokeball["Source"] = "Smokeball"

    # =========================================
    # SELECT COLUMNS
    # =========================================

    df_toggl_std = df_toggl[["Date1", "USER_CORRECT", "Hours", "Activity", "Source"]].copy()
    df_camplegal_std = df_camplegal[["Date1", "USER_CORRECT", "Hours", "Activity", "Source"]].copy()
    df_smokeball_std = df_smokeball[["Date1", "USER_CORRECT", "Hours", "Activity", "Source"]].copy()

    # =========================================
    # CONCATENATE
    # =========================================

    df_all_time = pd.concat(
        [df_toggl_std, df_camplegal_std, df_smokeball_std],
        ignore_index=True
    )

    # =========================================
    # DAILY REPORT
    # =========================================

    daily_report = (
        df_all_time
        .groupby(["Date1", "USER_CORRECT"], as_index=False)
        .agg(Total_Hours=("Hours", "sum"))
    )

    daily_report["Total_Hours"] = daily_report["Total_Hours"].round(2)
    daily_report["Status"] = "See Compliance Engine"

    # =========================================
    # EMPLOYEE INFO
    # =========================================

    employee_info = (
        df_names[["NAME CORRECT", "COMPANY", "DEPARTMENT", "TEAM"]]
        .drop_duplicates(subset=["NAME CORRECT"])
    )

    daily_report = daily_report.merge(
        employee_info,
        left_on="USER_CORRECT",
        right_on="NAME CORRECT",
        how="left"
    )

    # =========================================
    # USER SUMMARY
    # =========================================

    users_summary = (
        df_all_time
        .groupby("USER_CORRECT")
        .agg(Total_Hours=("Hours", "sum"), Entries=("Hours", "count"))
        .reset_index()
    )

    users_summary["Total_Hours"] = users_summary["Total_Hours"].round(2)

    users_summary = users_summary.merge(
        employee_info,
        left_on="USER_CORRECT",
        right_on="NAME CORRECT",
        how="left"
    )

    users_summary = users_summary.sort_values("Total_Hours", ascending=False)

    # =========================================
    # DETAIL REPORT
    # =========================================

    detail_report = df_all_time.copy()

    # =========================================
    # COMPLIANCE ENGINE
    # =========================================

    active_users = df_names[
        df_names["USER STATUS"].astype(str).str.upper().eq("ACTIVE")
    ]["NAME CORRECT"].dropna().unique()

    company_map = dict(zip(df_names["NAME CORRECT"], df_names["COMPANY"]))

    all_dates = pd.date_range(start=start_date, end=end_date, freq="D")

    compliance_records = []

    for current_day in all_dates:
        weekday = current_day.weekday()

        for user in active_users:
            normalized_user = normalize_name(user)
            user_company = str(company_map.get(user, "")).upper().strip()
            required_hours = 0
            holiday_assignments = pd.DataFrame()

            # =========================================
            # VERIFICAR SI EL USUARIO ESTÁ EN FESTIVO (NO TRABAJA)
            # =========================================
            holiday_user = df_special_days[
                (df_special_days["PERSONA_NORMALIZADA"] == normalized_user) &
                (df_special_days["TIPO_NORMALIZADO"] == "FESTIVO") &
                (df_special_days["Fecha"].dt.date == current_day.date())
            ]

            if len(holiday_user) > 0:
                continue

            # =========================================
            # VERIFICAR DÍA FESTIVO (PARA TODOS)
            # =========================================
            holiday_assignments = df_special_days[
                (df_special_days["PERSONA_NORMALIZADA"] == "TODOS") &
                (df_special_days["TIPO_NORMALIZADO"] == "FESTIVO") &
                (df_special_days["Fecha"].dt.date == current_day.date())
            ]

            is_holiday = len(holiday_assignments) > 0

            if is_holiday:
                holiday_worker = df_special_days[
                    (df_special_days["PERSONA_NORMALIZADA"] == normalized_user) &
                    (df_special_days["TIPO_NORMALIZADO"] == "FESTIVO_PROGRAMADO") &
                    (df_special_days["Fecha"].dt.date == current_day.date())
                ]
                if len(holiday_worker) > 0:
                    required_hours = 8
                else:
                    required_hours = 0
            else:
                # =========================================
                # LUNES A VIERNES
                # =========================================
                if weekday <= 4:
                    required_hours = 8

                # =========================================
                # SÁBADOS
                # =========================================
                elif weekday == 5:
                    if user_company == "MAX":
                        saturday_user = df_special_days[
                            (df_special_days["PERSONA_NORMALIZADA"] == normalized_user) &
                            (df_special_days["Fecha"].dt.date == current_day.date())
                        ]
                        if len(saturday_user) > 0:
                            required_hours = 4
                        else:
                            required_hours = 4

                    elif user_company == "CLG":
                        clg_saturday_user = df_clg_special_days[
                            (df_clg_special_days["PERSONA_NORMALIZADA"] == normalized_user) &
                            (df_clg_special_days["TIPO_NORMALIZADO"] == "SABADO") &
                            (df_clg_special_days["Fecha"].dt.date == current_day.date())
                        ]
                        if len(clg_saturday_user) > 0:
                            required_hours = 8
                        else:
                            required_hours = 0
                    else:
                        required_hours = 0

                # =========================================
                # DOMINGOS
                # =========================================
                else:
                    required_hours = 0

            # =========================================
            # SI NO HAY HORAS REQUERIDAS, SALTAR
            # =========================================
            if required_hours == 0:
                continue

            # =========================================
            # OBTENER HORAS TRABAJADAS
            # =========================================
            day_record = daily_report[
                (daily_report["USER_CORRECT"] == user) &
                (daily_report["Date1"].dt.date == current_day.date())
            ]

            worked_hours = 0.0
            if len(day_record) > 0:
                worked_hours = float(day_record["Total_Hours"].sum())

            # =========================================
            # OBTENER NOVEDAD
            # =========================================
            novelty = get_novelty_status(user, current_day, df_novelties)

            if novelty is None and user_company == "CLG":
                clg_novelty = df_clg_novelties[
                    (df_clg_novelties["PERSONA_NORMALIZADA"] == normalized_user) &
                    (current_day >= df_clg_novelties["Fecha Inicio"]) &
                    (current_day <= df_clg_novelties["Fecha Fin"])
                ]
                if len(clg_novelty) > 0:
                    novelty = str(clg_novelty.iloc[0]["Tipo de Novedad"])

            # =========================================
            # LÓGICA DE STATUS
            # =========================================
            if novelty is not None:
                status = f"🟡 {novelty}"
            elif worked_hours == 0:
                status = "❌ No registró tiempo"
            elif required_hours == 4:
                if worked_hours >= 3.5:
                    status = "✅ Cumple"
                else:
                    status = "❌ Horas insuficientes"
            elif worked_hours < required_hours:
                status = "❌ Horas insuficientes"
            else:
                status = "✅ Cumple"

            compliance_records.append({
                "Date": current_day.date(),
                "Weekday": weekday,
                "User": user,
                "Hours Worked": round(worked_hours, 2),
                "Hours Required": required_hours,
                "Novelty": novelty,
                "Status": status,
                "Holiday Count": len(holiday_assignments),
                "Current Day": current_day.date()
            })

    compliance_engine = pd.DataFrame(compliance_records)

    compliance_summary = (
        compliance_engine
        .groupby("User")
        .agg(
            Total_Hours=("Hours Worked", "sum"),
            Days_Evaluated=("Date", "count")
        )
        .reset_index()
    )

    success_counts = (
        compliance_engine[compliance_engine["Status"] == "✅ Cumple"]
        .groupby("User")
        .size()
    )

    failure_counts = (
        compliance_engine[
            compliance_engine["Status"].isin([
                "❌ No registró tiempo",
                "❌ Horas insuficientes"
            ])
        ]
        .groupby("User")
        .size()
    )

    novelty_counts = (
        compliance_engine[
            compliance_engine["Status"].astype(str).str.startswith("🟡")
        ]
        .groupby("User")
        .size()
    )

    compliance_summary["Compliant_Days"] = (
        compliance_summary["User"].map(success_counts).fillna(0)
    )

    compliance_summary["Non_Compliant_Days"] = (
        compliance_summary["User"].map(failure_counts).fillna(0)
    )

    compliance_summary["Justified_Days"] = (
        compliance_summary["User"].map(novelty_counts).fillna(0)
    )

    daily_status = (
        compliance_engine
        .sort_values("Date")
        .groupby("User")
        .apply(
            lambda df: " | ".join([
                f"{d.strftime('%m/%d')} {s}"
                for d, s in zip(df["Date"], df["Status"])
            ])
        )
    )

    compliance_summary["Daily_Status"] = (
        compliance_summary["User"].map(daily_status).fillna("")
    )

    # Calcular Failed_Dates para cada usuario
    failed_dates = (
        compliance_engine[
            compliance_engine["Status"].isin([
                "❌ No registró tiempo",
                "❌ Horas insuficientes"
            ])
        ]
        .groupby("User")["Date"]
        .apply(lambda dates: ", ".join(d.strftime("%m/%d") for d in dates))
    )

    compliance_summary["Failed_Dates"] = (
        compliance_summary["User"].map(failed_dates).fillna("Ninguna")
    )

    return daily_report, detail_report, users_summary, compliance_engine, compliance_summary


# ==================================================
# HEADER
# ==================================================

st.title("⏱️ Time Control Platform")
st.markdown("### Phase 8 - Toggl - Camp Legal - Smokeball Validation")

# ==================================================
# SIDEBAR
# ==================================================

st.sidebar.header("Upload Files")

resources_file = st.sidebar.file_uploader("Power BI Resources", type=["xlsx"])
toggl_file = st.sidebar.file_uploader("Toggl File", type=["xlsx"])
novelties_file = st.sidebar.file_uploader("Novedades RRHH", type=["xlsx"])
clg_novelties_file = st.sidebar.file_uploader("Novedades CLG", type=["xlsx"])
camplegal_file = st.sidebar.file_uploader("Camp Legal", type=["xlsx"])
smokeball_file = st.sidebar.file_uploader("Smokeball", type=["xlsx"])

st.sidebar.divider()

start_date = st.sidebar.date_input("Start Date", pd.Timestamp("2026-08-01"))
end_date = st.sidebar.date_input("End Date", pd.Timestamp("2026-08-31"))

# ==================================================
# PROCESS
# ==================================================

if (
    resources_file and
    toggl_file and
    novelties_file and
    camplegal_file and
    smokeball_file and
    clg_novelties_file
):

    daily_report, detail_report, users_summary, compliance_engine, compliance_summary = process_files(
        toggl_file,
        camplegal_file,
        smokeball_file,
        resources_file,
        novelties_file,
        clg_novelties_file,
        start_date,
        end_date
    )

    total_users = users_summary["USER_CORRECT"].nunique()

    # =========================================
    # CÁLCULO DE MÉTRICAS
    # =========================================
    compliant_days = len(
        compliance_engine[compliance_engine["Status"] == "✅ Cumple"]
    )

    non_compliant_days = len(
        compliance_engine[
            compliance_engine["Status"].isin([
                "❌ No registró tiempo",
                "❌ Horas insuficientes"
            ])
        ]
    )

    justified_days = len(
        compliance_engine[
            compliance_engine["Status"].astype(str).str.startswith("🟡")
        ]
    )

    total_hours = round(users_summary["Total_Hours"].sum(), 2)

    total_non_compliance = len(
        compliance_engine[
            compliance_engine["Status"].isin([
                "❌ No registró tiempo",
                "❌ Horas insuficientes"
            ])
        ]
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("👥 Users", total_users)
    c2.metric("✅ Compliant Days", compliant_days)
    c3.metric("❌ Non Compliant Days", non_compliant_days)
    c4.metric("🟡 Justified", justified_days)
    c5.metric("⏱️ Total Hours", total_hours)

    # =====================================
    # TABS
    # =====================================

    tab1, tab2, tab3, tab4 = st.tabs([
        "🚨 Compliance Engine",
        "📋 Activity Detail",
        "👥 Users Summary",
        "📊 Compliance Summary"
    ])

    # =====================================
    # TAB 1 - COMPLIANCE ENGINE
    # =====================================

    with tab1:
        st.subheader("Compliance Engine")

        compliance_filter = st.selectbox(
            "Compliance Status",
            ["All", "✅ Cumple", "❌ No registró tiempo", "❌ Horas insuficientes"]
        )

        compliance_users = sorted(compliance_engine["User"].dropna().unique().tolist())
        selected_compliance_user = st.selectbox(
            "Search User",
            ["All Users"] + compliance_users
        )

        engine = compliance_engine.copy()

        if compliance_filter != "All":
            engine = engine[engine["Status"] == compliance_filter]

        if selected_compliance_user != "All Users":
            engine = engine[engine["User"] == selected_compliance_user]

        st.dataframe(
            engine[["Date", "User", "Hours Worked", "Novelty", "Status"]],
            use_container_width=True
        )

    # =====================================
    # TAB 2 - ACTIVITY DETAIL
    # =====================================

    with tab2:
        st.subheader("Activity Detail")

        users_list = sorted(
            detail_report["USER_CORRECT"]
            .dropna()
            .unique()
            .tolist()
        )

        selected_user = st.selectbox(
            "Search User",
            ["All Users"] + users_list
        )

        detail_view = detail_report.copy()
        if selected_user != "All Users":
            detail_view = detail_view[
                detail_view["USER_CORRECT"] == selected_user
            ]

        st.dataframe(
            detail_view[["Date1", "USER_CORRECT", "Activity", "Hours", "Source"]],
            use_container_width=True
        )

    # =====================================
    # TAB 3 - USERS SUMMARY
    # =====================================

    with tab3:
        st.subheader("Total Hours by User")
        st.dataframe(
            users_summary[[
                "USER_CORRECT",
                "COMPANY",
                "DEPARTMENT",
                "TEAM",
                "Entries",
                "Total_Hours"
            ]],
            use_container_width=True
        )

        fig_users = px.bar(
            users_summary.head(25),
            x="USER_CORRECT",
            y="Total_Hours",
            title="Top Users by Hours"
        )
        st.plotly_chart(fig_users, use_container_width=True)

    # =====================================
    # TAB 4 - COMPLIANCE SUMMARY
    # =====================================

    with tab4:
        st.subheader("Compliance Summary By User")
        summary_view = compliance_summary.copy()

        summary_users = sorted(summary_view["User"].dropna().unique().tolist())
        selected_summary_user = st.selectbox(
            "Search User (Summary)",
            ["All Users"] + summary_users
        )

        if selected_summary_user != "All Users":
            summary_view = summary_view[summary_view["User"] == selected_summary_user]

        cols_per_row = 6

        for i in range(0, len(summary_view), cols_per_row):
            cols = st.columns(cols_per_row)

            for j in range(cols_per_row):
                if i + j >= len(summary_view):
                    continue

                row = summary_view.iloc[i + j]
                non_compliant = int(row["Non_Compliant_Days"])
                border_color = "#e74c3c" if non_compliant > 0 else "#27ae60"
                icon = "❌" if non_compliant > 0 else "✅"

                with cols[j]:
                    card_background = (
                        "#e8f5e9"
                        if non_compliant == 0
                        else "#fdeaea"
                    )
                    card_border = (
                        "#2e7d32"
                        if non_compliant == 0
                        else "#d32f2f"
                    )

                    st.markdown(
                        f"""
                        <div style="
                            background: {card_background};
                            border: 2px solid {card_border};
                            border-radius: 12px;
                            padding: 10px;
                            margin-bottom: 10px;
                        ">
                            <b>{icon} {row['User']}</b><br><br>
                            ⏱️ {row['Total_Hours']:.2f} hrs<br>
                            📅 {int(row['Days_Evaluated'])} días<br>
                            ✅ {int(row['Compliant_Days'])}<br>
                            ❌ {int(row['Non_Compliant_Days'])}<br>
                            🟡 {int(row['Justified_Days'])}<br>
                            <hr>
                            <b>Incidencias:</b><br>
                            {row['Failed_Dates']}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

    # =====================================
    # NON COMPLIANCE
    # =====================================

    st.divider()
    st.subheader("🚨 Non Compliant Records")

    non_compliance = compliance_engine[
        compliance_engine["Status"].isin([
            "❌ No registró tiempo",
            "❌ Horas insuficientes"
        ])
    ]

    if len(non_compliance) > 0:
        st.dataframe(
            non_compliance[["Date", "User", "Hours Worked", "Status"]],
            use_container_width=True
        )
    else:
        st.success("🎉 No non-compliant records found!")

else:
    st.info("📌 Upload all files to begin.")

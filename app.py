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
    for _, row in df_names.iterrows():
        correct_name = row.get("NAME CORRECT")
        tg_name = row.get("NAME TG")
        if pd.notna(correct_name) and pd.notna(tg_name):
            mappings.append({
                "SOURCE_NAME": normalize_name(tg_name),
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
def process_files(toggl_file, resources_file, novelties_file, start_date, end_date):
    
    # =========================================
    # READ FILES
    # =========================================
    
    df_toggl = pd.read_excel(toggl_file, sheet_name="DataBaseToggl", engine="openpyxl")
    df_names = pd.read_excel(resources_file, sheet_name="Names", engine="openpyxl")
    df_novelties = pd.read_excel(novelties_file, sheet_name="Novedades", engine="openpyxl")
    df_special_days = pd.read_excel(novelties_file, sheet_name="Novedades 2", engine="openpyxl")
    
    df_novelties["Fecha Inicio"] = pd.to_datetime(df_novelties["Fecha Inicio"], errors="coerce")
    df_novelties["Fecha Fin"] = pd.to_datetime(df_novelties["Fecha Fin"], errors="coerce")
    df_special_days["Fecha"] = pd.to_datetime(df_special_days["Fecha"], errors="coerce")
    
    # =========================================
    # USER MAP
    # =========================================
    
    user_map = build_user_mapping(df_names)
    
    # =========================================
    # NORMALIZE USERS
    # =========================================
    
    df_toggl["NORMALIZED_MEMBER"] = df_toggl["Member"].astype(str).apply(normalize_name)
    df_toggl["USER_CORRECT"] = df_toggl["NORMALIZED_MEMBER"].map(user_map).fillna(df_toggl["Member"])
    
    # =========================================
    # DATE
    # =========================================
    
    df_toggl["Date1"] = pd.to_datetime(df_toggl["Date1"], errors="coerce")
    
    # =========================================
    # FILTER DATES
    # =========================================
    
    df_toggl = df_toggl[
        (df_toggl["Date1"] >= pd.to_datetime(start_date)) &
        (df_toggl["Date1"] <= pd.to_datetime(end_date))
    ]
    
    # =========================================
    # HOURS
    # =========================================
    
    df_toggl["Hours"] = df_toggl["Duration"].apply(convert_duration_to_hours)
    
    # =========================================
    # DAILY REPORT
    # =========================================
    
    daily_report = (
        df_toggl
        .groupby(["Date1", "USER_CORRECT"], as_index=False)
        .agg(Total_Hours=("Hours", "sum"))
    )
    
    daily_report["Total_Hours"] = daily_report["Total_Hours"].round(2)
    daily_report["Status"] = np.where(
        daily_report["Total_Hours"] >= MINIMUM_DAILY_HOURS,
        "✅ Comply",
        "❌ Does Not Comply"
    )
    
    # =========================================
    # EMPLOYEE INFO
    # =========================================
    
    employee_info = (
        df_names[["NAME CORRECT", "COMPANY", "DEPARTMENT", "TEAM"]]
        .drop_duplicates()
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
        df_toggl
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
    
    active_users = df_names[
        df_names["USER STATUS"].astype(str).str.upper().eq("ACTIVE")
    ]["NAME CORRECT"].dropna().unique()
    
    all_dates = pd.date_range(start=start_date, end=end_date, freq="D")
    
    compliance_records = []
    
    for current_day in all_dates:
        weekday = current_day.weekday()
        
        for user in active_users:
            required_hours = 0
            
            # Lunes a Viernes
            if weekday <= 4:
                required_hours = 8
            
            # Sábado
            elif weekday == 5:
                saturday_user = df_special_days[
                    (df_special_days["Persona"] == user) &
                    (df_special_days["Tipo de Novedad"] == "Sábado") &
                    (df_special_days["Fecha"].dt.date == current_day.date())
                ]
                if len(saturday_user) > 0:
                    required_hours = 4
            
            # Domingo
            else:
                required_hours = 0
            
            # Festivos
            holiday_user = df_special_days[
                (df_special_days["Persona"] == user) &
                (df_special_days["Tipo de Novedad"] == "Festivo") &
                (df_special_days["Fecha"].dt.date == current_day.date())
            ]
            if len(holiday_user) > 0:
                required_hours = 8
            
            if required_hours == 0:
                continue
            
            day_record = daily_report[
                (daily_report["USER_CORRECT"] == user) &
                (daily_report["Date1"].dt.date == current_day.date())
            ]
            
            worked_hours = 0
            if len(day_record) > 0:
                worked_hours = day_record["Total_Hours"].sum()
            
            novelty = get_novelty_status(user, current_day, df_novelties)
            
            if novelty is not None:
                status = f"🟡 {novelty}"
            else:
                if worked_hours == 0:
                    status = "❌ No registró tiempo"
                elif worked_hours < required_hours:
                    status = "❌ Horas insuficientes"
                else:
                    status = "✅ Cumple"
            
            compliance_records.append({
                "Date": current_day.date(),
                "User": user,
                "Hours Worked": worked_hours,
                "Hours Required": required_hours,
                "Novelty": novelty,
                "Status": status
            })
    
    compliance_engine = pd.DataFrame(compliance_records)
    
    return daily_report, df_toggl, users_summary, compliance_engine


# ==================================================
# HEADER
# ==================================================

st.title("⏱️ Time Control Platform")
st.markdown("### Phase 1 - Toggl Validation")

# ==================================================
# SIDEBAR
# ==================================================

st.sidebar.header("Upload Files")

resources_file = st.sidebar.file_uploader("Power BI Resources", type=["xlsx"])
toggl_file = st.sidebar.file_uploader("Toggl File", type=["xlsx"])
novelties_file = st.sidebar.file_uploader("Novedades RRHH", type=["xlsx"])

st.sidebar.divider()

start_date = st.sidebar.date_input("Start Date", pd.Timestamp("2026-07-01"))
end_date = st.sidebar.date_input("End Date", pd.Timestamp("2026-07-31"))

# ==================================================
# PROCESS
# ==================================================

if resources_file and toggl_file and novelties_file:
    
    daily_report, detail_report, users_summary, compliance_engine = process_files(
        toggl_file,
        resources_file,
        novelties_file,
        start_date,
        end_date
    )
    
    total_users = users_summary["USER_CORRECT"].nunique()
    compliant_days = (daily_report["Status"] == "✅ Comply").sum()
    non_compliant_days = (daily_report["Status"] == "❌ Does Not Comply").sum()
    total_hours = round(users_summary["Total_Hours"].sum(), 2)
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Users", total_users)
    c2.metric("Compliant Days", compliant_days)
    c3.metric("Non Compliant Days", non_compliant_days)
    c4.metric("Total Hours", total_hours)
    
    # =====================================
    # TABS
    # =====================================
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "👥 Users Summary",
        "✅ Daily Compliance",
        "📋 Activity Detail",
        "🚨 Compliance Engine"
    ])
    
    # =====================================
    # TAB 1 - USERS SUMMARY
    # =====================================
    
    with tab1:
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
    # TAB 2 - DAILY COMPLIANCE
    # =====================================
    
    with tab2:
        status_filter = st.selectbox(
            "Status Filter",
            ["All", "✅ Comply", "❌ Does Not Comply"]
        )
        
        report = daily_report.copy()
        if status_filter != "All":
            report = report[report["Status"] == status_filter]
        
        st.dataframe(
            report[[
                "Date1",
                "USER_CORRECT",
                "COMPANY",
                "TEAM",
                "Total_Hours",
                "Status"
            ]],
            use_container_width=True
        )
    
    # =====================================
    # TAB 3 - ACTIVITY DETAIL
    # =====================================
    
    with tab3:
        st.subheader("Activity Detail")
        st.dataframe(
            detail_report[[
                "Date1",
                "USER_CORRECT",
                "Project (Activity)",
                "Duration",
                "Hours"
            ]],
            use_container_width=True
        )
    
    # =====================================
    # TAB 4 - COMPLIANCE ENGINE
    # =====================================
    
    with tab4:
        st.subheader("Compliance Engine")
        
        compliance_filter = st.selectbox(
            "Compliance Status",
            ["All", "✅ Cumple", "❌ No registró tiempo", "❌ Horas insuficientes"]
        )
        
        engine = compliance_engine.copy()
        if compliance_filter != "All":
            engine = engine[engine["Status"] == compliance_filter]
        
        st.dataframe(engine, use_container_width=True)
    
    # =====================================
    # NON COMPLIANCE
    # =====================================
    
    st.divider()
    st.subheader("🚨 Non Compliant Records")
    
    non_compliance = daily_report[daily_report["Status"] == "❌ Does Not Comply"]
    
    st.dataframe(
        non_compliance[[
            "Date1",
            "USER_CORRECT",
            "COMPANY",
            "TEAM",
            "Total_Hours"
        ]],
        use_container_width=True
    )

else:
    st.info("Upload both files to begin.")

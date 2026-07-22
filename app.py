import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import unicodedata

# =====================================================
# CONFIGURACION
# =====================================================

st.set_page_config(
    page_title="Time Control Platform",
    page_icon="⏱️",
    layout="wide"
)

MINIMUM_DAILY_HOURS = 8

# =====================================================
# FUNCIONES
# =====================================================

def normalize_name(text):
    """
    Normaliza nombres para hacer matching:
    - Mayúsculas
    - Sin tildes
    - Sin espacios dobles
    """

    if pd.isna(text):
        return ""

    text = str(text).strip().upper()

    text = unicodedata.normalize("NFD", text)
    text = "".join(
        c for c in text
        if unicodedata.category(c) != "Mn"
    )

    text = " ".join(text.split())

    return text


def build_user_mapping(df_names):
    """
    Crea diccionario de homologación usando NAME TG
    """

    mappings = []

    for _, row in df_names.iterrows():

        correct_name = row.get("NAME CORRECT")

        if pd.isna(correct_name):
            continue

        tg_name = row.get("NAME TG")

        if pd.notna(tg_name):

            mappings.append(
                {
                    "SOURCE_NAME": normalize_name(tg_name),
                    "NAME_CORRECT": correct_name
                }
            )

    map_df = pd.DataFrame(mappings)

    if len(map_df) == 0:
        return {}

    map_df = map_df.drop_duplicates()

    return dict(
        zip(
            map_df["SOURCE_NAME"],
            map_df["NAME_CORRECT"]
        )
    )


def convert_duration_to_hours(value):
    """
    Convierte la columna Dur de Excel a horas.
    Dur llega como fracción de día.
    """

    try:
        return float(value) * 24
    except:
        return 0


@st.cache_data
def process_files(
    toggl_file,
    resources_file,
    start_date,
    end_date
):
    # =====================================================
    # LEER ARCHIVOS
    # =====================================================

    df_toggl = pd.read_excel(
        toggl_file,
        sheet_name="DataBaseToggl",
        engine="openpyxl"
    )

    df_names = pd.read_excel(
        resources_file,
        sheet_name="Names",
        engine="openpyxl"
    )

    # =====================================================
    # MAPEO DE USUARIOS
    # =====================================================

    user_map = build_user_mapping(df_names)

    # =====================================================
    # NORMALIZAR NOMBRES
    # =====================================================

    df_toggl["NORMALIZED_MEMBER"] = (
        df_toggl["Member"]
        .astype(str)
        .apply(normalize_name)
    )

    df_toggl["USER_CORRECT"] = (
        df_toggl["NORMALIZED_MEMBER"]
        .map(user_map)
        .fillna(df_toggl["Member"])
    )

    # =====================================================
    # FECHA
    # =====================================================

    df_toggl["Date1"] = pd.to_datetime(
        df_toggl["Date1"],
        errors="coerce"
    )

    # =====================================================
    # FILTRO FECHAS
    # =====================================================

    df_toggl = df_toggl[
        (df_toggl["Date1"] >= pd.to_datetime(start_date))
        &
        (df_toggl["Date1"] <= pd.to_datetime(end_date))
    ]

    # =====================================================
    # HORAS
    # =====================================================

    df_toggl["Hours"] = (
        df_toggl["Dur"]
        .apply(convert_duration_to_hours)
    )

    # =====================================================
    # RESUMEN DIARIO
    # =====================================================

    daily_report = (
        df_toggl
        .groupby(
            ["Date1", "USER_CORRECT"],
            as_index=False
        )
        .agg(
            Total_Hours=("Hours", "sum")
        )
    )

    # =====================================================
    # ESTADO CUMPLIMIENTO
    # =====================================================

    daily_report["Status"] = np.where(
        daily_report["Total_Hours"] >= MINIMUM_DAILY_HOURS,
        "✅ Comply",
        "❌ Does Not Comply"
    )

    # =====================================================
    # INFORMACION ORGANIZACIONAL
    # =====================================================

    master = (
        df_names[
            [
                "NAME CORRECT",
                "COMPANY",
                "DEPARTMENT",
                "TEAM"
            ]
        ]
        .drop_duplicates()
    )

    daily_report = daily_report.merge(
        master,
        left_on="USER_CORRECT",
        right_on="NAME CORRECT",
        how="left"
    )

    # =====================================================
    # DETALLE ACTIVIDADES
    # =====================================================

    detail_report = df_toggl.copy()

    return daily_report, detail_report


# =====================================================
# HEADER
# =====================================================

st.title("⏱️ Time Control Platform")
st.markdown(
    "### Fase 1 - Validación de tiempos Toggl"
)

# =====================================================
# SIDEBAR
# =====================================================

st.sidebar.header("Archivos")

resources_file = st.sidebar.file_uploader(
    "Power BI Resources",
    type=["xlsx"]
)

toggl_file = st.sidebar.file_uploader(
    "Archivo Toggl",
    type=["xlsx"]
)

st.sidebar.markdown("---")

start_date = st.sidebar.date_input(
    "Fecha Inicial",
    pd.Timestamp("2026-07-01")
)

end_date = st.sidebar.date_input(
    "Fecha Final",
    pd.Timestamp("2026-07-31")
)

# =====================================================
# CARGA DE DATOS
# =====================================================

if resources_file and toggl_file:

    daily_report, detail_report = process_files(
        toggl_file,
        resources_file,
        start_date,
        end_date
    )

    # =====================================================
    # KPIs
    # =====================================================

    total_users = (
        daily_report["USER_CORRECT"]
        .nunique()
    )

    compliant_users = (
        daily_report[
            daily_report["Status"] == "✅ Comply"
        ]["USER_CORRECT"]
        .nunique()
    )

    non_compliant_users = (
        daily_report[
            daily_report["Status"] == "❌ Does Not Comply"
        ]["USER_CORRECT"]
        .nunique()
    )

    total_hours = round(
        daily_report["Total_Hours"].sum(),
        2
    )

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Usuarios",
        total_users
    )

    col2.metric(
        "Cumplen",
        compliant_users
    )

    col3.metric(
        "No Cumplen",
        non_compliant_users
    )

    col4.metric(
        "Horas Totales",
        total_hours
    )

    st.divider()

    # =====================================================
    # TABS
    # =====================================================

    tab1, tab2 = st.tabs(
        [
            "✅ Daily Compliance",
            "📋 Activity Detail"
        ]
    )

    # =====================================================
    # TAB 1
    # =====================================================

    with tab1:

        st.subheader(
            "Resumen Diario de Cumplimiento"
        )

        status_filter = st.selectbox(
            "Filtrar Estado",
            [
                "Todos",
                "✅ Comply",
                "❌ Does Not Comply"
            ]
        )

        display_df = daily_report.copy()

        if status_filter != "Todos":

            display_df = display_df[
                display_df["Status"] == status_filter
            ]

        # =====================================================
        # GRAFICO
        # =====================================================

        chart_df = (
            display_df
            .groupby("USER_CORRECT")["Total_Hours"]
            .sum()
            .reset_index()
            .sort_values(
                "Total_Hours",
                ascending=False
            )
            .head(20)
        )

        if len(chart_df) > 0:

            fig = px.bar(
                chart_df,
                x="USER_CORRECT",
                y="Total_Hours",
                title="Horas registradas por usuario"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

        # =====================================================
        # TABLA RESUMEN
        # =====================================================

        display_df["Date1"] = (
            display_df["Date1"]
            .dt.strftime("%Y-%m-%d")
        )

        st.dataframe(
            display_df[
                [
                    "Date1",
                    "USER_CORRECT",
                    "COMPANY",
                    "DEPARTMENT",
                    "TEAM",
                    "Total_Hours",
                    "Status"
                ]
            ],
            use_container_width=True
        )

        csv = display_df.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            label="⬇ Descargar Resumen",
            data=csv,
            file_name="daily_report.csv",
            mime="text/csv"
        )

    # =====================================================
    # TAB 2
    # =====================================================

    with tab2:

        st.subheader(
            "Detalle de Actividades"
        )

        detail_view = detail_report.copy()

        detail_view["Date1"] = (
            pd.to_datetime(
                detail_view["Date1"]
            )
            .dt.strftime("%Y-%m-%d")
        )

        st.dataframe(
            detail_view[
                [
                    "Date1",
                    "USER_CORRECT",
                    "Project (Activity)",
                    "Hours"
                ]
            ],
            use_container_width=True
        )

        detail_csv = detail_view.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            label="⬇ Descargar Detalle",
            data=detail_csv,
            file_name="activity_detail.csv",
            mime="text/csv"
        )

    # =====================================================
    # INCUMPLIMIENTOS
    # =====================================================

    st.divider()

    st.subheader(
        "🚨 Usuarios que NO Cumplen"
    )

    non_compliance = daily_report[
        daily_report["Status"]
        == "❌ Does Not Comply"
    ]

    st.dataframe(
        non_compliance[
            [
                "Date1",
                "USER_CORRECT",
                "COMPANY",
                "TEAM",
                "Total_Hours"
            ]
        ],
        use_container_width=True
    )

else:

    st.info(
        "Cargue el archivo de Toggl y el archivo Power BI Resources para comenzar."
    )

# app.py
import os
import base64
import json
import zlib
import unicodedata
from datetime import datetime, date

from io import BytesIO  # ✅ ADDED (Excel download)

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path

import streamlit.components.v1 as components
import json

from pandas.errors import DatabaseError as PandasDatabaseError

from datetime import datetime
import re
import pyodbc
from functools import lru_cache

# -------------------------------
# CONFIG
# -------------------------------
st.set_page_config(
    page_title="Ventas ExpertCell",
    page_icon="📊",
    layout="wide",
)

# -------------------------------
# GLOBAL CONSTANTS
# -------------------------------
EXCLUDED_VENDOR = "ABASTECEDORA Y SUMINISTROS ORTEGA/ISABEL VALDEZ JIMENEZ"

# -------------------------------
# THEME (READ ONLY — we do NOT force anything)
# -------------------------------
try:
    theme_base = st.get_option("theme.base") or "light"  # "light" | "dark"
except Exception:
    theme_base = "light"

IS_DARK = str(theme_base).lower() == "dark"
PLOTLY_TEMPLATE = "plotly_dark" if IS_DARK else "plotly_white"
# ✅ Pylance safe default (Streamlit can stop early with st.stop(), but Pylance doesn't know that)
center_sel: list[str] = ["CC2", "JV"]

# -------------------------------
# NEUTRAL, THEME-FRIENDLY CSS (no forced colors)
# -------------------------------
st.markdown(
    """
<style>
header[data-testid="stHeader"]{ background: rgba(0,0,0,0) !important; }
header[data-testid="stHeader"] [data-testid="stToolbar"]{ background: rgba(0,0,0,0) !important; }
header[data-testid="stHeader"] button,
header[data-testid="stHeader"] svg{
  color: var(--text-color) !important;
  fill: var(--text-color) !important;
}

.stApp{
  background-color: var(--background-color) !important;
  background-image:
    radial-gradient(circle at 1px 1px, rgba(127,127,127,0.14) 1px, transparent 0) !important;
  background-size: 18px 18px !important;
  color: var(--text-color) !important;
}
.block-container{ padding-top: 1.2rem; }

section[data-testid="stSidebar"]{
  background: var(--secondary-background-color) !important;
  border-right: 1px solid rgba(127,127,127,0.25) !important;
}
section[data-testid="stSidebar"] *{ color: var(--text-color) !important; }

section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea{
  background: var(--background-color) !important;
  border: 1px solid rgba(127,127,127,0.28) !important;
  color: var(--text-color) !important;
  border-radius: 10px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] > div{
  background: var(--background-color) !important;
  border: 1px solid rgba(127,127,127,0.28) !important;
  border-radius: 10px !important;
}
section[data-testid="stSidebar"] [data-baseweb="tag"]{
  background: rgba(127,127,127,0.25) !important;
  color: var(--text-color) !important;
  border-radius: 999px !important;
  font-weight: 800 !important;
}

.metric-card{
  background: var(--secondary-background-color) !important;
  border-radius: 14px;
  padding: 14px 16px;
  border: 1px solid rgba(127,127,127,0.22);
  box-shadow: 0 1px 0 rgba(0,0,0,0.10);
}
.metric-label{ font-size:0.92rem; opacity:0.78; font-weight:800; }
.metric-value{ font-size:2.25rem; font-weight:900; margin-top:4px; line-height:1; }
.metric-sub{ font-size:0.9rem; opacity:0.70; margin-top:6px; }

.kpi-mini{
  background: var(--secondary-background-color) !important;
  border-radius: 14px;
  padding: 12px 14px;
  border: 1px solid rgba(127,127,127,0.22);
}
.kpi-mini .t{ font-size:0.9rem; font-weight:900; opacity:0.78; }
.kpi-mini .v{ font-size:1.6rem; font-weight:900; margin-top:4px; }

div[data-testid="stPlotlyChart"] > div{ border-radius: 16px; }
</style>
""",
    unsafe_allow_html=True,
)

# -------------------------------
# HELPERS (format)
# -------------------------------
def fmt_int(x: float | int | None) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "-"
    return f"{int(round(x)):,}"


def fmt_money_short(x: float | int | None) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "$-"
    x = float(x)
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1_000_000:
        return f"{sign}${x/1_000_000:.2f}M"
    if x >= 1_000:
        return f"{sign}${x/1_000:.2f}K"
    return f"{sign}${x:,.2f}"


def fmt_pct(x: float | None) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "-"
    return f"{x*100:.2f}%"


def fmt_pct(x: float | None) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "-"
    return f"{x*100:.2f}%"


@lru_cache(maxsize=50000)
def normalize_name(s: str) -> str:
    if s is None:
        s = ""
    s = str(s).strip().upper()
    s = " ".join(s.split())
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s


def metric_card(label: str, value: str, sub: str | None = None):
    sub_html = f'<div class="metric-sub">{sub}</div>' if sub else ""
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_mini(label: str, value: str):
    st.markdown(
        f"""
        <div class="kpi-mini">
          <div class="t">{label}</div>
          <div class="v">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def month_key_to_name_es(ym: int) -> str:
    y = ym // 100
    m = ym % 100
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    if 1 <= m <= 12:
        return f"{meses[m-1]} {y}"
    return str(ym)


def add_bar_value_labels(fig: go.Figure) -> go.Figure:
    """
    Adds value labels to BAR traces ONLY when they don't already have text/texttemplate.
    Keeps existing charts unchanged if they already define labels.
    """
    try:
        for tr in getattr(fig, "data", []) or []:
            if getattr(tr, "type", "") != "bar":
                continue

            # If the bar already has labels, don't touch it
            has_text = tr.text is not None and np.size(tr.text) > 0
            has_template = bool(getattr(tr, "texttemplate", "") or "")
            if has_text or has_template:
                continue

            orient = getattr(tr, "orientation", None) or "v"
            if orient == "h":
                tr.update(texttemplate="%{x:,.0f}", textposition="outside", cliponaxis=False)
            else:
                tr.update(texttemplate="%{y:,.0f}", textposition="outside", cliponaxis=False)

        fig.update_layout(uniformtext_minsize=10, uniformtext_mode="hide")
    except Exception:
        pass
    return fig


def apply_plotly_theme(fig: go.Figure) -> go.Figure:
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    add_bar_value_labels(fig)
    return fig


# ✅ Totals row helper (for every table shown)
def add_totals_row(
    df: pd.DataFrame,
    label_col: str,
    totals: dict,
    label: str = "TOTAL",
) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    row = {c: np.nan for c in out.columns}
    if label_col in row:
        row[label_col] = label
    for k, v in totals.items():
        if k in row:
            row[k] = v
    return pd.concat([out, pd.DataFrame([row])], ignore_index=True)


# ✅ NEW: bold totals rows in st.dataframe using Styler
def style_totals_bold(df: pd.DataFrame, label_col: str):
    def _bold_row(row):
        v = row.get(label_col, "")
        if "TOTAL" in str(v).upper():
            return ["font-weight: 800"] * len(row)
        return [""] * len(row)

    return df.style.apply(_bold_row, axis=1)

# -------------------------------
# ✅ DATE & WORKABLE DAYS HELPERS (add these to your HELPERS section)
# -------------------------------
def _nth_weekday_of_month(year: int, month: int, weekday: int, n: int) -> date:
    d = date(year, month, 1)
    shift = (weekday - d.weekday()) % 7
    d = d.replace(day=1 + shift)
    d = d.replace(day=d.day + 7 * (n - 1))
    return d

def mexico_puentes(year: int) -> set[date]:
    """Returns a set of non-working holidays for Mexico."""
    return {
        date(year, 1, 1),             # Año Nuevo
        _nth_weekday_of_month(year, 2, 0, 1), # Constitución (1er lunes feb)
        _nth_weekday_of_month(year, 3, 0, 3), # Benito Juárez (3er lunes mar)
        date(year, 5, 1),             # Día del Trabajo
        date(year, 9, 16),            # Independencia
        _nth_weekday_of_month(year, 11, 0, 3), # Revolución (3er lunes nov)
        date(year, 12, 25),           # Navidad
    }

def workable_equiv_between(start_d: date, end_d: date) -> float:
    """Calculates effective working days: Mon-Fri=1.0, Sat=0.5, Sun=0, Holidays=0."""
    if end_d < start_d:
        return 0.0
    days = pd.date_range(pd.Timestamp(start_d), pd.Timestamp(end_d), freq="D")
    puentes = mexico_puentes(start_d.year)
    total = 0.0
    for dts in days:
        d = dts.date()
        if dts.year != start_d.year:
            puentes = puentes | mexico_puentes(dts.year)
        
        if d in puentes:
            continue
            
        wd = dts.weekday() # 0=Mon, 6=Sun
        if wd <= 4:    # Mon-Fri
            total += 1.0
        elif wd == 5:  # Sat
            total += 0.5
        # Sun = 0
    return float(total)

def month_bounds(ym_int: int) -> tuple[date, date]:
    """Returns the first and last date of a YYYYMM integer."""
    y = ym_int // 100
    m = ym_int % 100
    start = date(y, m, 1)
    end = (pd.Timestamp(y, m, 1) + pd.offsets.MonthEnd(1)).date()
    return start, end

# -------------------------------
# ✅ EXCEL EXPORT HELPERS (ADDED)
# -------------------------------
def _safe_sheet_name(name: str) -> str:
    name = str(name or "Sheet").strip()
    bad = [":", "\\", "/", "?", "*", "[", "]"]
    for b in bad:
        name = name.replace(b, " ")
    name = " ".join(name.split())
    return (name[:31] or "Sheet")


def build_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets.items():
            if df is None:
                continue
            df_to_write = df.copy()
            sheet = _safe_sheet_name(sname)
            df_to_write.to_excel(writer, index=False, sheet_name=sheet)

            try:
                from openpyxl.utils import get_column_letter
                ws = writer.sheets[sheet]
                max_rows = min(len(df_to_write), 500)
                for i, col in enumerate(df_to_write.columns, start=1):
                    col_vals = [str(col)]
                    if max_rows > 0:
                        col_vals += [str(v) for v in df_to_write[col].iloc[:max_rows].tolist()]
                    width = min(max(len(x) for x in col_vals) + 2, 55)
                    ws.column_dimensions[get_column_letter(i)].width = max(10, width)
            except Exception:
                pass

    return output.getvalue()


# -------------------------------
# DB (SQL Server via pyodbc)
# -------------------------------
def get_db_cfg():
    if "db" in st.secrets:
        return {
            "server": st.secrets["db"]["server"],
            "database": st.secrets["db"]["database"],
            "username": st.secrets["db"]["username"],
            "password": st.secrets["db"]["password"],
            "driver": st.secrets["db"].get("driver", "ODBC Driver 17 for SQL Server"),
        }
    return {
        "server": os.getenv("DB_SERVER", ""),
        "database": os.getenv("DB_DATABASE", ""),
        "username": os.getenv("DB_USERNAME", ""),
        "password": os.getenv("DB_PASSWORD", ""),
        "driver": os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server"),
    }


import pyodbc

@st.cache_resource(show_spinner=False)
def _get_conn():
    cfg = get_db_cfg()
    conn_str = (
        f"DRIVER={{{cfg['driver']}}};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};"
        f"PWD={cfg['password']};"
        f"TrustServerCertificate=yes;"
        f"Mars_Connection=yes;"
    )
    return pyodbc.connect(conn_str, autocommit=True)

@st.cache_data(ttl=600, show_spinner=False)
def read_sql(query: str) -> pd.DataFrame:
    conn = None
    try:
        conn = _get_conn()
        return pd.read_sql(query, conn)

    # ✅ pandas wraps DB/ODBC errors here sometimes
    except (pyodbc.Error, PandasDatabaseError):
        # rebuild connection once
        try:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
        finally:
            st.cache_resource.clear()

        conn2 = _get_conn()
        return pd.read_sql(query, conn2)

# -------------------------------
# POWER QUERY → PANDAS (Consulta2)
# -------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def load_empleados() -> pd.DataFrame:
    q = r"""
    SELECT
      Region,
      Plaza,
      Tienda AS Centro,
      [Nombre Completo] AS Nombre,
      Puesto,
      RFC,
      [Jefe Inmediato],
      Estatus,
      [Fecha Ingreso],
      [Fecha Baja],
      [Canal de Venta],
      Operacion
    FROM reporte_empleado('EMPRESA_MAESTRA',1,'','') AS e
    WHERE
      [Canal de Venta] IN ('ATT')
      AND [Operacion] IN ('CONTACT CENTER')
      AND [Tipo Tienda] IN ('VIRTUAL')
      AND (
        Estatus = 'ACTIVO'
        OR (
          Estatus = 'BAJA'
          AND [Fecha Baja] >= DATEADD(MONTH, -1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
          AND [Fecha Baja] <  DATEADD(MONTH,  1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
        )
      )
    """
    df = read_sql(q)
    df["Nombre"] = df["Nombre"].astype(str).str.strip()
    df["Jefe Inmediato"] = df["Jefe Inmediato"].astype(str).str.strip()
    df["Estatus"] = df["Estatus"].astype(str).str.strip()
    df["Centro"] = df["Centro"].astype(str).str.strip()
    df["Puesto"] = df["Puesto"].astype(str).str.strip()
    return df


# -------------------------------
# POWER QUERY → PANDAS (Consulta1)
# -------------------------------
def build_ventas_query(start_yyyymmdd: str, end_yyyymmdd: str) -> str:
    return f"""
    SELECT
      FOLIO,
      [PTO. DE VENTA] AS CENTRO,
      [OPERACION PDV],
      [ESTATUS],
      [EJECUTIVO],
      [FECHA DE CAPTURA],
      [PLAN],
      [RENTA SIN IMPUESTOS],
      [PRECIO],
      [SUBREGION]
    FROM reporte_ventas_no_conciliadas('EMPRESA_MAESTRA', 4, '{start_yyyymmdd}', '{end_yyyymmdd}', 1, '19000101', '20990101')
    WHERE
      [OPERACION PDV] = 'CONTACT CENTER'
      AND [PTO. DE VENTA] LIKE 'EXP ATT C CENTER%'
    """


@st.cache_data(ttl=600, show_spinner=False)
def load_ventas(start_yyyymmdd: str, end_yyyymmdd: str) -> pd.DataFrame:
    q = build_ventas_query(start_yyyymmdd, end_yyyymmdd)
    df = read_sql(q)

    df["EJECUTIVO"] = df["EJECUTIVO"].astype(str).str.strip()
    df["CENTRO"] = df["CENTRO"].astype(str).str.strip()

    c = df["CENTRO"].astype(str).str.upper()
    df["CENTRO"] = np.where(
        c.str.contains("JUAREZ", na=False),
        "EXP ATT C CENTER JUAREZ",
        np.where(c.str.contains("CENTER 2", na=False), "EXP ATT C CENTER 2", df["CENTRO"])
    )

    df["EJECUTIVO"] = df["EJECUTIVO"].replace(
        {
            "CESAR JAHACIEL ALONSO GARCIAA": "CESAR JAHACIEL ALONSO GARCIA",
            "VICTOR BETANZO FUENTES": "VICTOR BETANZOS FUENTES",
        }
    )

    df["FECHA DE CAPTURA"] = pd.to_datetime(df["FECHA DE CAPTURA"], errors="coerce")
    df["Fecha"] = df["FECHA DE CAPTURA"].dt.date
    df["Hora"] = df["FECHA DE CAPTURA"].dt.time
    df["Año"] = df["FECHA DE CAPTURA"].dt.year
    df["Mes"] = df["FECHA DE CAPTURA"].dt.month
    df["NombreMes"] = df["FECHA DE CAPTURA"].dt.strftime("%B").str.lower()
    df["AñoMes"] = df["Año"] * 100 + df["Mes"]

    iso = df["FECHA DE CAPTURA"].dt.isocalendar()
    df["ISOYear"] = iso.year.astype(int)
    df["ISOWeek"] = iso.week.astype(int)
    df["SemanaAño"] = df["ISOWeek"]
    df["WeekKey"] = df["ISOYear"] * 100 + df["ISOWeek"]
    df["SemanaISO"] = df["ISOYear"].astype(str) + "-W" + df["ISOWeek"].astype(str).str.zfill(2)

    df["DiaSemana"] = df["FECHA DE CAPTURA"].dt.day_name().astype(str)
    df["DiaNum"] = df["FECHA DE CAPTURA"].dt.day.astype(int)

    for col in ["PRECIO", "RENTA SIN IMPUESTOS"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["CentroKey"] = np.where(df["CENTRO"].str.upper().str.contains("JUAREZ", na=False), "JV", "CC2")

    # ✅ ONLY ADD PHONE HERE
    try:
        q_tel = f"""
        SELECT
          LTRIM(RTRIM(CAST([Folio] AS VARCHAR(100)))) AS Folio_key,
          MAX(NULLIF(LTRIM(RTRIM(CAST([Telefono] AS VARCHAR(100)))), '')) AS [Telefono cliente]
        FROM reporte_programacion_entrega('empresa_maestra', 4, '{start_yyyymmdd}', '{end_yyyymmdd}')
        WHERE [Tienda solicita] LIKE 'EXP ATT C CENTER%'
        GROUP BY LTRIM(RTRIM(CAST([Folio] AS VARCHAR(100))))
        """

        df_tel = read_sql(q_tel)

        df["_folio_key"] = (
            df["FOLIO"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

        df["Telefono cliente"] = ""

        if df_tel is not None and not df_tel.empty:
            df_tel["Folio_key"] = (
                df_tel["Folio_key"]
                .astype(str)
                .str.strip()
                .str.replace(r"\.0$", "", regex=True)
            )

            df_tel["Telefono cliente"] = (
                df_tel["Telefono cliente"]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            tel_map = (
                df_tel[df_tel["Telefono cliente"].ne("")]
                .drop_duplicates(subset=["Folio_key"], keep="first")
                .set_index("Folio_key")["Telefono cliente"]
                .to_dict()
            )

            df["Telefono cliente"] = df["_folio_key"].map(tel_map).fillna("")

        # ✅ tiny fallback ONLY for remaining blanks in TERMINADA
        faltantes = (
            df.loc[
                df["Telefono cliente"].astype(str).str.strip().eq("")
                & df["ESTATUS"].astype(str).str.upper().str.contains("TERMINAD", na=False),
                "_folio_key"
            ]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )

        # safeguard: only run fallback when the list is small
        if 0 < len(faltantes) <= 200:
            dt_ini = pd.to_datetime(start_yyyymmdd, format="%Y%m%d", errors="coerce")
            tel_ini_fb = (dt_ini - pd.Timedelta(days=365)).strftime("%Y%m%d") if pd.notna(dt_ini) else start_yyyymmdd

            in_sql = ", ".join("'" + str(f).replace("'", "''") + "'" for f in faltantes)

            q_tel_fb = f"""
            SELECT
              LTRIM(RTRIM(CAST([Folio] AS VARCHAR(100)))) AS Folio_key,
              MAX(NULLIF(LTRIM(RTRIM(CAST([Telefono] AS VARCHAR(100)))), '')) AS [Telefono cliente]
            FROM reporte_programacion_entrega('empresa_maestra', 4, '{tel_ini_fb}', '{end_yyyymmdd}')
            WHERE
              [Tienda solicita] LIKE 'EXP ATT C CENTER%'
              AND LTRIM(RTRIM(CAST([Folio] AS VARCHAR(100)))) IN ({in_sql})
            GROUP BY LTRIM(RTRIM(CAST([Folio] AS VARCHAR(100))))
            """

            df_tel_fb = read_sql(q_tel_fb)

            if df_tel_fb is not None and not df_tel_fb.empty:
                df_tel_fb["Folio_key"] = (
                    df_tel_fb["Folio_key"]
                    .astype(str)
                    .str.strip()
                    .str.replace(r"\.0$", "", regex=True)
                )

                df_tel_fb["Telefono cliente"] = (
                    df_tel_fb["Telefono cliente"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                tel_map_fb = (
                    df_tel_fb[df_tel_fb["Telefono cliente"].ne("")]
                    .drop_duplicates(subset=["Folio_key"], keep="first")
                    .set_index("Folio_key")["Telefono cliente"]
                    .to_dict()
                )

                mask_fill = (
                    df["Telefono cliente"].astype(str).str.strip().eq("")
                    & df["ESTATUS"].astype(str).str.upper().str.contains("TERMINAD", na=False)
                )

                df.loc[mask_fill, "Telefono cliente"] = (
                    df.loc[mask_fill, "_folio_key"].map(tel_map_fb).fillna("")
                )

        df.drop(columns=["_folio_key"], inplace=True, errors="ignore")

    except Exception:
        df["Telefono cliente"] = ""

    return df


def add_empleado_join(ventas: pd.DataFrame, empleados: pd.DataFrame) -> pd.DataFrame:
    emp = empleados[["Nombre", "Jefe Inmediato"]].copy()
    emp["Nombre"] = emp["Nombre"].astype(str).str.strip()

    out = ventas.merge(emp, left_on="EJECUTIVO", right_on="Nombre", how="left")
    out.rename(columns={"Jefe Inmediato": "Supervisor"}, inplace=True)
    out["Supervisor"] = out["Supervisor"].fillna("").replace({"": "BAJA"})
    out["Supervisor_norm"] = out["Supervisor"].apply(normalize_name)
    out["EJECUTIVO_norm"] = out["EJECUTIVO"].apply(normalize_name)
    return out




# -------------------------------
# METAS (EXCEL POR MES) ✅ generic monthly metas logic
# -------------------------------
MESES_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

def month_name_es_only(ym_int: int) -> str:
    try:
        m = int(ym_int) % 100
        return MESES_ES.get(m, "")
    except Exception:
        return ""

def _metas_norm_txt(x: str) -> str:
    x = str(x or "").strip().lower()
    x = unicodedata.normalize("NFKD", x)
    x = "".join(c for c in x if not unicodedata.combining(c))
    x = x.replace("_", " ").replace("-", " ")
    x = " ".join(x.split())
    return x

def _metas_pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm_map = {_metas_norm_txt(c): c for c in df.columns}

    # exact match
    for cand in candidates:
        k = _metas_norm_txt(cand)
        if k in norm_map:
            return norm_map[k]

    # contains match
    for cand in candidates:
        k = _metas_norm_txt(cand)
        for nk, real in norm_map.items():
            if k and k in nk:
                return real

    return None

def _metas_pick_meta_month_col(df: pd.DataFrame, ym_int: int) -> str | None:
    mes = month_name_es_only(ym_int)
    if not mes:
        return None

    # Prefer explicit month column
    col = _metas_pick_col(
        df,
        [
            f"Meta {mes}",
            f"META {mes.upper()}",
            f"Meta_{mes}",
            f"meta {mes}",
            f"meta_{mes}",
        ],
    )
    if col:
        return col

    # Fallbacks if the workbook uses a generic "current month" style column
    col = _metas_pick_col(df, ["meta_mes_actual", "meta mes actual", "meta mes"])
    if col:
        return col

    # Last resort: any column containing both "meta" and current month name
    for c in df.columns:
        nc = _metas_norm_txt(c)
        if ("meta" in nc) and (mes in nc):
            return c

    # Ultimate fallback: single generic meta column
    generic = _metas_pick_col(df, ["Meta", "META", "meta"])
    if generic:
        return generic

    return None

def _metas_find_month_excel(ym_int: int) -> Path:
    base_dir = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    search_dirs = []
    for d in [base_dir, Path.cwd()]:
        if d not in search_dirs and d.exists():
            search_dirs.append(d)

    mes = month_name_es_only(ym_int)
    year = int(ym_int) // 100

    all_meta_files: list[Path] = []
    for d in search_dirs:
        for f in d.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() not in (".xlsx", ".xls"):
                continue
            nm = _metas_norm_txt(f.name)
            if "meta" in nm:
                all_meta_files.append(f)

    if not all_meta_files:
        raise FileNotFoundError(
            f"No encontré archivos Excel de metas en: {base_dir}\n"
            f"Coloca el archivo junto a app.py."
        )

    # Prefer files containing the selected month in the filename
    month_specific = [f for f in all_meta_files if mes and mes in _metas_norm_txt(f.name)]
    candidates = month_specific if month_specific else all_meta_files

    def _rank(p: Path) -> tuple[int, int, str]:
        nm = _metas_norm_txt(p.name)
        score = 0
        if "metas" in nm:
            score += 3
        if mes and mes in nm:
            score += 5
        if str(year) in nm:
            score += 1
        if "cc" in nm:
            score += 2
        return (score, -len(nm), p.name.lower())

    candidates = sorted(candidates, key=_rank, reverse=True)
    return candidates[0]

@st.cache_data(ttl=600, show_spinner=False)
def _load_metas_excel_df(ym_int: int) -> pd.DataFrame:
    """
    Loads metas Excel for the selected month and auto-selects the sheet
    that contains EJECUTIVO.
    """
    meta_path = _metas_find_month_excel(ym_int)

    xl = pd.ExcelFile(meta_path)
    best_sheet = None

    # Pick the sheet that has EJECUTIVO in headers
    for sn in xl.sheet_names:
        cols = list(pd.read_excel(meta_path, sheet_name=sn, nrows=0).columns)
        if any(_metas_norm_txt(c) == "ejecutivo" for c in cols):
            best_sheet = sn
            break

    # fallback: a sheet name that looks like metas_mes
    if best_sheet is None:
        for sn in xl.sheet_names:
            snn = _metas_norm_txt(sn)
            if "metas" in snn and "mes" in snn:
                best_sheet = sn
                break

    if best_sheet is None:
        best_sheet = xl.sheet_names[0]

    df = pd.read_excel(meta_path, sheet_name=best_sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# -------------------------------
# METAS (MANUAL TABLE)
# -------------------------------
# ✅ EDIT METAS HERE (by hand)
METAS_MANUAL_ROWS = [
    # --- Metas Centro ---
    {"IDCenter": "CC1", "Nivel": "Centro", "Nombre": "EDUARDO AGUILA SANCHEZ", "Centro": "CC2", "Meta": 600},
    {"IDCenter": "JV1", "Nivel": "Centro", "Nombre": "MARIA LUISA MEZA GOEL",  "Centro": "JV",  "Meta": 300},

    # --- Metas Supervisor (JV) ---
    {"IDCenter": "JV2", "Nivel": "Supervisor", "Nombre": "JORGE MIGUEL UREÑA ZARATE",        "Centro": "JV",  "Meta": 157},
    {"IDCenter": "JV3", "Nivel": "Supervisor", "Nombre": "MARIA FERNANDA MARTINEZ BISTRAIN", "Centro": "JV",  "Meta": 143},

    # --- Metas Supervisor (CC2) ---
    {"IDCenter": "CC2", "Nivel": "Supervisor", "Nombre": "ALFREDO CABRERA PADRON",          "Centro": "CC2", "Meta": 131},
    {"IDCenter": "CC4", "Nivel": "Supervisor", "Nombre": "REYNA LIZZETTE MARTINEZ GARCIA",  "Centro": "CC2", "Meta": 161},
    {"IDCenter": "CC3", "Nivel": "Supervisor", "Nombre": "CARLOS ALBERTO AGUILAR CANO",  "Centro": "CC2", "Meta": 149},
    {"IDCenter": "CC5", "Nivel": "Supervisor", "Nombre": "ALAN UZIEL SALAZAR AGUILAR",     "Centro": "CC2", "Meta": 159},

    # ❌ JULIO is intentionally NOT included
]


def load_metas_df(ym_int: int) -> pd.DataFrame:
    # ✅ If monthly Excel exists, use it automatically for Centro/Supervisor metas (Tab 8)
    try:
        df_excel = _load_metas_excel_df(ym_int)
        if df_excel is None or df_excel.empty:
            raise ValueError("Excel de metas vacío.")

        col_ej = _metas_pick_col(df_excel, ["EJECUTIVO", "Ejecutivo"])
        col_sup = _metas_pick_col(df_excel, ["Supervisor", "SUPERVISOR"])
        col_cent = _metas_pick_col(df_excel, ["Centro", "CENTRO"])
        col_meta = _metas_pick_meta_month_col(df_excel, ym_int)

        if not col_ej or not col_sup or not col_cent or not col_meta:
            raise KeyError(
                f"No pude detectar columnas requeridas. Columnas: {list(df_excel.columns)}"
            )

        tmp = df_excel.copy()
        tmp[col_ej] = tmp[col_ej].astype(str).str.strip()
        tmp[col_sup] = tmp[col_sup].astype(str).str.strip()
        tmp[col_cent] = tmp[col_cent].astype(str).str.strip().str.upper()
        tmp[col_meta] = pd.to_numeric(tmp[col_meta], errors="coerce").fillna(0)

        # Coordinators (keep names from manual rows, but meta from Excel sum)
        coord_map = {
            str(r.get("Centro", "")).strip().upper(): (str(r.get("IDCenter", "")), str(r.get("Nombre", "")).strip())
            for r in METAS_MANUAL_ROWS
            if str(r.get("Nivel", "")).strip().lower() == "centro"
        }

        center_tot = tmp.groupby(col_cent, as_index=False)[col_meta].sum()

        rows: list[dict] = []
        for _, rr in center_tot.iterrows():
            cent = str(rr[col_cent]).strip().upper()
            meta_val = float(rr[col_meta] or 0)
            idc, name = coord_map.get(cent, (f"{cent}_COORD", "COORDINADOR"))
            rows.append(
                {
                    "IDCenter": idc,
                    "Nivel": "Centro",
                    "Nombre": name,
                    "Centro": cent,
                    "Meta": meta_val,
                }
            )

        # Supervisor metas from Excel (sum of exec metas)
        sup_tot = tmp.groupby([col_cent, col_sup], as_index=False)[col_meta].sum()
        for i, rr in sup_tot.iterrows():
            cent = str(rr[col_cent]).strip().upper()
            sup = str(rr[col_sup]).strip()
            rows.append(
                {
                    "IDCenter": f"{cent}_SUP{i+1}",
                    "Nivel": "Supervisor",
                    "Nombre": sup,
                    "Centro": cent,
                    "Meta": float(rr[col_meta] or 0),
                }
            )

        df = pd.DataFrame(rows, columns=["IDCenter", "Nivel", "Nombre", "Centro", "Meta"])
        df["Nombre"] = df["Nombre"].astype(str).str.strip()
        df["Centro"] = df["Centro"].astype(str).str.strip().str.upper()
        df["Nivel"] = df["Nivel"].astype(str).str.strip()
        df["Meta"] = pd.to_numeric(df["Meta"], errors="coerce")
        df["Nombre_norm"] = df["Nombre"].apply(normalize_name)
        return df

    except Exception:
        # Fallback to your manual table
        df = pd.DataFrame(METAS_MANUAL_ROWS, columns=["IDCenter", "Nivel", "Nombre", "Centro", "Meta"])
        if df.empty:
            return pd.DataFrame(columns=["IDCenter", "Nivel", "Nombre", "Centro", "Meta", "Nombre_norm"])

        df["Nombre"] = df["Nombre"].astype(str).str.strip()
        df["Centro"] = df["Centro"].astype(str).str.strip().str.upper()
        df["Nivel"] = df["Nivel"].astype(str).str.strip()
        df["Meta"] = pd.to_numeric(df["Meta"], errors="coerce")
        df["Nombre_norm"] = df["Nombre"].apply(normalize_name)
        return df



# -------------------------------
# MEASURES
# -------------------------------
def total_folios(df: pd.DataFrame) -> int:
    return int(len(df))


def total_precio(df: pd.DataFrame) -> float:
    return float(df["PRECIO"].sum(skipna=True))


def total_renta(df: pd.DataFrame) -> float:
    return float(df["RENTA SIN IMPUESTOS"].sum(skipna=True))


def arpu(df: pd.DataFrame) -> float:
    fol = total_folios(df)
    return (total_precio(df) / fol) if fol else np.nan


def arpu_siva(df: pd.DataFrame) -> float:
    fol = total_folios(df)
    return (total_renta(df) / fol) if fol else np.nan


def distinct_days_with_sales(df: pd.DataFrame, exclude_sunday: bool = True) -> int:
    if df.empty:
        return 0
    tmp = df.copy()
    tmp["Fecha_dt"] = pd.to_datetime(tmp["Fecha"])
    if exclude_sunday:
        tmp = tmp[tmp["Fecha_dt"].dt.weekday <= 5]
    return int(tmp["Fecha"].nunique())


def _ventas_sin_domingo(df: pd.DataFrame) -> pd.DataFrame:
    """Rows excluding Sundays (used ONLY for average calculations)."""
    if df.empty:
        return df
    fecha_dt = pd.to_datetime(df["Fecha"], errors="coerce")
    wd = fecha_dt.dt.weekday  # Mon=0 ... Sun=6
    mask = fecha_dt.notna() & (wd != 6)  # exclude Sunday
    return df.loc[mask].copy()


def _dias_equivalentes_lun_sab(df: pd.DataFrame) -> float:
    """
    Working-day equivalent count:
      - Mon-Fri = 1 day
      - Saturday = 0.5 day
      - Sunday excluded (should not be present if you use _ventas_sin_domingo)
    """
    if df.empty:
        return 0.0

    fecha_dt = pd.to_datetime(df["Fecha"], errors="coerce")
    fecha_dt = fecha_dt[fecha_dt.notna()]
    if fecha_dt.empty:
        return 0.0

    # unique calendar days present in data (avoid counting multiple sales same day)
    unique_days = pd.to_datetime(pd.Series(fecha_dt.dt.normalize().unique()))
    if unique_days.empty:
        return 0.0

    wd_u = pd.DatetimeIndex(unique_days).weekday  # array of weekdays
    mon_fri = int(np.sum(wd_u <= 4))
    sat = int(np.sum(wd_u == 5))
    return float(mon_fri + 0.5 * sat)


def promedio_diario(df: pd.DataFrame) -> float:
    """
    ✅ Average rules:
      - Sunday sales DO NOT count in average
      - Saturday counts as 0.5 working day (2 Saturdays = 1 full day)
      - Totals (Total de Ventas) are NOT affected elsewhere
    """
    if df.empty:
        return np.nan

    df_avg = _ventas_sin_domingo(df)  # exclude Sunday sales from numerator
    ventas_avg = int(df_avg.shape[0])

    dias_eq = _dias_equivalentes_lun_sab(df_avg)  # Mon-Fri=1, Sat=0.5
    return (ventas_avg / dias_eq) if dias_eq else np.nan


def weekly_series(df: pd.DataFrame) -> pd.DataFrame:
    """
    Weekly series where:
      - Bar (Ventas) = TOTAL sales including Sundays
      - Line (PromDiarioSemana) = average excluding Sundays + Saturday=0.5 day
    """
    if df.empty:
        return pd.DataFrame(columns=["SemanaISO", "WeekKey", "Ventas", "VentasNoDomingo", "PromDiarioSemana", "DiasSemanaEq"])

    # TOTAL weekly sales (includes Sundays) -> for the BAR
    ventas_total = (
        df.groupby(["WeekKey", "SemanaISO"], as_index=False)
        .size()
        .rename(columns={"size": "Ventas"})
    )

    # Data for AVERAGE (exclude Sundays)
    tmp = df.copy()
    tmp["_Fecha_dt"] = pd.to_datetime(tmp["Fecha"], errors="coerce")
    tmp = tmp[tmp["_Fecha_dt"].notna()].copy()
    tmp["_wd"] = tmp["_Fecha_dt"].dt.weekday

    tmp_avg = tmp[tmp["_wd"] != 6].copy()  # exclude Sunday rows for average numerator/denominator

    # Weekly sales excluding Sunday -> numerator for average
    ventas_no_dom = (
        tmp_avg.groupby(["WeekKey", "SemanaISO"], as_index=False)
        .size()
        .rename(columns={"size": "VentasNoDomingo"})
    )

    # Weighted equivalent days per week: Mon-Fri=1, Sat=0.5
    days_unique = tmp_avg.drop_duplicates(["WeekKey", "_Fecha_dt"])[["WeekKey", "_Fecha_dt", "_wd"]].copy()
    days_unique["w"] = np.where(days_unique["_wd"] == 5, 0.5, 1.0)
    dias_eq = (
        days_unique.groupby("WeekKey", as_index=False)["w"]
        .sum()
        .rename(columns={"w": "DiasSemanaEq"})
    )

    out = ventas_total.merge(ventas_no_dom, on=["WeekKey", "SemanaISO"], how="left")
    out = out.merge(dias_eq, on="WeekKey", how="left")

    out["VentasNoDomingo"] = out["VentasNoDomingo"].fillna(0).astype(int)
    out["PromDiarioSemana"] = out["VentasNoDomingo"] / out["DiasSemanaEq"].replace({0: np.nan})

    return out.sort_values("WeekKey")






def max_folios_dia(df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    g = df.groupby("Fecha", as_index=False).size()
    return int(g["size"].max())


def daily_series(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            {
                "Fecha": pd.to_datetime([]),
                "Ventas": pd.Series([], dtype="int64"),
            }
        )
    out = df.groupby("Fecha", as_index=False).size().rename(columns={"size": "Ventas"})
    out["Fecha"] = pd.to_datetime(out["Fecha"], errors="coerce")
    return out.sort_values("Fecha")



    ventas_sem = (
        df.groupby(["WeekKey", "SemanaISO"], as_index=False)
        .size()
        .rename(columns={"size": "Ventas"})
    )

    tmp = df.copy()
    tmp["Fecha_dt"] = pd.to_datetime(tmp["Fecha"])
    tmp = tmp[tmp["Fecha_dt"].dt.weekday <= 5]  # lunes-sábado

    dias_sem = (
        tmp.groupby("WeekKey")["Fecha"]
        .nunique()
        .reset_index()
        .rename(columns={"Fecha": "DiasSemana"})
    )

    out = ventas_sem.merge(dias_sem, on="WeekKey", how="left")
    out["PromDiarioSemana"] = out["Ventas"] / out["DiasSemana"].replace({0: np.nan})
    return out.sort_values("WeekKey")


def filter_month(df: pd.DataFrame, ym: int) -> pd.DataFrame:
    return df[df["AñoMes"] == ym].copy()


def cut_month_mode(df: pd.DataFrame, mode: int, day_cut: int) -> pd.DataFrame:
    if mode == 1:
        return df[df["DiaNum"] <= day_cut].copy()
    return df.copy()


# -------------------------------
# GAUGE
# -------------------------------
def gauge_fig(value: float, meta: float, title: str):
    value = 0 if value is None or np.isnan(value) else float(value)
    meta = 0 if meta is None or np.isnan(meta) else float(meta)
    axis_max = max(meta, value, 1)

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=value,
            number={"font": {"size": 38}},
            title={"text": title, "font": {"size": 12}},
            gauge={
                "axis": {"range": [0, axis_max], "tickwidth": 0},
                "bar": {"color": "rgba(127,127,127,0.65)", "thickness": 0.35},
                "bgcolor": "rgba(0,0,0,0)",
                "borderwidth": 0,
                "steps": [{"range": [0, axis_max], "color": "rgba(127,127,127,0.18)"}],
            },
        )
    )
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        margin=dict(l=8, r=8, t=40, b=10),
        height=250,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


# -------------------------------
# DETALLE HELPERS (CC2 / JV + pies)
# -------------------------------
def build_detalle_matrix(df_: pd.DataFrame) -> pd.DataFrame:
    if df_.empty:
        return pd.DataFrame(columns=["Supervisor", "Ejecutivo", "TotalVentas", "MontoVendido", "ARPU"])

    by_sup = (
        df_.groupby("Supervisor", as_index=False)
        .agg(
            TotalVentas=("FOLIO", "count"),
            MontoVendido=("PRECIO", "sum"),
            ARPU=("PRECIO", lambda s: s.sum() / len(s) if len(s) else np.nan),
        )
        .sort_values("TotalVentas", ascending=False)
    )

    by_ej = (
        df_.groupby(["Supervisor", "EJECUTIVO"], as_index=False)
        .agg(
            TotalVentas=("FOLIO", "count"),
            MontoVendido=("PRECIO", "sum"),
            ARPU=("PRECIO", lambda s: s.sum() / len(s) if len(s) else np.nan),
        )
        .sort_values(["Supervisor", "TotalVentas"], ascending=[True, False])
    )

    rows = []
    for _, sr in by_sup.iterrows():
        sup = sr["Supervisor"]
        rows.append(
            {
                "Supervisor": str(sup),
                "Ejecutivo": "",
                "TotalVentas": int(sr["TotalVentas"]),
                "MontoVendido": float(sr["MontoVendido"]),
                "ARPU": float(sr["ARPU"]) if pd.notna(sr["ARPU"]) else np.nan,
            }
        )
        sub = by_ej[by_ej["Supervisor"] == sup].sort_values("TotalVentas", ascending=False)
        for _, er in sub.iterrows():
            rows.append(
                {
                    "Supervisor": "",
                    "Ejecutivo": "   " + str(er["EJECUTIVO"]),
                    "TotalVentas": int(er["TotalVentas"]),
                    "MontoVendido": float(er["MontoVendido"]),
                    "ARPU": float(er["ARPU"]) if pd.notna(er["ARPU"]) else np.nan,
                }
            )

    out = pd.DataFrame(rows, columns=["Supervisor", "Ejecutivo", "TotalVentas", "MontoVendido", "ARPU"])
    return out


def donut_compare_fig(labels: list[str], values: list[float], title: str, value_formatter):
    vals = [0.0 if (v is None or (isinstance(v, float) and np.isnan(v))) else float(v) for v in values]
    total = float(np.sum(vals))
    if total <= 0:
        return None

    texts = [value_formatter(v) for v in vals]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=vals,
                hole=0.55,
                text=texts,
                textinfo="label+text+percent",
                insidetextorientation="radial",
                sort=False,
            )
        ]
    )
    fig.update_layout(title=title, height=320)
    apply_plotly_theme(fig)
    return fig


# -------------------------------
# LOAD DATA
# -------------------------------
st.sidebar.header("⚙️ Parámetros")

# Refresh button
if "last_refresh" not in st.session_state:
    st.session_state["last_refresh"] = None

btn_cols = st.sidebar.columns([1, 1])
with btn_cols[0]:
    if st.button("🔄 Actualizar datos", use_container_width=True):
        st.cache_data.clear()
        st.cache_resource.clear()  # ✅ IMPORTANT: clears the cached pyodbc connection
        st.session_state["last_refresh"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        st.rerun()
with btn_cols[1]:
    st.caption(f"🕒 {st.session_state['last_refresh']}" if st.session_state["last_refresh"] else "")

# ✅ Default date range: Inicio fijo 2025-12-29, Fin = día actual
default_start_dt = date(2025, 12, 29)
default_end_dt = date.today()

d1, d2 = st.sidebar.columns(2)
start_dt = d1.date_input("Inicio", value=default_start_dt, format="YYYY-MM-DD")
end_dt = d2.date_input("Fin", value=default_end_dt, format="YYYY-MM-DD")

if start_dt > end_dt:
    st.sidebar.error("⚠️ Inicio no puede ser mayor que Fin.")
    st.stop()

start_yyyymmdd = start_dt.strftime("%Y%m%d")
end_yyyymmdd = end_dt.strftime("%Y%m%d")

with st.spinner("Cargando datos desde SQL Server…"):
    empleados = load_empleados()
    ventas_raw = load_ventas(start_yyyymmdd, end_yyyymmdd)
    ventas = add_empleado_join(ventas_raw, empleados)


# ==============================================================

# ✅ ONLY FOR FILTER OPTIONS (do NOT change data logic):
#    - remove BAJA supervisors (people no longer with you)
#    - remove EDUARDO AGUILA SANCHEZ from Supervisor filters (he is coordinator of supervisors)
EXCLUDED_SUP_NORMS = {normalize_name("BAJA"), normalize_name("EDUARDO AGUILA SANCHEZ")}
ventas_filtros = ventas[~ventas["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS)].copy()

if ventas.empty:
    st.error("No hay datos en el rango seleccionado.")
    st.stop()

# ✅ Months should come from the selected date range (even if 0 sales)
months_range = pd.period_range(start=start_dt.replace(day=1), end=end_dt.replace(day=1), freq="M")
meses_disponibles = sorted([p.year * 100 + p.month for p in months_range])

# Labels for all months in range
mes_labels = {ym: month_key_to_name_es(int(ym)) for ym in meses_disponibles}

# (Optional) still keep labels for any unexpected months from data
for ym in ventas["AñoMes"].dropna().unique().tolist():
    ym = int(ym)
    if ym not in mes_labels:
        mes_labels[ym] = month_key_to_name_es(ym)


mes_opciones = ["Todos"] + meses_disponibles

mes_choice = st.sidebar.selectbox(
    "Mes",
    options=mes_opciones,
    format_func=lambda x: "Todos" if x == "Todos" else mes_labels.get(x, str(x)),
    index=0,
)

mes_sel_all = mes_choice == "Todos"
meses_sel = meses_disponibles if mes_sel_all else [int(mes_choice)]

# Dejamos mes_sel como un solo mes para no romper el resto del código
mes_sel = int(meses_sel[-1])

metas = load_metas_df(int(mes_sel))

# Texto opcional por si luego quieres mostrar el título correcto
mes_title = "Todos los meses del intervalo" if mes_sel_all else mes_labels.get(mes_sel, str(mes_sel))

# ==========================================================
# ✅ Mantener a Maria Luisa visible en el filtro de supervisores
# ==========================================================
ml_norm = normalize_name("MARIA LUISA MEZA GOEL")

# Si es marzo 2026, conservar este ajuste por si la BD la trae como BAJA
if int(mes_sel) == 202603:
    empleados.loc[empleados["Nombre"].apply(normalize_name) == ml_norm, "Estatus"] = "ACTIVO"
# ==========================================================

center_keys = ["CC2", "JV"]
center_sel = st.sidebar.multiselect("Centro (CC2 / JV)", options=center_keys, default=center_keys)

# ✅ Prevent Streamlit from keeping excluded supervisors in session state selections
if "Supervisor" in st.session_state:
    st.session_state["Supervisor"] = [
        s for s in st.session_state["Supervisor"]
        if normalize_name(s) not in EXCLUDED_SUP_NORMS
    ]

# ✅ Supervisor options (exclude BAJA + Eduardo only in the filter list)
supervisores = sorted([s for s in ventas_filtros["Supervisor"].dropna().unique().tolist()])
sup_sel = st.sidebar.multiselect("Supervisor", options=supervisores, default=[])

# ✅ Ejecutivo options (DO NOT exclude BAJA here — exclusion is ONLY in Tendencia Ejecutivo tab)
ejecutivos = sorted([e for e in ventas["EJECUTIVO"].dropna().unique().tolist()])

if "Ejecutivo" in st.session_state:
    st.session_state["Ejecutivo"] = [e for e in st.session_state["Ejecutivo"] if e in ejecutivos]
ej_sel = st.sidebar.multiselect("Ejecutivo", options=ejecutivos, default=[])

subregs = sorted([s for s in ventas["SUBREGION"].dropna().unique().tolist()])
sub_sel = st.sidebar.multiselect("Subregión", options=subregs, default=[])

# Filters (month + sidebar filters)
df_base = ventas.copy()
df_base = df_base[df_base["AñoMes"].isin(meses_sel)]
if center_sel:
    df_base = df_base[df_base["CentroKey"].isin(center_sel)]
if sup_sel:
    df_base = df_base[df_base["Supervisor"].isin(sup_sel)]
if ej_sel:
    df_base = df_base[df_base["EJECUTIVO"].isin(ej_sel)]
if sub_sel:
    df_base = df_base[df_base["SUBREGION"].isin(sub_sel)]

# -------------------------------
# ✅ DATA LOADERS (Add these functions)
# -------------------------------


from pathlib import Path

@st.cache_data(ttl=600, show_spinner=False)
def load_metas_from_csv(ym_int: int) -> dict:
    """Loads executive metas from the Excel of the selected month. Returns {EJ_NORM: meta}."""
    try:
        df = _load_metas_excel_df(ym_int)
        if df is None or df.empty:
            return {}

        col_ej = _metas_pick_col(df, ["EJECUTIVO", "Ejecutivo"])
        col_meta = _metas_pick_meta_month_col(df, ym_int)

        if not col_ej:
            raise KeyError(f"Falta columna 'EJECUTIVO'. Columnas: {list(df.columns)}")
        if not col_meta:
            raise KeyError(f"No encontré la columna de meta del mes seleccionado. Columnas: {list(df.columns)}")

        tmp = df.copy()
        tmp[col_ej] = tmp[col_ej].astype(str).str.strip()
        tmp["EJ_NORM"] = tmp[col_ej].apply(normalize_name)
        tmp[col_meta] = pd.to_numeric(tmp[col_meta], errors="coerce").fillna(0)

        return tmp.groupby("EJ_NORM")[col_meta].sum().to_dict()

    except Exception as e:
        st.warning(f"No se pudo cargar el archivo de metas del mes seleccionado: {e}")
        return {}


@st.cache_data(ttl=600, show_spinner=False)
def load_metas_supervisor_from_excel(ym_int: int) -> dict:
    """Loads supervisor metas from the Excel of the selected month. Returns {SUP_NORM: meta_sum}."""
    try:
        df = _load_metas_excel_df(ym_int)
        if df is None or df.empty:
            return {}

        col_sup = _metas_pick_col(df, ["Supervisor", "SUPERVISOR"])
        col_meta = _metas_pick_meta_month_col(df, ym_int)

        if not col_sup or not col_meta:
            return {}

        tmp = df.copy()
        tmp[col_sup] = tmp[col_sup].astype(str).str.strip()
        tmp[col_meta] = pd.to_numeric(tmp[col_meta], errors="coerce").fillna(0)

        tmp["SUP_NORM"] = tmp[col_sup].apply(normalize_name)
        return tmp.groupby("SUP_NORM")[col_meta].sum().to_dict()

    except Exception:
        return {}
    
# =========================================================
# ✅ TRANSITO LOGIC FOR HTML DASHBOARD (BACK OFFICE & PIPELINE)
# =========================================================
def parse_backoffice_datetime(series: pd.Series, window_start: date | None = None, window_end: date | None = None) -> pd.Series:
    """Exact date parser from transito.py to guarantee 1-to-1 match."""
    s = series.astype(str).str.strip()
    s = s.replace({"nan": "", "None": "", "NaT": ""})
    s = s.where(s != "", np.nan)

    if s.notna().any():
        pat = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?)|(\d{4}[/-]\d{1,2}[/-]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)"
        ext = s.astype(str).str.extract(pat)
        ext = ext[0].fillna(ext[1])
        s2 = ext.where(ext.notna(), s)
    else:
        s2 = s

    dt_dayfirst = pd.to_datetime(s2, errors="coerce", dayfirst=True)
    dt_monthfirst = pd.to_datetime(s2, errors="coerce", dayfirst=False)

    if window_start is None or window_end is None:
        return dt_dayfirst

    w0 = pd.Timestamp(window_start)
    w1 = pd.Timestamp(window_end) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

    in1 = dt_dayfirst.between(w0, w1)
    in2 = dt_monthfirst.between(w0, w1)

    out = dt_dayfirst.copy()
    out = out.where(~(in2 & ~in1), dt_monthfirst)
    out = out.where(~(dt_dayfirst.isna() & dt_monthfirst.notna()), dt_monthfirst)
    return out

def choose_backoffice_dt_html(df: pd.DataFrame, window_start: date, window_end: date) -> pd.Series:
    # Same logic as Transito Global 2.0
    if "BO_DT_DF" in df.columns and "BO_DT_MF" in df.columns:
        dt_dayfirst = df["BO_DT_DF"]
        dt_monthfirst = df["BO_DT_MF"]

        w0 = pd.Timestamp(window_start)
        w1 = pd.Timestamp(window_end) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

        in1 = dt_dayfirst.between(w0, w1)
        in2 = dt_monthfirst.between(w0, w1)

        out = dt_dayfirst.copy()
        out = out.where(~(in2 & ~in1), dt_monthfirst)
        out = out.where(~(dt_dayfirst.isna() & dt_monthfirst.notna()), dt_monthfirst)
        return out

    # fallback
    return parse_backoffice_datetime(df["Back Office"], window_start=window_start, window_end=window_end)

# =========================================================
# ✅ SINGLE CACHED FETCH FOR ALL TRANSITO & HTML DATA
# =========================================================
@st.cache_data(ttl=600, show_spinner=False)
def fetch_programacion_history(end_date: date) -> pd.DataFrame:
    """
    Trae la misma base necesaria para Interfaz Custom,
    pero con la misma limpieza y parsing usada en Transito Global 2.0
    para que Back Office cuadre 1:1.
    """
    fi_str = "20251101"
    ff_str = end_date.strftime("%Y%m%d")

    q = f"""
    SELECT
        LTRIM(RTRIM([Vendedor]))        AS EJECUTIVO,
        LTRIM(RTRIM([Estatus]))         AS Estatus,
        [Venta],
        [Back Office],
        [Fecha creacion],
        [Tienda solicita]               AS Centro
    FROM reporte_programacion_entrega('empresa_maestra', 4, '{fi_str}', '{ff_str}')
    WHERE
        [Tienda solicita] LIKE 'EXP ATT C CENTER%'
        AND [Estatus] IN ('En entrega','Canc Error','Entregado','En preparacion','Back Office','Solicitado')
    """

    try:
        df = read_sql(q)

        if df.empty:
            return df

        # -------------------------
        # Limpieza exacta estilo Tránsito
        # -------------------------
        for col in ["EJECUTIVO", "Estatus", "Back Office", "Centro"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({"nan": np.nan, "None": np.nan})

        if "Venta" in df.columns:
            df["Venta"] = df["Venta"].replace({"nan": np.nan, "None": np.nan})

        # ✅ misma exclusión que Transito
        if "EJECUTIVO" in df.columns:
            df = df[df["EJECUTIVO"].str.upper() != EXCLUDED_VENDOR].copy()

        # ✅ mismas correcciones de nombre que ventas
        df["EJECUTIVO"] = df["EJECUTIVO"].replace(
            {
                "CESAR JAHACIEL ALONSO GARCIAA": "CESAR JAHACIEL ALONSO GARCIA",
                "VICTOR BETANZO FUENTES": "VICTOR BETANZOS FUENTES",
            }
        )

        df["EJ_NORM"] = df["EJECUTIVO"].astype(str).apply(normalize_name)
        df["Estatus_upper"] = df["Estatus"].astype(str).str.strip().str.upper()
        df["Venta_Vacia"] = df["Venta"].isna() | (df["Venta"].astype(str).str.strip() == "")
        df["Fecha_Creacion_DT"] = pd.to_datetime(df["Fecha creacion"], errors="coerce", dayfirst=True).dt.date

        # CentroKey igual que en ventas
        df["CentroKey"] = np.where(
            df["Centro"].astype(str).str.upper().str.contains("JUAREZ", na=False),
            "JV",
            "CC2",
        )

        # -------------------------
        # Pre-parse exacto Back Office
        # (igual filosofía que Transito Global 2.0)
        # -------------------------
        s = df["Back Office"].astype(str).str.strip()
        s = s.replace({"nan": "", "None": "", "NaT": ""})
        s = s.where(s != "", np.nan)

        if s.notna().any():
            pat = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?)|(\d{4}[/-]\d{1,2}[/-]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)"
            ext = s.astype(str).str.extract(pat)
            ext = ext[0].fillna(ext[1])
            s2 = ext.where(ext.notna(), s)
        else:
            s2 = s

        df["BO_DT_DF"] = pd.to_datetime(s2, errors="coerce", dayfirst=True)
        df["BO_DT_MF"] = pd.to_datetime(s2, errors="coerce", dayfirst=False)

        return df

    except Exception as e:
        print(f"Error loading history: {e}")
        return pd.DataFrame()


# -------------------------------
# TABS
# -------------------------------
tabs = st.tabs(
    [
        "🌐 Global Mes",
        "📅 Semanas",
        "🆚 JV vs CC2",
        "📉 Mes vs Mes",
        "📋 Detalle",
        "🗺️ Región",
        "🏆 Tops",
        "🎯 Metas",
        "📈 Tendencia Ejecutivo",
        "✨ Interfaz Custom", # <-- NEW TAB
    ]
)

# ======================================================
# TAB 1: Global del Mes
# ======================================================
with tabs[0]:
    st.markdown(f"## Ventas Globales del Mes — **{mes_labels[mes_sel]}**")

    # --- Selection ARPU (day selected in the bar chart) ---
    sel_arpu_label = None
    sel_arpu_val = None
    sel_arpu_siva_val = None

    colA, colB = st.columns([0.78, 0.22], gap="large")

    with colA:
        k1, k2, k3, k4 = st.columns(4, gap="medium")
        with k1:
            metric_card("Promedio de Ventas", fmt_int(promedio_diario(df_base)))
        with k2:
            metric_card("Total de Ventas", fmt_int(total_folios(df_base)))
        with k3:
            metric_card("Max Ventas en un Día", fmt_int(max_folios_dia(df_base)))
        with k4:
            metric_card("Monto Vendido", fmt_money_short(total_precio(df_base)))

        s = daily_series(df_base)
        avg_line = promedio_diario(df_base)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=s["Fecha"], y=s["Ventas"], name="Total de Ventas"))
        fig.add_trace(go.Scatter(x=s["Fecha"], y=[avg_line] * len(s), mode="lines", name="Promedio Diario de Ventas"))
        fig.update_layout(
            title="Vista General de Ventas",
            xaxis_title="Fecha",
            yaxis_title="Total de Ventas",
            height=380,
        )
        apply_plotly_theme(fig)
        # ✅ Enable click/selection on bars to retrieve selected day(s)
        event = st.plotly_chart(
            fig,
            width="stretch",
            key=f"t0_global_mes_{mes_sel}",
            on_select="rerun",
            selection_mode="points",
        )

        # --- Extract selected day(s) robustly (prefer point_indices to avoid date parsing quirks) ---
        sel_dates = []

        try:
            idxs = list(getattr(event.selection, "point_indices", []))
        except Exception:
            idxs = []

        # Fallback if the object behaves more like a dict
        if not idxs:
            try:
                idxs = list(event["selection"]["point_indices"])
            except Exception:
                idxs = []

        # Map indices back to the daily series (s) to get the exact date(s)
        for ix in idxs:
            try:
                ix = int(ix)
                if 0 <= ix < len(s):
                    d = pd.to_datetime(s.iloc[ix]["Fecha"], errors="coerce")
                    if pd.notna(d):
                        sel_dates.append(d.date())
            except Exception:
                pass

        # Fallback: if no indices, try reading x directly from points
        if not sel_dates:
            try:
                pts = getattr(event.selection, "points", [])
            except Exception:
                pts = []
            if not pts:
                try:
                    pts = event.get("selection", {}).get("points", [])
                except Exception:
                    pts = []

            for p in pts:
                try:
                    xval = p.get("x", None)
                    d = pd.to_datetime(xval, errors="coerce")
                    if pd.notna(d):
                        sel_dates.append(d.date())
                except Exception:
                    pass

        if sel_dates:
            sel_dates = sorted(set(sel_dates))
            df_sel_day = df_base[df_base["Fecha"].isin(sel_dates)].copy()

            if len(sel_dates) == 1:
                sel_arpu_label = sel_dates[0].strftime("%d/%m/%Y")
            else:
                sel_arpu_label = f"{sel_dates[0].strftime('%d/%m/%Y')} → {sel_dates[-1].strftime('%d/%m/%Y')}"

            sel_arpu_val = arpu(df_sel_day)
            sel_arpu_siva_val = arpu_siva(df_sel_day)

        top_days = s.sort_values("Ventas", ascending=False).head(8).copy()
        top_days["Fecha"] = top_days["Fecha"].dt.strftime("%A, %d %B %Y")

        # ✅ ADD TOTAL ROW (month total) + ✅ BOLD TOTAL
        top_days_show = top_days.rename(columns={"Ventas": "Total de Ventas"})
        top_days_show = add_totals_row(
            top_days_show,
            label_col="Fecha",
            totals={"Total de Ventas": total_folios(df_base)},
            label="TOTAL (Mes)",
        )

        st.dataframe(
            style_totals_bold(top_days_show, label_col="Fecha").format({"Total de Ventas": "{:,.0f}"}),
            hide_index=True,
            width="stretch",
        )

        # ---------------------------------------------------------
        # ✅ NEW: Ventas por Equipo (Supervisor) — intervalo seleccionado (TAB 1)
        #     - If you selected days in the bar chart => uses ONLY those days
        #     - Else => uses full month (df_base)
        # ---------------------------------------------------------
        st.markdown("---")
        st.markdown("### 👥 Ventas por Equipo (Supervisor) — intervalo seleccionado")

        df_interval = df_base.copy()

        # If user selected points (days) in the chart, use those dates only
        if sel_dates:
            df_interval = df_base[df_base["Fecha"].isin(sel_dates)].copy()
            st.caption(f"Intervalo: {sel_arpu_label}")
        else:
            st.caption("Intervalo: Mes completo (según filtros)")

        # Optional: exclude BAJA + Eduardo from this visualization (consistent with filter list)
        if "Supervisor_norm" in df_interval.columns:
            df_interval = df_interval[~df_interval["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS)].copy()

        if df_interval.empty:
            st.info("No hay datos para graficar por equipo con el intervalo/filtros actuales.")
        else:
            team_kpi = (
                df_interval.groupby("Supervisor", as_index=False)
                .agg(
                    Ventas=("FOLIO", "count"),
                    Ejecutivos=("EJECUTIVO", "nunique"),
                    MontoVendido=("PRECIO", "sum"),
                )
                .copy()
            )

            team_kpi["ARPU"] = np.where(team_kpi["Ventas"] > 0, team_kpi["MontoVendido"] / team_kpi["Ventas"], 0.0)

            # Nice label
            team_kpi["Etiqueta"] = (
                team_kpi["Ventas"].astype(int).map(lambda x: f"{x:,}")
                + " ventas | "
                + team_kpi["Ejecutivos"].astype(int).map(lambda x: f"{x:,}")
                + " ejecutivos"
            )

            team_kpi = team_kpi.sort_values("Ventas", ascending=False).reset_index(drop=True)

            dyn_h = max(360, 140 + 32 * len(team_kpi))

            fig_team = px.bar(
                team_kpi.sort_values("Ventas", ascending=True),
                x="Ventas",
                y="Supervisor",
                orientation="h",
                text="Etiqueta",
                title="Ventas por Supervisor (Equipo)",
                hover_data={"MontoVendido": ":,.2f", "ARPU": ":,.2f", "Ventas": True, "Ejecutivos": True, "Etiqueta": False},
                template=PLOTLY_TEMPLATE,
            )
            fig_team.update_traces(textposition="outside", cliponaxis=False)
            fig_team.update_layout(
                height=dyn_h,
                showlegend=False,
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                margin=dict(l=320, r=30, t=70, b=20),
                xaxis_title="Ventas",
                yaxis_title="Supervisor",
            )
            apply_plotly_theme(fig_team)

            st.plotly_chart(fig_team, width="stretch", key=f"t0_team_interval_{mes_sel}_{len(sel_dates) if sel_dates else 0}")

            # Table summary (with TOTAL row)
            team_show = team_kpi[["Supervisor", "Ventas", "Ejecutivos", "MontoVendido", "ARPU"]].copy()
            team_show = add_totals_row(
                team_show,
                label_col="Supervisor",
                totals={
                    "Ventas": int(team_show["Ventas"].sum()),
                    "Ejecutivos": int(team_show["Ejecutivos"].sum()),
                    "MontoVendido": float(team_show["MontoVendido"].sum()),
                    "ARPU": (float(team_show["MontoVendido"].sum()) / int(team_show["Ventas"].sum())) if int(team_show["Ventas"].sum()) > 0 else 0.0,
                },
                label="TOTAL",
            )

            st.dataframe(
                style_totals_bold(team_show, label_col="Supervisor").format(
                    {"Ventas": "{:,.0f}", "Ejecutivos": "{:,.0f}", "MontoVendido": "${:,.2f}", "ARPU": "${:,.2f}"}
                ),
                hide_index=True,
                width="stretch",
            )

        # ✅ DOWNLOAD EXCEL (Global Mes) — UPDATED (includes ALL ventas)

        # 1️⃣ Full ventas detail (
        ventas_full = df_base.copy()

        # Optional: reorder / clean columns for export
        cols_order = [
            "FOLIO",
            "Telefono cliente",
            "Fecha",
            "Hora",
            "EJECUTIVO",
            "Supervisor",
            "CENTRO",
            "CentroKey",
            "PLAN",
            "PRECIO",
            "RENTA SIN IMPUESTOS",
            "SUBREGION",
            "ESTATUS",
        ]
        cols_order = [c for c in cols_order if c in ventas_full.columns]
        ventas_full = ventas_full[cols_order]

        # 2️⃣ Build sheets
        sheets_gm = {
            "Todas las Ventas (Detalle)": ventas_full,  
            "Top dias (tabla)": top_days_show.copy(),
            "Serie diaria (grafica)": s.rename(columns={"Ventas": "Total de Ventas"}).copy(),
        }

        if sel_dates:
            sheets_gm["Seleccion (dia)"] = df_sel_day.copy()


        st.download_button(
            "⬇️ Descargar Excel (Global Mes)",
            data=build_excel_bytes(sheets_gm),
            file_name=f"GlobalMes_{mes_sel}_{date.today():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_global_mes_{mes_sel}",
            use_container_width=True,
        )

        st.caption(f"🕒 Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    with colB:
        st.markdown("### ARPU")
        kpi_mini("ARPU", fmt_money_short(arpu(df_base)))
        kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_base)))

        # ✅ Show ARPU for the selected day in the bar chart (if any)
        if sel_arpu_label:
            st.markdown("#### ARPU del día seleccionado")
            kpi_mini(f"ARPU ({sel_arpu_label})", fmt_money_short(sel_arpu_val))
            kpi_mini(f"ARPU S/IVA ({sel_arpu_label})", fmt_money_short(sel_arpu_siva_val))
        else:
            st.caption("👆 Selecciona un día en la barra para ver el ARPU de ese día.")


# ======================================================
# TAB 2: Semanas
# ======================================================
with tabs[1]:
    st.markdown("## Vista por Semanas")

    # ✅ Export holders (ADDED)
    by_day_export = None
    cmp_export = None
    interval_summary_export = None
    hour_breakdown_export = None

    df_sem = ventas.copy()
    if center_sel:
        df_sem = df_sem[df_sem["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_sem = df_sem[df_sem["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_sem = df_sem[df_sem["EJECUTIVO"].isin(ej_sel)]
    if sub_sel:
        df_sem = df_sem[df_sem["SUBREGION"].isin(sub_sel)]

    meses_sem = sorted(df_sem["AñoMes"].dropna().unique().tolist())
    sem_labels = ["Todos los meses"] + [mes_labels.get(m, str(m)) for m in meses_sem]
    sem_choice = st.selectbox("Mes (Semanas)", options=sem_labels, index=0, key="sem_mes_choice")

    if sem_choice != "Todos los meses":
        inv = {mes_labels.get(m, str(m)): m for m in meses_sem}
        mes_sem_sel = inv.get(sem_choice)
        if mes_sem_sel is not None:
            month_rows = df_sem[df_sem["AñoMes"] == mes_sem_sel].copy()
            week_keys = month_rows["WeekKey"].dropna().unique().tolist()
            df_sem = df_sem[df_sem["WeekKey"].isin(week_keys)].copy()

    w = weekly_series(df_sem)
    prom_global_sem = np.nanmean(w["PromDiarioSemana"].values) if not w.empty else np.nan

    col1, col2, col3 = st.columns([0.33, 0.34, 0.33])
    with col1:
        metric_card("Promedio de Ventas", fmt_int(prom_global_sem))
    with col2:
        metric_card("Total de Ventas", fmt_int(total_folios(df_sem)))
    with col3:
        if not df_sem.empty:
            last_week_key = int(df_sem["WeekKey"].max())
            tmpw = df_sem[df_sem["WeekKey"] == last_week_key].copy()
            dias = distinct_days_with_sales(tmpw, exclude_sunday=True)
            metric_card("Semana actual", f"{dias} días")

    figw = go.Figure()
    figw.add_trace(go.Bar(x=w["SemanaISO"], y=w["Ventas"], name="Total de Ventas"))
    figw.add_trace(go.Scatter(x=w["SemanaISO"], y=w["PromDiarioSemana"], mode="lines+markers", name="Promedio Diario por Semana"))
    figw.update_layout(
        title="Vista General de Ventas",
        xaxis_title="SEMANA",
        yaxis_title="Total de Ventas",
        height=460,
    )
    apply_plotly_theme(figw)
    st.plotly_chart(figw, width="stretch", key="t1_semanas_general")

    st.markdown("### ARPU")
    kpi_mini("ARPU", fmt_money_short(arpu(df_sem)))
    kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_sem)))

    # ==========================================================
    # ✅ NEW (TAB SEMANAS): Vista por meses y semanas (como en Mes vs Mes)
    # ==========================================================
    st.markdown("---")
    st.markdown("## ✅ Vista por meses y semanas")

    df_sem_mvw_ctx = ventas.copy()
    if center_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["EJECUTIVO"].isin(ej_sel)]
    if sub_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["SUBREGION"].isin(sub_sel)]

    df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["FECHA DE CAPTURA"].notna()].copy()
    df_sem_mvw_ctx["M_DT"] = pd.to_datetime(df_sem_mvw_ctx["FECHA DE CAPTURA"], errors="coerce")
    df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["M_DT"].notna()].copy()

    df_sem_mvw_ctx["M_Fecha"] = df_sem_mvw_ctx["M_DT"].dt.date
    df_sem_mvw_ctx["M_Hora"] = df_sem_mvw_ctx["M_DT"].dt.hour

    if df_sem_mvw_ctx.empty:
        st.info("No hay datos disponibles para la vista por meses/semanas con los filtros actuales.")
    else:
        df_sem_mvw_ctx["M_MonthKey"] = df_sem_mvw_ctx["M_DT"].dt.strftime("%Y-%m")
        df_sem_mvw_ctx["M_MonthName"] = df_sem_mvw_ctx["M_DT"].dt.strftime("%B")
        df_sem_mvw_ctx["M_MonthLabel"] = df_sem_mvw_ctx["M_MonthKey"] + " (" + df_sem_mvw_ctx["M_MonthName"] + ")"

        month_start = df_sem_mvw_ctx["M_DT"].dt.to_period("M").dt.to_timestamp()
        first_wd = month_start.dt.weekday
        df_sem_mvw_ctx["M_WeekOfMonth"] = ((df_sem_mvw_ctx["M_DT"].dt.day + first_wd - 1) // 7) + 1

        st.markdown("### Vista por meses y semanas (Ventas)")

        month_map = (
            df_sem_mvw_ctx[["M_MonthKey", "M_MonthLabel"]]
            .dropna()
            .drop_duplicates()
            .sort_values("M_MonthKey")
        )
        m_options = month_map["M_MonthLabel"].tolist()

        # Default: mes seleccionado en sidebar + (si existe) el mes anterior por orden
        def _ym_to_monthkey_from_ym(ym_int: int) -> str:
            y = ym_int // 100
            m = ym_int % 100
            return f"{y}-{m:02d}"

        cur_key = _ym_to_monthkey_from_ym(int(mes_sel)) if mes_sel else None
        keys_sorted = month_map["M_MonthKey"].tolist()

        defaults = []
        if cur_key and cur_key in keys_sorted:
            idx = keys_sorted.index(cur_key)
            defaults_keys = [keys_sorted[idx]]
            if idx - 1 >= 0:
                defaults_keys.insert(0, keys_sorted[idx - 1])
            defaults = month_map[month_map["M_MonthKey"].isin(defaults_keys)]["M_MonthLabel"].tolist()

        if not defaults:
            defaults = m_options[-2:] if len(m_options) >= 2 else m_options

        m_sel = st.multiselect(
            "Selecciona uno o más meses (Ventas)",
            options=m_options,
            default=defaults,
            key="sem_mvw_months_multi",
        )

        df_mvw = df_sem_mvw_ctx.copy()
        if m_sel:
            df_mvw = df_mvw[df_mvw["M_MonthLabel"].isin(m_sel)].copy()
        else:
            df_mvw = df_mvw.iloc[0:0].copy()

        if df_mvw.empty:
            st.info("No hay datos para los meses seleccionados.")
        else:
            df_mvw["M_WeekLabel"] = df_mvw["M_MonthLabel"] + " - Semana " + df_mvw["M_WeekOfMonth"].astype(int).astype(str)

            w_map = (
                df_mvw[["M_MonthKey", "M_MonthLabel", "M_WeekOfMonth", "M_WeekLabel"]]
                .dropna()
                .drop_duplicates()
                .sort_values(["M_MonthKey", "M_WeekOfMonth"])
            )
            w_options = w_map["M_WeekLabel"].tolist()

            w_sel = st.multiselect(
                "Selecciona Semana(s) del mes (Ventas)",
                options=w_options,
                default=w_options,
                key="sem_mvw_weeks_multi",
            )

            if w_sel:
                df_mvw = df_mvw[df_mvw["M_WeekLabel"].isin(w_sel)].copy()
            else:
                df_mvw = df_mvw.iloc[0:0].copy()

            if df_mvw.empty:
                st.info("No hay datos para las semanas seleccionadas.")
            else:
                by_day = df_mvw.groupby("M_Fecha", as_index=False).size().rename(columns={"size": "Ventas"})
                by_day_export = by_day.copy()  # ✅ ADDED (export)
                fig_mvw = px.bar(
                    by_day,
                    x="M_Fecha",
                    y="Ventas",
                    title="Total por día (Ventas) — filtro por Mes(es) y Semana(s)",
                    labels={"Ventas": "Ventas", "M_Fecha": "Fecha"},
                    template=PLOTLY_TEMPLATE,
                )
                fig_mvw.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                add_bar_value_labels(fig_mvw)
                st.plotly_chart(fig_mvw, width="stretch", key="t1_sem_mvw_bar")

                st.markdown("### Comparativo día vs día (entre meses seleccionados) — Ventas")

                df_mvw["M_DiaDelMes"] = df_mvw["M_DT"].dt.day
                cmp = (
                    df_mvw.groupby(["M_MonthLabel", "M_DiaDelMes"], as_index=False)
                    .size()
                    .rename(columns={"size": "Ventas"})
                )
                cmp_export = cmp.copy()  # ✅ ADDED (export)

                fig_cmp = px.line(
                    cmp,
                    x="M_DiaDelMes",
                    y="Ventas",
                    color="M_MonthLabel",
                    markers=True,
                    title="Comparativo por día del mes (Ventas)",
                    labels={"M_DiaDelMes": "Día del mes", "Ventas": "Ventas", "M_MonthLabel": "Mes"},
                    template=PLOTLY_TEMPLATE,
                )
                fig_cmp.update_xaxes(dtick=1)
                fig_cmp.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_cmp, width="stretch", key="t1_sem_mvw_day_vs_day")

                st.markdown("#### Comparar dos intervalos de tiempo ")

                def _trim_time_to_minute_sem(tobj):
                    try:
                        return tobj.replace(second=0, microsecond=0)
                    except Exception:
                        return tobj

                def _t_sem(h: int, m: int):
                    return datetime.strptime(f"{h:02d}:{m:02d}", "%H:%M").time()

                avail_dts = df_mvw["M_DT"].dropna()
                avail_dates = sorted(df_mvw["M_Fecha"].dropna().unique().tolist())

                if avail_dts.empty or not avail_dates:
                    st.info("No hay fechas/horas disponibles para comparar con los filtros actuales (Semanas).")
                else:
                    min_dt = avail_dts.min().to_pydatetime()
                    max_dt = avail_dts.max().to_pydatetime()

                    d_def_2 = avail_dates[-1]
                    d_def_1 = avail_dates[-2] if len(avail_dates) >= 2 else avail_dates[-1]

                    s1_def = max(datetime.combine(d_def_1, _t_sem(0, 0)), min_dt)
                    e1_def = min(datetime.combine(d_def_1, _t_sem(23, 59)), max_dt)
                    s2_def = max(datetime.combine(d_def_2, _t_sem(0, 0)), min_dt)
                    e2_def = min(datetime.combine(d_def_2, _t_sem(23, 59)), max_dt)

                    if s1_def > e1_def:
                        s1_def, e1_def = min_dt, max_dt
                    if s2_def > e2_def:
                        s2_def, e2_def = min_dt, max_dt

                    cA, cB = st.columns(2, gap="large")

                    with cA:
                        st.markdown("**Fecha 1 (intervalo a comparar)**")
                        a1, a2 = st.columns(2)
                        with a1:
                            s1_date = st.date_input(
                                "Inicio (Fecha 1) — Semanas",
                                value=s1_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i1_start_date",
                            )
                        with a2:
                            s1_time = st.time_input(
                                "Inicio (Hora 1) — Semanas",
                                value=_trim_time_to_minute_sem(s1_def.time()),
                                key="sem_mvw_i1_start_time",
                            )
                        b1, b2 = st.columns(2)
                        with b1:
                            e1_date = st.date_input(
                                "Fin (Fecha 1) — Semanas",
                                value=e1_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i1_end_date",
                            )
                        with b2:
                            e1_time = st.time_input(
                                "Fin (Hora 1) — Semanas",
                                value=_trim_time_to_minute_sem(e1_def.time()),
                                key="sem_mvw_i1_end_time",
                            )

                    with cB:
                        st.markdown("**Fecha 2 (intervalo a comparar)**")
                        a1, a2 = st.columns(2)
                        with a1:
                            s2_date = st.date_input(
                                "Inicio (Fecha 2) — Semanas",
                                value=s2_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i2_start_date",
                            )
                        with a2:
                            s2_time = st.time_input(
                                "Inicio (Hora 2) — Semanas",
                                value=_trim_time_to_minute_sem(s2_def.time()),
                                key="sem_mvw_i2_start_time",
                            )
                        b1, b2 = st.columns(2)
                        with b1:
                            e2_date = st.date_input(
                                "Fin (Fecha 2) — Semanas",
                                value=e2_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i2_end_date",
                            )
                        with b2:
                            e2_time = st.time_input(
                                "Fin (Hora 2) — Semanas",
                                value=_trim_time_to_minute_sem(e2_def.time()),
                                key="sem_mvw_i2_end_time",
                            )

                    s1 = datetime.combine(s1_date, s1_time)
                    e1 = datetime.combine(e1_date, e1_time)
                    s2 = datetime.combine(s2_date, s2_time)
                    e2 = datetime.combine(e2_date, e2_time)

                    if s1 < min_dt: s1 = min_dt
                    if e1 > max_dt: e1 = max_dt
                    if s2 < min_dt: s2 = min_dt
                    if e2 > max_dt: e2 = max_dt

                    if s1 > e1:
                        st.warning("En Fecha 1 (Semanas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                        s1, e1 = e1, s1
                    if s2 > e2:
                        st.warning("En Fecha 2 (Semanas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                        s2, e2 = e2, s2

                    df_d1 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s1)) & (df_mvw["M_DT"] <= pd.Timestamp(e1))].copy()
                    df_d2 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s2)) & (df_mvw["M_DT"] <= pd.Timestamp(e2))].copy()

                    v1 = int(df_d1.shape[0])
                    v2 = int(df_d2.shape[0])

                    m1 = float(df_d1["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d1.columns else 0.0
                    m2 = float(df_d2["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d2.columns else 0.0

                    k1, k2, k3, k4 = st.columns(4, gap="medium")
                    with k1:
                        metric_card("Ventas (Fecha 1)", fmt_int(v1), sub=f"{s1:%d/%m/%Y %H:%M} → {e1:%d/%m/%Y %H:%M}")
                    with k2:
                        metric_card("Ventas (Fecha 2)", fmt_int(v2), sub=f"{s2:%d/%m/%Y %H:%M} → {e2:%d/%m/%Y %H:%M}")
                    with k3:
                        metric_card("Diferencia (1-2)", f"{(v1 - v2):+,}")
                    with k4:
                        metric_card("Diferencia Monto (1-2)", f"{(m1 - m2):+,.2f}")

                    comp_df = pd.DataFrame(
                        {
                            "Comparación": ["Fecha 1", "Fecha 2"],
                            "Ventas": [v1, v2],
                            "MontoVendido": [m1, m2],
                            "Inicio": [s1, s2],
                            "Fin": [e1, e2],
                        }
                    )
                    interval_summary_export = comp_df.copy()  # ✅ ADDED (export)

                    fig_dates = px.bar(
                        comp_df,
                        x="Comparación",
                        y="Ventas",
                        title="Comparativo Ventas — Fecha 1 vs Fecha 2",
                        labels={"Ventas": "Ventas"},
                        hover_data={"Inicio": True, "Fin": True, "MontoVendido": True, "Comparación": False},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_dates.update_layout(height=360, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    add_bar_value_labels(fig_dates)
                    st.plotly_chart(fig_dates, width="stretch", key="t1_sem_mvw_interval_bar")

                    df_d1["H"] = pd.to_datetime(df_d1["M_DT"], errors="coerce").dt.hour
                    df_d2["H"] = pd.to_datetime(df_d2["M_DT"], errors="coerce").dt.hour

                    h1 = df_d1.groupby("H").size()
                    h2 = df_d2.groupby("H").size()

                    hours = list(range(0, 24))
                    hour_df = pd.DataFrame(
                        {
                            "Hora": hours,
                            "Fecha 1": [int(h1.get(h, 0)) for h in hours],
                            "Fecha 2": [int(h2.get(h, 0)) for h in hours],
                        }
                    )
                    hour_breakdown_export = hour_df.copy()  # ✅ ADDED (export)

                    hour_long = hour_df.melt(id_vars="Hora", var_name="Fecha", value_name="Ventas")

                    fig_hour2 = px.bar(
                        hour_long,
                        x="Hora",
                        y="Ventas",
                        color="Fecha",
                        barmode="group",
                        title="Comparativo por hora — Fecha 1 vs Fecha 2",
                        labels={"Ventas": "Ventas", "Hora": "Hora"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_hour2.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    add_bar_value_labels(fig_hour2)
                    st.plotly_chart(fig_hour2, width="stretch", key="t1_sem_mvw_interval_hour")

    # ✅ DOWNLOAD EXCEL (Semanas) — ADDED
    filtros_sem = pd.DataFrame(
        {
            "Filtro": ["Centro", "Supervisor", "Ejecutivo", "Subregión", "Mes (Semanas)"],
            "Valor": [
                ", ".join(center_sel) if center_sel else "Todos",
                ", ".join(sup_sel) if sup_sel else "Todos",
                ", ".join(ej_sel) if ej_sel else "Todos",
                ", ".join(sub_sel) if sub_sel else "Todos",
                str(sem_choice),
            ],
        }
    )

    sheets_sem = {
        "Filtros": filtros_sem,
        "Semanas (serie)": w.copy(),
    }
    if by_day_export is not None:
        sheets_sem["MesSemana - Total por dia"] = by_day_export
    if cmp_export is not None:
        sheets_sem["Comparativo dia vs dia"] = cmp_export
    if interval_summary_export is not None:
        sheets_sem["Intervalo resumen"] = interval_summary_export
    if hour_breakdown_export is not None:
        sheets_sem["Intervalo por hora"] = hour_breakdown_export

    st.download_button(
        "⬇️ Descargar Excel (Semanas)",
        data=build_excel_bytes(sheets_sem),
        file_name=f"Semanas_{mes_sel}_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_semanas_{mes_sel}",
        use_container_width=True,
    )

# ======================================================
# TAB 3: JV vs CC2
# ======================================================
with tabs[2]:
    st.markdown(f"## Ventas del Mes — Comparativo (JV vs CC2) — **{mes_labels[mes_sel]}**")

    df_m = filter_month(ventas, mes_sel)
    if sub_sel:
        df_m = df_m[df_m["SUBREGION"].isin(sub_sel)]
    if sup_sel:
        df_m = df_m[df_m["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_m = df_m[df_m["EJECUTIVO"].isin(ej_sel)]
    if center_sel:
        df_m = df_m[df_m["CentroKey"].isin(center_sel)]

    df_jv = df_m[df_m["CentroKey"] == "JV"].copy()
    df_cc2 = df_m[df_m["CentroKey"] == "CC2"].copy()

    cL, cR = st.columns(2, gap="large")

    def render_group(col, title, df_, key_prefix: str):
        with col:
            st.markdown(f"### {title}")
            a, b, c, d = st.columns(4, gap="medium")
            with a:
                metric_card("Monto Vendido", fmt_money_short(total_precio(df_)))
            with b:
                metric_card("Promedio de Ventas", fmt_int(promedio_diario(df_)))
            with c:
                metric_card("Total de Ventas", fmt_int(total_folios(df_)))
            with d:
                metric_card("Max Ventas en un Día", fmt_int(max_folios_dia(df_)))

            s = daily_series(df_)
            avg = promedio_diario(df_)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=s["Fecha"], y=s["Ventas"], name="Total de Ventas"))
            fig.add_trace(go.Scatter(x=s["Fecha"], y=[avg] * len(s), mode="lines", name="Promedio Diario de Venta"))
            fig.update_layout(title="Vista General de Ventas", height=340)
            apply_plotly_theme(fig)
            st.plotly_chart(fig, width="stretch", key=f"t2_{key_prefix}_general_{mes_sel}")

            top = s.sort_values("Ventas", ascending=False).head(6).copy()
            top["Fecha"] = top["Fecha"].dt.strftime("%A, %d %B %Y")

            # ✅ ADD TOTAL ROW (month total for that center) + ✅ BOLD TOTAL
            top_show = top.rename(columns={"Ventas": "Total de Ventas"})
            top_show = add_totals_row(
                top_show,
                label_col="Fecha",
                totals={"Total de Ventas": total_folios(df_)},
                label="TOTAL (Mes)",
            )

            st.dataframe(
                style_totals_bold(top_show, label_col="Fecha").format({"Total de Ventas": "{:,.0f}"}),
                hide_index=True,
                width="stretch",
            )

            kpi_mini("ARPU", fmt_money_short(arpu(df_)))
            kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_)))

    render_group(cL, "JV (Juárez)", df_jv, key_prefix="jv")
    render_group(cR, "CC2 (Center 2)", df_cc2, key_prefix="cc2")

# ======================================================
# TAB 4: Mes vs Mes  ✅ + Vista por meses/semanas + día vs día + intervalo
# ======================================================
with tabs[3]:
    st.markdown("## Mes vs Mes")

    months_all = sorted(ventas["AñoMes"].dropna().unique().tolist())
    if len(months_all) < 2:
        st.warning("Se requieren al menos 2 meses para comparar.")
    else:
        colS1, colS2, colS3 = st.columns([0.35, 0.35, 0.30])
        with colS1:
            mes_actual = st.selectbox(
                "Mes Actual",
                options=months_all,
                format_func=lambda ym: mes_labels.get(ym, str(ym)),
                index=len(months_all) - 1,
                key="mvm_mes_actual",
            )
        with colS2:
            mes_comp = st.selectbox(
                "Mes Comparado",
                options=months_all,
                format_func=lambda ym: mes_labels.get(ym, str(ym)),
                index=max(0, len(months_all) - 2),
                key="mvm_mes_comp",
            )
        with colS3:
            modo = st.selectbox(
                "ModoComparación",
                options=[0, 1],
                format_func=lambda x: "0 = Mes completo" if x == 0 else "1 = Cortar al día de hoy",
                key="mvm_modo",
            )

        dfA = filter_month(ventas, mes_actual)
        dfB = filter_month(ventas, mes_comp)

        if center_sel:
            dfA = dfA[dfA["CentroKey"].isin(center_sel)]
            dfB = dfB[dfB["CentroKey"].isin(center_sel)]
        if sup_sel:
            dfA = dfA[dfA["Supervisor"].isin(sup_sel)]
            dfB = dfB[dfB["Supervisor"].isin(sup_sel)]
        if ej_sel:
            dfA = dfA[dfA["EJECUTIVO"].isin(ej_sel)]
            dfB = dfB[dfB["EJECUTIVO"].isin(ej_sel)]
        if sub_sel:
            dfA = dfA[dfA["SUBREGION"].isin(sub_sel)]
            dfB = dfB[dfB["SUBREGION"].isin(sub_sel)]

        today_day = date.today().day
        dfA_m = cut_month_mode(dfA, modo, today_day)
        dfB_m = cut_month_mode(dfB, modo, today_day)

        ventasA = total_folios(dfA_m)
        ventasB = total_folios(dfB_m)

        dif = ventasA - ventasB
        pct = (dif / ventasB) if ventasB else np.nan
        arrow = "↑" if dif > 0 else ("↓" if dif < 0 else "→")

        st.markdown(
            f"""
            <div class="metric-card" style="text-align:center;">
              <div class="metric-value" style="font-size:1.6rem;">
                {dif:+,} ventas ({fmt_pct(pct)}) {arrow}
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        c1, c2 = st.columns(2, gap="large")
        with c1:
            metric_card("Promedio Diario", fmt_int(promedio_diario(dfA_m)))
            metric_card("Ventas", fmt_int(ventasA))
            sA = daily_series(dfA_m)
            avgA = promedio_diario(dfA_m)
            figA = go.Figure()
            figA.add_trace(go.Bar(x=sA["Fecha"], y=sA["Ventas"], name="Total Folios (Modo)"))
            figA.add_trace(go.Scatter(x=sA["Fecha"], y=[avgA] * len(sA), mode="lines", name="Promedio Diario Mes (Modo)"))
            figA.update_layout(title="Vista General de Ventas", height=340)
            apply_plotly_theme(figA)
            st.plotly_chart(figA, width="stretch", key=f"t3_mesA_{mes_actual}_{modo}")

        with c2:
            metric_card("Promedio Diario", fmt_int(promedio_diario(dfB_m)))
            metric_card("Ventas", fmt_int(ventasB))
            sB = daily_series(dfB_m)
            avgB = promedio_diario(dfB_m)
            figB = go.Figure()
            figB.add_trace(go.Bar(x=sB["Fecha"], y=sB["Ventas"], name="Total Folios (Modo)"))
            figB.add_trace(go.Scatter(x=sB["Fecha"], y=[avgB] * len(sB), mode="lines", name="Promedio Diario Mes (Modo)"))
            figB.update_layout(title="Vista General de Ventas", height=340)
            add_bar_value_labels(figB)
            apply_plotly_theme(figB)
            st.plotly_chart(figB, width="stretch", key=f"t3_mesB_{mes_comp}_{modo}")

        emoji = "📈🔥" if dif > 0 else ("📉⚠️" if dif < 0 else "➖")
        msg = f"{emoji} {mes_labels[mes_actual]} {'mejoró' if dif>0 else ('bajó' if dif<0 else 'mantiene')} {fmt_pct(pct)} vs {mes_labels[mes_comp]} ({dif:+,} ventas)."
        st.markdown(
            f"""
            <div class="metric-card" style="text-align:center;">
              <div class="metric-value" style="font-size:1.25rem;">{msg}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ==========================================================
        # ✅ NEW: Vista por meses y semanas + comparativo día vs día + intervalo (como tu ejemplo)
        # ==========================================================
        st.markdown("---")
        st.markdown("## ✅ Vista por meses y semanas ")

        df_mvm_ctx = ventas.copy()
        if center_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["CentroKey"].isin(center_sel)]
        if sup_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["Supervisor"].isin(sup_sel)]
        if ej_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["EJECUTIVO"].isin(ej_sel)]
        if sub_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["SUBREGION"].isin(sub_sel)]

        df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["FECHA DE CAPTURA"].notna()].copy()
        df_mvm_ctx["M_DT"] = pd.to_datetime(df_mvm_ctx["FECHA DE CAPTURA"], errors="coerce")
        df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["M_DT"].notna()].copy()

        df_mvm_ctx["M_Fecha"] = df_mvm_ctx["M_DT"].dt.date
        df_mvm_ctx["M_Hora"] = df_mvm_ctx["M_DT"].dt.hour

        if df_mvm_ctx.empty:
            st.info("No hay datos disponibles para la vista por meses/semanas con los filtros actuales.")
        else:
            df_mvm_ctx["M_MonthKey"] = df_mvm_ctx["M_DT"].dt.strftime("%Y-%m")
            df_mvm_ctx["M_MonthName"] = df_mvm_ctx["M_DT"].dt.strftime("%B")
            df_mvm_ctx["M_MonthLabel"] = df_mvm_ctx["M_MonthKey"] + " (" + df_mvm_ctx["M_MonthName"] + ")"

            month_start = df_mvm_ctx["M_DT"].dt.to_period("M").dt.to_timestamp()
            first_wd = month_start.dt.weekday
            df_mvm_ctx["M_WeekOfMonth"] = ((df_mvm_ctx["M_DT"].dt.day + first_wd - 1) // 7) + 1

            st.markdown("### Vista por meses y semanas")

            m_options = sorted(df_mvm_ctx["M_MonthLabel"].dropna().unique().tolist())

            # Default: los 2 meses del comparativo (si existen), si no: todos
            def _ym_to_monthkey(ym: int) -> str:
                y = ym // 100
                m = ym % 100
                return f"{y}-{m:02d}"

            want_keys = {_ym_to_monthkey(int(mes_actual)), _ym_to_monthkey(int(mes_comp))}
            defaults = [x for x in m_options if str(x).startswith(tuple(sorted(want_keys)))]
            if not defaults:
                defaults = m_options

            m_sel = st.multiselect(
                "Selecciona uno o más meses (Ventas)",
                options=m_options,
                default=defaults,
                key="mvm_months_multi",
            )

            df_mvw = df_mvm_ctx.copy()
            if m_sel:
                df_mvw = df_mvw[df_mvw["M_MonthLabel"].isin(m_sel)].copy()
            else:
                df_mvw = df_mvw.iloc[0:0].copy()

            if df_mvw.empty:
                st.info("No hay datos para los meses seleccionados.")
            else:
                df_mvw["M_WeekLabel"] = df_mvw["M_MonthLabel"] + " - Semana " + df_mvw["M_WeekOfMonth"].astype(int).astype(str)
                w_options = sorted(df_mvw["M_WeekLabel"].dropna().unique().tolist())
                w_sel_default = w_options

                w_sel = st.multiselect(
                    "Selecciona Semana(s) del mes (Ventas)",
                    options=w_options,
                    default=w_sel_default,
                    key="mvm_weeks_multi",
                )

                if w_sel:
                    df_mvw = df_mvw[df_mvw["M_WeekLabel"].isin(w_sel)].copy()
                else:
                    df_mvw = df_mvw.iloc[0:0].copy()

                if df_mvw.empty:
                    st.info("No hay datos para las semanas seleccionadas.")
                else:
                    by_day = df_mvw.groupby("M_Fecha", as_index=False).size().rename(columns={"size": "Ventas"})
                    fig_mvw = px.bar(
                        by_day,
                        x="M_Fecha",
                        y="Ventas",
                        title="Total por día (Ventas) — filtro por Mes(es) y Semana(s)",
                        labels={"Ventas": "Ventas", "M_Fecha": "Fecha"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_mvw.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    add_bar_value_labels(fig_mvw)
                    st.plotly_chart(fig_mvw, width="stretch", key="t3_mvw_bar")

                    st.markdown("### Comparativo día vs día (mes contra mes) — Ventas")

                    df_mvw["M_DiaDelMes"] = df_mvw["M_DT"].dt.day
                    cmp = (
                        df_mvw.groupby(["M_MonthLabel", "M_DiaDelMes"], as_index=False)
                        .size()
                        .rename(columns={"size": "Ventas"})
                    )

                    fig_cmp = px.line(
                        cmp,
                        x="M_DiaDelMes",
                        y="Ventas",
                        color="M_MonthLabel",
                        markers=True,
                        title="Comparativo por día del mes (Ventas)",
                        labels={"M_DiaDelMes": "Día del mes", "Ventas": "Ventas", "M_MonthLabel": "Mes"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_cmp.update_xaxes(dtick=1)
                    fig_cmp.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_cmp, width="stretch", key="t3_mvw_day_vs_day")

                    st.markdown("#### Comparar dos intervalos de tiempo")

                    def _trim_time_to_minute(tobj):
                        try:
                            return tobj.replace(second=0, microsecond=0)
                        except Exception:
                            return tobj

                    def _t(h: int, m: int):
                        return datetime.strptime(f"{h:02d}:{m:02d}", "%H:%M").time()

                    avail_dts = df_mvw["M_DT"].dropna()
                    avail_dates = sorted(df_mvw["M_Fecha"].dropna().unique().tolist())

                    if avail_dts.empty or not avail_dates:
                        st.info("No hay fechas/horas disponibles para comparar con los filtros actuales (Ventas).")
                    else:
                        min_dt = avail_dts.min().to_pydatetime()
                        max_dt = avail_dts.max().to_pydatetime()

                        d_def_2 = avail_dates[-1]
                        d_def_1 = avail_dates[-2] if len(avail_dates) >= 2 else avail_dates[-1]

                        s1_def = max(datetime.combine(d_def_1, _t(0, 0)), min_dt)
                        e1_def = min(datetime.combine(d_def_1, _t(23, 59)), max_dt)
                        s2_def = max(datetime.combine(d_def_2, _t(0, 0)), min_dt)
                        e2_def = min(datetime.combine(d_def_2, _t(23, 59)), max_dt)

                        if s1_def > e1_def:
                            s1_def, e1_def = min_dt, max_dt
                        if s2_def > e2_def:
                            s2_def, e2_def = min_dt, max_dt

                        cA, cB = st.columns(2, gap="large")

                        with cA:
                            st.markdown("**Fecha 1 (intervalo a comparar)**")
                            a1, a2 = st.columns(2)
                            with a1:
                                s1_date = st.date_input(
                                    "Inicio (Fecha 1) — Ventas",
                                    value=s1_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i1_start_date",
                                )
                            with a2:
                                s1_time = st.time_input(
                                    "Inicio (Hora 1) — Ventas",
                                    value=_trim_time_to_minute(s1_def.time()),
                                    key="mvm_i1_start_time",
                                )
                            b1, b2 = st.columns(2)
                            with b1:
                                e1_date = st.date_input(
                                    "Fin (Fecha 1) — Ventas",
                                    value=e1_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i1_end_date",
                                )
                            with b2:
                                e1_time = st.time_input(
                                    "Fin (Hora 1) — Ventas",
                                    value=_trim_time_to_minute(e1_def.time()),
                                    key="mvm_i1_end_time",
                                )

                        with cB:
                            st.markdown("**Fecha 2 (intervalo a comparar)**")
                            a1, a2 = st.columns(2)
                            with a1:
                                s2_date = st.date_input(
                                    "Inicio (Fecha 2) — Ventas",
                                    value=s2_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i2_start_date",
                                )
                            with a2:
                                s2_time = st.time_input(
                                    "Inicio (Hora 2) — Ventas",
                                    value=_trim_time_to_minute(s2_def.time()),
                                    key="mvm_i2_start_time",
                                )
                            b1, b2 = st.columns(2)
                            with b1:
                                e2_date = st.date_input(
                                    "Fin (Fecha 2) — Ventas",
                                    value=e2_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i2_end_date",
                                )
                            with b2:
                                e2_time = st.time_input(
                                    "Fin (Hora 2) — Ventas",
                                    value=_trim_time_to_minute(e2_def.time()),
                                    key="mvm_i2_end_time",
                                )

                        s1 = datetime.combine(s1_date, s1_time)
                        e1 = datetime.combine(e1_date, e1_time)
                        s2 = datetime.combine(s2_date, s2_time)
                        e2 = datetime.combine(e2_date, e2_time)

                        if s1 < min_dt: s1 = min_dt
                        if e1 > max_dt: e1 = max_dt
                        if s2 < min_dt: s2 = min_dt
                        if e2 > max_dt: e2 = max_dt

                        if s1 > e1:
                            st.warning("En Fecha 1 (Ventas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                            s1, e1 = e1, s1
                        if s2 > e2:
                            st.warning("En Fecha 2 (Ventas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                            s2, e2 = e2, s2

                        df_d1 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s1)) & (df_mvw["M_DT"] <= pd.Timestamp(e1))].copy()
                        df_d2 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s2)) & (df_mvw["M_DT"] <= pd.Timestamp(e2))].copy()

                        v1 = int(df_d1.shape[0])
                        v2 = int(df_d2.shape[0])

                        m1 = float(df_d1["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d1.columns else 0.0
                        m2 = float(df_d2["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d2.columns else 0.0

                        k1, k2, k3, k4 = st.columns(4, gap="medium")
                        with k1:
                            metric_card("Ventas (Fecha 1)", fmt_int(v1), sub=f"{s1:%d/%m/%Y %H:%M} → {e1:%d/%m/%Y %H:%M}")
                        with k2:
                            metric_card("Ventas (Fecha 2)", fmt_int(v2), sub=f"{s2:%d/%m/%Y %H:%M} → {e2:%d/%m/%Y %H:%M}")
                        with k3:
                            metric_card("Diferencia (1-2)", f"{(v1 - v2):+,}")
                        with k4:
                            metric_card("Diferencia Monto (1-2)", f"{(m1 - m2):+,.2f}")

                        comp_df = pd.DataFrame(
                            {
                                "Comparación": ["Fecha 1", "Fecha 2"],
                                "Ventas": [v1, v2],
                                "MontoVendido": [m1, m2],
                                "Inicio": [s1, s2],
                                "Fin": [e1, e2],
                            }
                        )
                        fig_dates = px.bar(
                            comp_df,
                            x="Comparación",
                            y="Ventas",
                            title="Comparativo Ventas — Fecha 1 vs Fecha 2",
                            labels={"Ventas": "Ventas"},
                            hover_data={"Inicio": True, "Fin": True, "MontoVendido": True, "Comparación": False},
                            template=PLOTLY_TEMPLATE,
                        )
                        fig_dates.update_layout(height=360, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        add_bar_value_labels(fig_dates)
                        st.plotly_chart(fig_dates, width="stretch", key="t3_mvw_interval_bar")

                        df_d1["H"] = pd.to_datetime(df_d1["M_DT"], errors="coerce").dt.hour
                        df_d2["H"] = pd.to_datetime(df_d2["M_DT"], errors="coerce").dt.hour

                        h1 = df_d1.groupby("H").size()
                        h2 = df_d2.groupby("H").size()

                        hours = list(range(0, 24))
                        hour_df = pd.DataFrame(
                            {
                                "Hora": hours,
                                "Fecha 1": [int(h1.get(h, 0)) for h in hours],
                                "Fecha 2": [int(h2.get(h, 0)) for h in hours],
                            }
                        )
                        hour_long = hour_df.melt(id_vars="Hora", var_name="Fecha", value_name="Ventas")

                        fig_hour = px.bar(
                            hour_long,
                            x="Hora",
                            y="Ventas",
                            color="Fecha",
                            barmode="group",
                            title="Comparativo por hora — Fecha 1 vs Fecha 2",
                            labels={"Ventas": "Ventas", "Hora": "Hora"},
                            template=PLOTLY_TEMPLATE,
                        )
                        fig_hour.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_hour, width="stretch", key="t3_mvw_interval_hour")

# ======================================================
# TAB 5: Detalle (Metas + Tránsito + Resumen Supervisor)
#   ✅ FIX 1: evita StreamlitDuplicateElementKey (keys únicos en plotly/download)
#   ✅ FIX 2: NO mostrar NINGÚN registro de estos "boss" (no supervisores)
#   ✅ FIX 3: NO tocar/mutar EXCLUDED_SUP_NORMS global (solo TAB 5)
#   ✅ FIX 4: ARREGLA INDENTACIÓN (tu bloque de plots estaba indentado de más)
# ======================================================
with tabs[4]:
    st.markdown(f"## Detalle General de Ventas — **{mes_labels[mes_sel]}**")

    # ------------------------------------------------
    # 0) Boss supervisors to exclude from TAB 5
    # ------------------------------------------------
    _BOSS_SUPS = [
        "BLANCA LETICIA HERNANDEZ CARRILLO",
        #"MARIA LUISA MEZA GOEL",
        "ANTONIO HERNAN GOMEZ OLVERA",
        "JONATHAN ARTURO TENORIO DEL AGUILA",
        "NELLY MASHIEL CAMPOS JUAREZ",
    ]
    _BOSS_SUP_NORMS = set(normalize_name(x) for x in _BOSS_SUPS)

    # ✅ local-only exclude list for this TAB (does NOT change global var)
    _EXCL_TAB5 = set(EXCLUDED_SUP_NORMS) if "EXCLUDED_SUP_NORMS" in globals() else set()
    EXCLUDED_SUP_NORMS_TAB5 = _EXCL_TAB5 | _BOSS_SUP_NORMS

    # ------------------------------------------------
    # 1) External data (Metas & Tránsito)
    # ------------------------------------------------
    metas_map = load_metas_from_csv(int(mes_sel))

    # Query tránsito only for the selected month window
    m_start, m_end = month_bounds(int(mes_sel))
    
    # ⚡ Load cached history instantly
    prog_history = fetch_programacion_history(end_dt)
    
    transito_map = {}
    activadas_map = {}

    if not prog_history.empty:
        # ✅ Use ALL fetched history for traceability, not only selected month
        df_hist = prog_history.copy()

        # Calculate Transito Map
        transito_mask = (
            df_hist["Estatus_upper"].isin([
                'EN ENTREGA', 'EN PREPARACION', 'EN PREPARACIÓN',
                'SOLICITADO', 'BACK OFFICE', 'BACKOFFICE',
                'EN TRANSITO', 'EN TRÁNSITO'
            ])
        ) | (
            (df_hist["Estatus_upper"] == 'ENTREGADO') & df_hist["Venta_Vacia"]
        )

        if not df_hist[transito_mask].empty:
            transito_map = df_hist[transito_mask].groupby("EJ_NORM").size().to_dict()

        # Calculate Activadas Map
        activadas_mask = (df_hist["Estatus_upper"] == 'ENTREGADO') & (~df_hist["Venta_Vacia"])
        if not df_hist[activadas_mask].empty:
            activadas_map = df_hist[activadas_mask].groupby("EJ_NORM").size().to_dict()

    metas_sup_map = load_metas_supervisor_from_excel(int(mes_sel))

    # ------------------------------------------------
    # 2) Working days (Sanity logic)
    # ------------------------------------------------
    today = date.today()

    dias_hab_total_real = workable_equiv_between(m_start, m_end)

    # Remaining days from today within the selected month
    if today < m_start:
        cutoff_start = m_start
    elif today > m_end:
        cutoff_start = m_end
    else:
        cutoff_start = today

    dias_hab_restantes_real = workable_equiv_between(cutoff_start, m_end)

    # ✅ Solo para abril, fijar Días Hab Mes en 22.5
    if int(mes_sel) % 100 == 4:
        dias_hab_total = 22.5
        dias_hab_restantes = float(dias_hab_restantes_real)
        dias_hab_transcurridos = float(max(0.0, dias_hab_total - dias_hab_restantes))
    else:
        dias_hab_total = dias_hab_total_real
        dias_hab_restantes = dias_hab_restantes_real
        dias_hab_transcurridos = float(max(0.0, dias_hab_total - dias_hab_restantes))

    if dias_hab_total <= 0:
        dias_hab_total = 1.0
    if dias_hab_restantes < 0:
        dias_hab_restantes = 0.0

    # ------------------------------------------------
    # 3) Build detalle per Ejecutivo (INCLUDE 0 SALES)
    #     Build roster from empleados, then left-join sales stats
    # ------------------------------------------------
    df_d = df_base.copy()

    # ---- ROSTER (employees who should appear even with 0 ventas) ----
    emp_roster = empleados.copy()



    # ✅ Include ACTIVO + anyone with Programadas (EnTransito/Activadas) to match dashboard2
    emp_roster["Estatus"] = emp_roster["Estatus"].astype(str).str.strip().str.upper()
    _keep_norms = set(list(transito_map.keys()) + list(activadas_map.keys()))
    emp_roster["_EJ_NORM_TMP"] = emp_roster["Nombre"].astype(str).str.strip().apply(normalize_name)
    emp_roster = emp_roster[(emp_roster["Estatus"] == "ACTIVO") | (emp_roster["_EJ_NORM_TMP"].isin(_keep_norms))].copy()
    emp_roster.drop(columns=["_EJ_NORM_TMP"], inplace=True, errors="ignore")

    # Keep only executivos (avoid supervisors/coordinators)
    emp_roster["Puesto"] = emp_roster["Puesto"].astype(str).str.strip().str.upper()
    ml_norm_tab5 = normalize_name("MARIA LUISA MEZA GOEL")
    emp_roster = emp_roster[
        (~emp_roster["Puesto"].str.contains("SUPERV", na=False))
        & (~emp_roster["Puesto"].str.contains("COORD", na=False))
        & (~emp_roster["Puesto"].str.contains("GEREN", na=False))
        & (~emp_roster["Puesto"].str.contains("JEFE", na=False))
        | (emp_roster["Jefe Inmediato"].astype(str).apply(normalize_name) == ml_norm_tab5) # ✅ Salva a su equipo
    ].copy()

    # Normalize fields
    emp_roster["Nombre"] = emp_roster["Nombre"].astype(str).str.strip()
    emp_roster["Jefe Inmediato"] = emp_roster["Jefe Inmediato"].astype(str).str.strip()
    emp_roster["Centro"] = emp_roster["Centro"].astype(str).str.strip()

    # CentroKey mapping (same logic as ventas)
    emp_roster["CentroKey"] = np.where(
        emp_roster["Centro"].str.upper().str.contains("JUAREZ", na=False),
        "JV",
        "CC2",
    )

    roster = emp_roster.rename(columns={"Nombre": "EJECUTIVO", "Jefe Inmediato": "Supervisor"})[
        ["CentroKey", "Supervisor", "EJECUTIVO"]
    ].copy()

    roster["EJECUTIVO_norm"] = roster["EJECUTIVO"].apply(normalize_name)
    roster["Supervisor_norm"] = roster["Supervisor"].apply(normalize_name)

    # ✅ Exclude supervisors requested ONLY in TAB 5
    roster = roster[~roster["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS_TAB5)].copy()

    # Apply SAME sidebar filters to roster
    if center_sel:
        roster = roster[roster["CentroKey"].isin(center_sel)].copy()
    if sup_sel:
        roster = roster[roster["Supervisor"].isin(sup_sel)].copy()
    if ej_sel:
        roster = roster[roster["EJECUTIVO"].isin(ej_sel)].copy()

    # ---- SALES STATS (only those who sold) ----
    if df_d.empty:
        stats = pd.DataFrame(
            columns=[
                "CentroKey",
                "Supervisor",
                "EJECUTIVO",
                "Ventas",
                "MontoVendido",
                "EJECUTIVO_norm",
                "Supervisor_norm",
            ]
        )
    else:
        # ✅ extra safety: remove bosses from sales base too
        df_d["Supervisor"] = df_d["Supervisor"].astype(str).str.strip()
        df_d["Supervisor_norm"] = df_d["Supervisor"].apply(normalize_name)
        df_d = df_d[~df_d["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS_TAB5)].copy()

        stats = (
            df_d.groupby(["CentroKey", "Supervisor", "EJECUTIVO"], as_index=False)
            .agg(
                Ventas=("FOLIO", "count"),
                MontoVendido=("PRECIO", "sum"),
            )
            .copy()
        )

        stats["EJECUTIVO"] = stats["EJECUTIVO"].astype(str).str.strip()
        stats["Supervisor"] = stats["Supervisor"].astype(str).str.strip()
        stats["EJECUTIVO_norm"] = stats["EJECUTIVO"].apply(normalize_name)
        stats["Supervisor_norm"] = stats["Supervisor"].apply(normalize_name)

        # ✅ Exclude bosses in stats too
        stats = stats[~stats["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS_TAB5)].copy()

    # ---- LEFT JOIN roster <- stats  (THIS IS THE KEY) ----
    base_matrix = roster.merge(
        stats[["CentroKey", "Supervisor_norm", "EJECUTIVO_norm", "Ventas", "MontoVendido"]],
        on=["CentroKey", "Supervisor_norm", "EJECUTIVO_norm"],
        how="left",
    )

    base_matrix["Supervisor"] = base_matrix["Supervisor"].fillna("")
    base_matrix["EJECUTIVO"] = base_matrix["EJECUTIVO"].fillna("")
    base_matrix["Ventas"] = pd.to_numeric(base_matrix["Ventas"], errors="coerce").fillna(0).astype(int)
    base_matrix["MontoVendido"] = pd.to_numeric(base_matrix["MontoVendido"], errors="coerce").fillna(0.0).astype(float)

    # ------------------------------------------------
    # 3b) Build rows from base_matrix (includes 0 ventas)
    # ------------------------------------------------
    rows = []
    for _, r in base_matrix.iterrows():
        ej_name = str(r["EJECUTIVO"]).strip()
        ej_norm = normalize_name(ej_name)

        # ✅ TRUE sales for the month (from reporte_ventas_no_conciliadas via df_base/stats)
        ventas_reales = int(r.get("Ventas", 0) or 0)

        monto = float(r["MontoVendido"])
        arpu_val = (monto / ventas_reales) if ventas_reales > 0 else 0.0

        meta_val = float(metas_map.get(ej_norm, 0) or 0)
        transito_val = int(transito_map.get(ej_norm, 0) or 0)

        expected_to_date = (float(meta_val) / float(dias_hab_total)) * float(dias_hab_transcurridos) if dias_hab_total > 0 else 0.0
        gap_to_date = float(expected_to_date) - float(ventas_reales)


        gap_to_meta = meta_val - ventas_reales
        daily_needed_avg = (meta_val / dias_hab_total) if dias_hab_total > 0 else 0.0

        if dias_hab_restantes > 0 and gap_to_meta > 0:
            daily_needed_now = gap_to_meta / dias_hab_restantes
        else:
            daily_needed_now = 0.0

        rows.append(
            {
                "CentroKey": r["CentroKey"],
                "Supervisor": str(r["Supervisor"]).strip(),
                "Ejecutivo": ej_name,
                "Meta": float(meta_val),
                "Ventas": ventas_reales,
                "Monto Vendido": monto,
                "ARPU": float(arpu_val),
                "Gap": float(gap_to_date),
                "En Transito": int(transito_val),
                "Dias Hab Mes": float(dias_hab_total),
                "Ventas Diarias Necesarias (Avg)": float(daily_needed_avg),
                "Dias Hab Restantes": float(dias_hab_restantes),
                "Ventas Diarias Necesarias (Hoy)": float(daily_needed_now),
            }
        )

    df_detalle_full = pd.DataFrame(rows)

    # final safety: remove any boss supervisor if something slipped
    if not df_detalle_full.empty:
        df_detalle_full["SUP_NORM"] = df_detalle_full["Supervisor"].astype(str).apply(normalize_name)
        df_detalle_full = df_detalle_full[~df_detalle_full["SUP_NORM"].isin(EXCLUDED_SUP_NORMS_TAB5)].copy()
        df_detalle_full.drop(columns=["SUP_NORM"], inplace=True, errors="ignore")

    # ------------------------------------------------
    # 4) Resumen por Supervisor (por centro)
    # ------------------------------------------------
    df_sup = pd.DataFrame()
    if not df_detalle_full.empty:
        df_sup = (
            df_detalle_full.groupby(["CentroKey", "Supervisor"], as_index=False)
            .agg(
                Meta_Ejecutivos=("Meta", "sum"),
                Ventas=("Ventas", "sum"),
                Monto_Vendido=("Monto Vendido", "sum"),
                En_Transito=("En Transito", "sum"),
            )
            .copy()
        )

        df_sup["SUP_NORM"] = df_sup["Supervisor"].astype(str).apply(normalize_name)
        df_sup = df_sup[~df_sup["SUP_NORM"].isin(EXCLUDED_SUP_NORMS_TAB5)].copy()

        df_sup["Meta Supervisor"] = df_sup["SUP_NORM"].map(lambda k: metas_sup_map.get(k, np.nan))
        df_sup["Meta Supervisor"] = df_sup["Meta Supervisor"].fillna(df_sup["Meta_Ejecutivos"])

        df_sup["Expected_To_Date"] = np.where(
            dias_hab_total > 0,
            (df_sup["Meta Supervisor"].astype(float) / float(dias_hab_total)) * float(dias_hab_transcurridos),
            0.0,
        )

        df_sup["Gap"] = df_sup["Expected_To_Date"].astype(float) - df_sup["Ventas"].astype(float)


        df_sup["Gap"] = df_sup["Expected_To_Date"] - df_sup["Ventas"]
        df_sup["Gap_Meta"] = df_sup["Meta Supervisor"] - df_sup["Ventas"]
        df_sup["ARPU"] = np.where(df_sup["Ventas"] > 0, df_sup["Monto_Vendido"] / df_sup["Ventas"], 0.0)

        df_sup["Dias Hab Mes"] = float(dias_hab_total)
        df_sup["Dias Hab Restantes"] = float(dias_hab_restantes)
        df_sup["Ventas Diarias Necesarias (Hoy)"] = np.where(
            (df_sup["Gap_Meta"] > 0) & (dias_hab_restantes > 0),
            df_sup["Gap_Meta"] / dias_hab_restantes,
            0.0,
        )

    # ------------------------------------------------
    # 5) Styling helpers (Gap red, Tránsito yellow, TOTAL row bold)
    # ------------------------------------------------
    def _style_detalle(df_cols):
        def _row_style(row):
            styles = [""] * len(df_cols)

            is_total = str(row.get("Ejecutivo", "")).strip().upper() == "TOTAL"
            if is_total:
                return ["font-weight: 900; background-color: rgba(127,127,127,0.12);"] * len(df_cols)

            try:
                gap_val = float(row.get("Gap", 0) or 0)
                if gap_val > 0 and "Gap" in df_cols:
                    idx_gap = df_cols.index("Gap")
                    styles[idx_gap] = "background-color: rgba(211,47,47,0.18); font-weight: 800;"
            except Exception:
                pass

            try:
                tr_val = float(row.get("En Transito", 0) or 0)
                if tr_val > 0 and "En Transito" in df_cols:
                    idx_tr = df_cols.index("En Transito")
                    styles[idx_tr] = "background-color: rgba(255,235,59,0.35); font-weight: 800;"
            except Exception:
                pass

            return styles

        return _row_style

    def _style_sup(df_cols):
        def _row_style(row):
            styles = [""] * len(df_cols)

            is_total = str(row.get("Supervisor", "")).strip().upper() == "TOTAL"
            if is_total:
                return ["font-weight: 900; background-color: rgba(127,127,127,0.12);"] * len(df_cols)

            try:
                gap_val = float(row.get("Gap", 0) or 0)
                if gap_val > 0 and "Gap" in df_cols:
                    idx_gap = df_cols.index("Gap")
                    styles[idx_gap] = "background-color: rgba(211,47,47,0.18); font-weight: 800;"
            except Exception:
                pass

            try:
                tr_val = float(row.get("En_Transito", 0) or 0)
                if tr_val > 0 and "En_Transito" in df_cols:
                    idx_tr = df_cols.index("En_Transito")
                    styles[idx_tr] = "background-color: rgba(255,235,59,0.35); font-weight: 800;"
            except Exception:
                pass

            return styles

        return _row_style

    # ------------------------------------------------
    # 6) Render detalle tables (CC2 then JV)
    # ------------------------------------------------
    def display_center_table(center_code: str, title: str):
        st.markdown(f"### {title}")

        if df_detalle_full.empty:
            st.info("No hay datos para este periodo.")
            return

        df_c = df_detalle_full[df_detalle_full["CentroKey"] == center_code].copy()
        if df_c.empty:
            st.info(f"No hay datos para {title}.")
            return

        df_c = df_c.sort_values(["Supervisor", "Gap", "Ventas"], ascending=[True, False, False])

        meta_sum = float(df_c["Meta"].sum())
        ventas_sum = int(df_c["Ventas"].sum())
        monto_sum = float(df_c["Monto Vendido"].sum())
        expected_sum_to_date = (float(meta_sum) / float(dias_hab_total)) * float(dias_hab_transcurridos) if dias_hab_total > 0 else 0.0
        gap_sum = float(expected_sum_to_date) - float(ventas_sum)

        tr_sum = int(df_c["En Transito"].sum())

        total_row = {
            "CentroKey": center_code,
            "Supervisor": "",
            "Ejecutivo": "TOTAL",
            "Meta": meta_sum,
            "Ventas": ventas_sum,
            "Monto Vendido": monto_sum,
            "ARPU": (monto_sum / ventas_sum) if ventas_sum else 0.0,
            "Gap": gap_sum,
            "En Transito": tr_sum,
            "Dias Hab Mes": float(dias_hab_total),
            "Ventas Diarias Necesarias (Avg)": (meta_sum / dias_hab_total) if dias_hab_total else 0.0,
            "Dias Hab Restantes": float(dias_hab_restantes),
            "Ventas Diarias Necesarias (Hoy)": (max(0.0, (meta_sum - ventas_sum)) / dias_hab_restantes) if dias_hab_restantes > 0 else 0.0,
        }

        df_show = pd.concat([df_c, pd.DataFrame([total_row])], ignore_index=True)

        cols_to_show = [
            "Supervisor",
            "Ejecutivo",
            "Meta",
            "Ventas",
            "Monto Vendido",
            "ARPU",
            "Gap",
            "En Transito",
            "Dias Hab Mes",
            "Ventas Diarias Necesarias (Avg)",
            "Dias Hab Restantes",
            "Ventas Diarias Necesarias (Hoy)",
        ]

        st.dataframe(
            df_show[cols_to_show]
            .style
            .apply(_style_detalle(cols_to_show), axis=1)
            .format(
                {
                    "Meta": "{:,.0f}",
                    "Ventas": "{:,.0f}",
                    "Monto Vendido": "${:,.2f}",
                    "ARPU": "${:,.2f}",
                    "Gap": "{:,.2f}",
                    "En Transito": "{:,.0f}",
                    "Dias Hab Mes": "{:,.1f}",
                    "Ventas Diarias Necesarias (Avg)": "{:,.2f}",
                    "Dias Hab Restantes": "{:,.1f}",
                    "Ventas Diarias Necesarias (Hoy)": "{:,.2f}",
                }
            ),
            hide_index=True,
            use_container_width=True,
        )

    display_center_table("CC2", "CC2 (Center 2)")
    st.markdown("---")
    display_center_table("JV", "JV (Juárez)")

    # ------------------------------------------------
    # 7) Supervisor summary tables (CC2 then JV)
    # ------------------------------------------------
    st.markdown("---")
    st.markdown("## 🎯 Resumen por Supervisor (Meta, Ventas, Gap, En Tránsito)")

    def display_supervisor_summary(center_code: str, title: str):
        st.markdown(f"### {title}")

        if df_sup.empty:
            st.info("No hay datos para el resumen por supervisor.")
            return

        dfx = df_sup[df_sup["CentroKey"] == center_code].copy()
        if dfx.empty:
            st.info(f"No hay datos para {title}.")
            return

        dfx = dfx.sort_values(["Gap", "Ventas"], ascending=[False, False])

        meta_sum = float(dfx["Meta Supervisor"].sum())
        ventas_sum = int(dfx["Ventas"].sum())
        monto_sum = float(dfx["Monto_Vendido"].sum())
        gap_sum = meta_sum - ventas_sum
        tr_sum = int(dfx["En_Transito"].sum())

        total_row = {
            "CentroKey": center_code,
            "Supervisor": "TOTAL",
            "Meta Supervisor": meta_sum,
            "Ventas": ventas_sum,
            "Monto_Vendido": monto_sum,
            "ARPU": (monto_sum / ventas_sum) if ventas_sum else 0.0,
            "Gap": gap_sum,
            "En_Transito": tr_sum,
            "Dias Hab Mes": float(dias_hab_total),
            "Dias Hab Restantes": float(dias_hab_restantes),
            "Ventas Diarias Necesarias (Hoy)": (max(0.0, gap_sum) / dias_hab_restantes) if dias_hab_restantes > 0 else 0.0,
        }

        show_cols = [
            "Supervisor",
            "Meta Supervisor",
            "Ventas",
            "Monto_Vendido",
            "ARPU",
            "Gap",
            "En_Transito",
            "Dias Hab Mes",
            "Dias Hab Restantes",
            "Ventas Diarias Necesarias (Hoy)",
        ]

        dfx_show = pd.concat([dfx, pd.DataFrame([total_row])], ignore_index=True)

        st.dataframe(
            dfx_show[show_cols]
            .style
            .apply(_style_sup(show_cols), axis=1)
            .format(
                {
                    "Meta Supervisor": "{:,.0f}",
                    "Ventas": "{:,.0f}",
                    "Monto_Vendido": "${:,.2f}",
                    "ARPU": "${:,.2f}",
                    "Gap": "{:,.2f}",
                    "En_Transito": "{:,.0f}",
                    "Dias Hab Mes": "{:,.1f}",
                    "Dias Hab Restantes": "{:,.1f}",
                    "Ventas Diarias Necesarias (Hoy)": "{:,.2f}",
                }
            ),
            hide_index=True,
            use_container_width=True,
        )

    display_supervisor_summary("CC2", "CC2 (Center 2) — Supervisores")
    st.markdown("---")
    display_supervisor_summary("JV", "JV (Juárez) — Supervisores")

    # ------------------------------------------------
    # 8) PLOTS: Donuts CC2 vs JV
    #    - % Monto Vendido
    #    - % ARPU
    # ------------------------------------------------
    st.markdown("---")

    monto_cc2 = 0.0
    monto_jv = 0.0
    ventas_cc2 = 0
    ventas_jv = 0

    if not df_detalle_full.empty:
        df_cc2_plot = df_detalle_full[df_detalle_full["CentroKey"] == "CC2"].copy()
        df_jv_plot = df_detalle_full[df_detalle_full["CentroKey"] == "JV"].copy()

        monto_cc2 = float(df_cc2_plot["Monto Vendido"].sum()) if "Monto Vendido" in df_cc2_plot.columns else 0.0
        monto_jv = float(df_jv_plot["Monto Vendido"].sum()) if "Monto Vendido" in df_jv_plot.columns else 0.0

        ventas_cc2 = int(df_cc2_plot["Ventas"].sum()) if "Ventas" in df_cc2_plot.columns else 0
        ventas_jv = int(df_jv_plot["Ventas"].sum()) if "Ventas" in df_jv_plot.columns else 0

    arpu_cc2_val = (monto_cc2 / ventas_cc2) if ventas_cc2 > 0 else 0.0
    arpu_jv_val = (monto_jv / ventas_jv) if ventas_jv > 0 else 0.0

    pieL, pieR = st.columns(2, gap="large")

    fig_monto = donut_compare_fig(
        ["CC2", "JV"],
        [monto_cc2, monto_jv],
        "% Monto Vendido",
        fmt_money_short,
    )

    fig_arpu = donut_compare_fig(
        ["CC2", "JV"],
        [arpu_cc2_val, arpu_jv_val],
        "% ARPU",
        fmt_money_short,
    )

    with pieL:
        if fig_monto is None:
            st.info("Sin monto suficiente para graficar % Monto Vendido.")
        else:
            st.plotly_chart(fig_monto, use_container_width=True, key=f"tab5_pie_monto_{mes_sel}_{start_yyyymmdd}_{end_yyyymmdd}")

    with pieR:
        if fig_arpu is None:
            st.info("Sin ARPU suficiente para graficar % ARPU.")
        else:
            st.plotly_chart(fig_arpu, use_container_width=True, key=f"tab5_pie_arpu_{mes_sel}_{start_yyyymmdd}_{end_yyyymmdd}")

    # ------------------------------------------------
    # 9) Download Excel (Detalle + Resumen)
    # ------------------------------------------------
    if not df_detalle_full.empty:
        sheets = {"Detalle Ejecutivos": df_detalle_full.copy()}
        if not df_sup.empty:
            sheets["Resumen Supervisores"] = df_sup.copy()

        st.download_button(
            "⬇️ Descargar TAB 5 (Excel)",
            data=build_excel_bytes(sheets),
            file_name=f"TAB5_Detalle_{mes_sel}_{date.today():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            # ✅ key realmente único (evita colisiones si cambias filtros rápido)
            key=f"dl_tab5_{mes_sel}_{start_yyyymmdd}_{end_yyyymmdd}_{'-'.join(sorted(center_sel or []))}_{normalize_name('|'.join(sorted(sup_sel or [])))}_{normalize_name('|'.join(sorted(ej_sel or [])))}",
        )


# ======================================================
# TAB 6: Región (Vista Universo Completo)
# ======================================================
with tabs[5]:
    st.markdown(f"## Distribución de Ventas por Región — **{mes_labels[mes_sel]}**")
    st.caption("Nota: Se muestran todas las regiones existentes en la base de datos cargada, incluyendo aquellas con 0 ventas este mes.")

    # 1. Obtener el UNIVERSO de regiones (Catalogo)
    # Usamos 'ventas' completo (sin filtros de fecha ni persona) para conocer todas las subregiones posibles
    all_regions = ventas["SUBREGION"].dropna().unique()
    df_universe = pd.DataFrame({"SUBREGION": all_regions})
    
    # Aseguramos limpieza
    df_universe["SUBREGION"] = df_universe["SUBREGION"].astype(str).str.strip().str.upper()
    df_universe = df_universe.drop_duplicates("SUBREGION").sort_values("SUBREGION")

    # 2. Obtener VENTAS del mes seleccionado (Sin filtros de Centro/Supervisor)
    df_month_sales = ventas[ventas["AñoMes"] == mes_sel].copy()
    
    # Agrupamos las ventas reales del mes
    df_month_sales["SUBREGION"] = df_month_sales["SUBREGION"].fillna("SIN REGION").astype(str).str.strip().str.upper()
    sales_counts = df_month_sales.groupby("SUBREGION", as_index=False).size().rename(columns={"size": "Ventas"})

    # 3. Cruzar Universo con Ventas (Left Join)
    # Esto asegura que aparezcan las regiones con 0 ventas
    df_final = df_universe.merge(sales_counts, on="SUBREGION", how="left")
    df_final["Ventas"] = df_final["Ventas"].fillna(0).astype(int)

    # 4. Calcular Porcentajes
    total_sales = df_final["Ventas"].sum()
    df_final["%"] = np.where(total_sales > 0, df_final["Ventas"] / total_sales, 0.0)

    # Ordenar: primero las que tienen ventas (descendente), luego alfabético
    df_final = df_final.sort_values(["Ventas", "SUBREGION"], ascending=[False, True])

    # 5. Visualización
    if df_final.empty:
        st.warning("No se encontraron regiones en la base de datos.")
    else:
        c1, c2 = st.columns([0.65, 0.35], gap="large")

        with c1:
            # Gráfico de Dona
            # Filtramos solo las que tienen ventas > 0 para el gráfico (para que no se vea saturado de ceros),
            # pero la tabla de al lado sí mostrará todo.
            df_plot = df_final[df_final["Ventas"] > 0].copy()
            
            if df_plot.empty:
                st.info("No hay ventas registradas en ninguna región este mes.")
            else:
                fig = px.pie(
                    df_plot, 
                    names="SUBREGION", 
                    values="Ventas", 
                    title="Participación (Regiones con Venta)", 
                    hole=0.45, 
                    template=PLOTLY_TEMPLATE
                )
                fig.update_traces(textposition='inside', textinfo='percent+label')
                fig.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", 
                    plot_bgcolor="rgba(0,0,0,0)",
                    height=450,
                    margin=dict(t=40, b=20, l=20, r=20),
                    showlegend=False 
                )
                st.plotly_chart(fig, width="stretch", key="t5_region_pie_universe")

        with c2:
            st.markdown("#### Detalle Global")
            
            # Agregamos fila de TOTAL
            reg_table_show = add_totals_row(
                df_final, 
                label_col="SUBREGION", 
                totals={"Ventas": total_sales, "%": 1.0},
                label="TOTAL"
            )

            # Mostramos la tabla completa (incluyendo ceros)
            st.dataframe(
                style_totals_bold(reg_table_show, label_col="SUBREGION")
                .format({
                    "Ventas": "{:,.0f}", 
                    "%": "{:.1%}"
                }),
                hide_index=True,
                use_container_width=True,
                height=450
            )

        # Botón de descarga
        st.download_button(
            "⬇️ Descargar Todas las Regiones",
            data=build_excel_bytes({"Todas_Regiones": reg_table_show}),
            file_name=f"Regiones_Universo_{mes_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_region_universe"
        )

# ======================================================
# TAB 7: TOPS
# ======================================================
with tabs[6]:
    st.markdown(f"## TOP Ventas — **{mes_labels[mes_sel]}**")

    df_m_all = filter_month(ventas, mes_sel)

    st.markdown("### TOP Ventas x Ejecutivo (por Centro)")
    centro_top = st.selectbox("Centro para TOP Ejecutivo", options=["JV", "CC2"], index=0, key="top_exec_center")
    df_top = df_m_all[df_m_all["CentroKey"] == centro_top].copy()

    top_ej = (
        df_top.groupby("EJECUTIVO", as_index=False)
        .size()
        .rename(columns={"size": "VENTAS"})
        .sort_values("VENTAS", ascending=False)
        .head(8)
    )

    fig1 = go.Figure()
    fig1.add_trace(
        go.Bar(
            x=top_ej["VENTAS"],
            y=top_ej["EJECUTIVO"],
            orientation="h",
            text=top_ej["VENTAS"],
            textposition="outside",
        )
    )
    fig1.update_layout(
        title="TOP EJECUTIVOS",
        xaxis_title="VENTAS",
        yaxis_title="EJECUTIVO",
        height=420,
        yaxis=dict(categoryorder="total ascending"),
    )
    apply_plotly_theme(fig1)
    st.plotly_chart(fig1, width="stretch", key=f"t6_top_ej_{centro_top}_{mes_sel}")

    st.markdown("---")

    st.markdown("### TOP Ventas Globales — (Color por Centro)")
    df_g = df_m_all.copy()

    g = (
        df_g.groupby(["EJECUTIVO", "CentroKey"], as_index=False)
        .size()
        .rename(columns={"size": "VENTAS"})
    )
    total_exec = g.groupby("EJECUTIVO", as_index=False)["VENTAS"].sum().rename(columns={"VENTAS": "VENTAS_TOTAL"})
    dominant = g.sort_values("VENTAS", ascending=False).drop_duplicates("EJECUTIVO")
    dom = dominant.merge(total_exec, on="EJECUTIVO", how="left")
    top_global = dom.sort_values("VENTAS_TOTAL", ascending=False).head(10)

    fig2 = px.bar(
        top_global,
        x="VENTAS_TOTAL",
        y="EJECUTIVO",
        orientation="h",
        text="VENTAS_TOTAL",
        color="CentroKey",
        title="TOP EJECUTIVOS GLOBAL",
        template=PLOTLY_TEMPLATE,
    )
    fig2.update_layout(
        height=460,
        yaxis=dict(categoryorder="total ascending"),
        xaxis_title="VENTAS",
        legend_title_text="CENTRO",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    fig2.update_traces(textposition="inside", insidetextanchor="end")
    st.plotly_chart(fig2, width="stretch", key=f"t6_top_global_{mes_sel}")

    st.markdown("---")
    st.markdown("### TOP Ventas por Equipo (cada equipo = Supervisor) — por Centro")

    centro_equipo = st.selectbox(
        "Centro para TOP por Equipo",
        options=["CC2", "JV"],
        index=0,
        key="top_equipo_centro",
    )

    df_eq = df_m_all[df_m_all["CentroKey"] == centro_equipo].copy()

    if df_eq.empty:
        st.info(f"Sin datos para {centro_equipo} en el mes seleccionado.")
    else:
        df_eq_sup = df_eq.copy()
        df_eq_sup_nb = df_eq_sup[df_eq_sup["Supervisor"].astype(str).str.upper() != "BAJA"].copy()
        if not df_eq_sup_nb.empty:
            df_eq_sup = df_eq_sup_nb

        top_sup = (
            df_eq_sup.groupby("Supervisor", as_index=False)
            .size()
            .rename(columns={"size": "VENTAS"})
            .sort_values("VENTAS", ascending=False)
            .head(6)
        )

        fig_sup = go.Figure()
        fig_sup.add_trace(
            go.Bar(
                x=top_sup["VENTAS"],
                y=top_sup["Supervisor"],
                orientation="h",
                text=top_sup["VENTAS"],
                textposition="outside",
            )
        )
        fig_sup.update_layout(
            title="TOP SUPERVISORES",
            xaxis_title="VENTAS",
            yaxis_title="SUPERVISOR",
            height=330,
            yaxis=dict(categoryorder="total ascending"),
        )
        apply_plotly_theme(fig_sup)
        st.plotly_chart(fig_sup, width="stretch", key=f"t6_top_sup_{centro_equipo}_{mes_sel}")

        equipos_n = 4 if centro_equipo == "CC2" else 2

        equipos_top = (
            df_eq_sup.groupby("Supervisor", as_index=False)
            .size()
            .rename(columns={"size": "VENTAS"})
            .sort_values("VENTAS", ascending=False)
        )

        equipos = equipos_top["Supervisor"].head(min(equipos_n, len(equipos_top))).tolist()

        if not equipos:
            st.info("No hay equipos para mostrar con los filtros actuales.")
        else:
            grid = st.columns(2, gap="large")
            for i, sup in enumerate(equipos):
                df_team = df_eq[df_eq["Supervisor"] == sup].copy()

                top_exec_team = (
                    df_team.groupby("EJECUTIVO", as_index=False)
                    .size()
                    .rename(columns={"size": "VENTAS"})
                    .sort_values("VENTAS", ascending=False)
                    .head(6)
                )

                fig_team = go.Figure()
                fig_team.add_trace(
                    go.Bar(
                        x=top_exec_team["EJECUTIVO"],
                        y=top_exec_team["VENTAS"],
                        text=top_exec_team["VENTAS"],
                        textposition="outside",
                    )
                )
                fig_team.update_layout(
                    title=f"EQUIPO: {sup}",
                    xaxis_title="",
                    yaxis_title="VENTAS",
                    height=320,
                )
                apply_plotly_theme(fig_team)
                fig_team.update_layout(margin=dict(l=20, r=20, t=60, b=100))
                fig_team.update_xaxes(tickangle=-35)

                with grid[i % 2]:
                    st.plotly_chart(fig_team, width="stretch", key=f"top_team_{centro_equipo}_{normalize_name(sup)}_{i}")

    st.caption(f"Día: {datetime.now().strftime('%d/%m/%Y')}")

# ======================================================
# TAB 8: METAS
# ======================================================
with tabs[7]:
    st.markdown(f"## Metas — **{mes_labels[mes_sel]}**")

    df_mes = filter_month(ventas, mes_sel)

    if center_sel:
        df_mes = df_mes[df_mes["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_mes = df_mes[df_mes["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_mes = df_mes[df_mes["EJECUTIVO"].isin(ej_sel)]
    if sub_sel:
        df_mes = df_mes[df_mes["SUBREGION"].isin(sub_sel)]

    metas_centro = metas[metas["Nivel"].str.lower() == "centro"].copy()
    metas_sup = metas[metas["Nivel"].str.lower() == "supervisor"].copy()

    metas_centro_f = metas_centro.copy()
    if center_sel:
        metas_centro_f = metas_centro_f[metas_centro_f["Centro"].isin([c.upper() for c in center_sel])]

    # ✅ Supervisores ACTIVOS REALES (no jefes) = empleados activos cuyo Puesto contiene SUPERV
    emp_sup = empleados.copy()
    emp_sup["Estatus"] = emp_sup["Estatus"].astype(str).str.strip().str.upper()
    emp_sup["Puesto"] = emp_sup["Puesto"].astype(str).str.strip().str.upper()
    emp_sup["Nombre"] = emp_sup["Nombre"].astype(str).str.strip()
    emp_sup["Centro"] = emp_sup["Centro"].astype(str).str.strip()

    emp_sup = emp_sup[
        (emp_sup["Estatus"] == "ACTIVO")
        & (emp_sup["Puesto"].str.contains("SUPERV", na=False))
        & (emp_sup["Centro"].str.upper().str.contains("EXP ATT C CENTER", na=False))
    ].copy()

    emp_sup["CentroKey"] = np.where(
        emp_sup["Centro"].str.upper().str.contains("JUAREZ", na=False),
        "JV",
        "CC2",
    )
    
    emp_sup["Supervisor"] = emp_sup["Nombre"]
    emp_sup["Supervisor_norm"] = emp_sup["Supervisor"].apply(normalize_name)

    _YARELI_NORM = normalize_name("YARELI SILVA ZEFERINO")
    emp_sup = emp_sup[emp_sup["Supervisor_norm"] != _YARELI_NORM].copy()

    # ❌ Hide any "JULIO ..."
    emp_sup["FirstName"] = emp_sup["Supervisor_norm"].astype(str).str.split().str[0]
    emp_sup = emp_sup[emp_sup["FirstName"] != "JULIO"].copy()
    emp_sup.drop(columns=["FirstName"], inplace=True, errors="ignore")

    # Respect sidebar supervisor filter (if exists)
    sup_norm_filter = {normalize_name(s) for s in sup_sel} if sup_sel else None
    if sup_norm_filter:
        emp_sup = emp_sup[emp_sup["Supervisor_norm"].isin(sup_norm_filter)].copy()

    # Dedup
    # Dedup
    emp_sup = emp_sup.drop_duplicates("Supervisor_norm", keep="first")

    active_supervisores_norm = set(emp_sup["Supervisor_norm"].dropna().tolist())

    # ==========================================================
    # ✅ Mostrar a Maria Luisa en Metas siempre que venga en el Excel del mes
    # ==========================================================
    ml_norm_meta = normalize_name("MARIA LUISA MEZA GOEL")
    ml_has_meta_this_month = (
        not metas_sup.empty
        and ml_norm_meta in set(metas_sup["Nombre_norm"].astype(str))
    )

    if ml_has_meta_this_month and ml_norm_meta not in active_supervisores_norm:
        nueva_fila = pd.DataFrame([{
            "Supervisor": "MARIA LUISA MEZA GOEL",
            "Supervisor_norm": ml_norm_meta,
            "CentroKey": "JV"
        }])
        emp_sup = pd.concat([emp_sup, nueva_fila], ignore_index=True)
        active_supervisores_norm.add(ml_norm_meta)
    # ==========================================================

    metas_sup_show = metas_sup[metas_sup["Nombre_norm"].isin(active_supervisores_norm)].copy()


    if center_sel:
        metas_sup_show["Centro"] = metas_sup_show["Centro"].astype(str).str.strip().str.upper()
        metas_sup_show = metas_sup_show[metas_sup_show["Centro"].isin([c.upper() for c in center_sel])].copy()

    st.markdown("### Metas Globales")
    achieved_global = int(len(df_mes))

    if not metas_sup_show.empty:
        meta_global = int(pd.to_numeric(metas_sup_show["Meta"], errors="coerce").fillna(0).sum())

        # --- INICIO DEL CAMBIO ---
        # Rescatar la meta perdida de Maria Fernanda y sumarla a la Meta Global
        ml_norm_meta = normalize_name("MARIA LUISA MEZA GOEL")
        mf_norm_meta = normalize_name("MARIA FERNANDA MARTINEZ BISTRAIN")

        if ml_norm_meta in active_supervisores_norm:
            # Buscamos cuánto era la meta de Maria Fernanda en el Excel original
            mf_row = metas_sup[metas_sup["Nombre_norm"] == mf_norm_meta]
            if not mf_row.empty:
                extra_meta = int(pd.to_numeric(mf_row["Meta"].iloc[0], errors="coerce") or 0)
                # Si Maria Fernanda está inactiva y no se sumó, la sumamos a la fuerza
                if mf_norm_meta not in active_supervisores_norm:
                    meta_global += extra_meta
        # --- FIN DEL CAMBIO ---

    else:
        meta_global = int(pd.to_numeric(metas_centro_f["Meta"], errors="coerce").fillna(0).sum()) if not metas_centro_f.empty else 0

    faltan_global = meta_global - achieved_global

    colg1, colg2 = st.columns([0.72, 0.28], gap="large")
    with colg1:
        fig = gauge_fig(achieved_global, meta_global, "VISOR GLOBAL")
        st.plotly_chart(fig, width="stretch", key="gauge_global")
    with colg2:
        metric_card("Meta Global", fmt_int(meta_global))
        metric_card("Alcanzado", fmt_int(achieved_global))
        metric_card("FALTAN", fmt_int(faltan_global))

    st.markdown("---")

    st.markdown("### Metas x Centro")
    if metas_centro_f.empty:
        st.info("No hay metas de Centro (con los filtros actuales).")
    else:
        ach_center = df_mes.groupby("CentroKey").size().to_dict()
        cols = st.columns(2, gap="large")

        for i, centro in enumerate(["CC2", "JV"]):
            row = metas_centro_f[metas_centro_f["Centro"] == centro].head(1)
            if row.empty:
                continue
            meta_val = int(pd.to_numeric(row["Meta"].iloc[0], errors="coerce") or 0)
            coord = str(row["Nombre"].iloc[0])
            achieved = int(ach_center.get(centro, 0))
            faltan = meta_val - achieved

            with cols[i]:
                fig = gauge_fig(achieved, meta_val, "VISOR DE METAS X COORDINADOR")
                st.plotly_chart(fig, width="stretch", key=f"gauge_centro_{centro}")
                st.markdown(f"**COORDINADOR:** {coord}")
                st.markdown(f"**FALTAN:** {fmt_int(faltan)}")

    st.markdown("---")

    # Metas x Supervisor — SOLO ACTIVOS REALES
    st.markdown("### Metas x Supervisor — (Solo Activos, comparativo CC2 vs JV)")

    centers_to_show = center_sel if center_sel else ["CC2", "JV"]

    show_cc2 = "CC2" in centers_to_show
    show_jv = "JV" in centers_to_show

    if show_cc2 and show_jv:
        col_cc2, col_jv = st.columns(2, gap="large")
        center_slots = {"CC2": col_cc2, "JV": col_jv}
        inner_grid_cols = 1
    else:
        center_slots = {"CC2": st.container(), "JV": st.container()}
        inner_grid_cols = 2

    metas_sup_local = metas_sup.copy()
    metas_sup_local["Centro"] = metas_sup_local["Centro"].astype(str).str.strip().str.upper()

    for centro in ["CC2", "JV"]:
        if centro not in centers_to_show:
            continue

        with center_slots[centro]:
            st.markdown(f"#### {'CC2 (Center 2)' if centro=='CC2' else 'JV (Juárez)'}")

            cand = emp_sup[emp_sup["CentroKey"] == centro][["Supervisor", "Supervisor_norm"]].copy()
            if cand.empty:
                st.info("No hay supervisores ACTIVOS (puesto supervisor) para este centro con los filtros actuales.")
                continue

            df_c = df_mes[df_mes["CentroKey"] == centro].copy()
            ach_map = df_c.groupby("Supervisor_norm").size().to_dict() if not df_c.empty else {}

            cand["Achieved"] = cand["Supervisor_norm"].map(lambda n: int(ach_map.get(n, 0)))
            cand = cand.sort_values("Achieved", ascending=False).reset_index(drop=True)

            grid_cols = st.columns(inner_grid_cols, gap="large")

            for idx, rr in cand.iterrows():
                sup_name = str(rr["Supervisor"]).strip()
                sup_norm = str(rr["Supervisor_norm"]).strip()
                achieved = int(rr["Achieved"])

                meta_val = np.nan

                # ✅ Primero buscar la meta real del supervisor en el Excel del mes.
                # ✅ Solo si no existe y es Maria Luisa, usar Maria Fernanda como fallback.
                _lookup_norm = sup_norm

                mrow = metas_sup_local[
                    (metas_sup_local["Nombre_norm"] == _lookup_norm) & (metas_sup_local["Centro"] == centro)
                ]

                if mrow.empty and _lookup_norm == normalize_name("MARIA LUISA MEZA GOEL"):
                    mrow = metas_sup_local[
                        (metas_sup_local["Nombre_norm"] == normalize_name("MARIA FERNANDA MARTINEZ BISTRAIN"))
                        & (metas_sup_local["Centro"] == centro)
                    ]

                if not mrow.empty:
                    mv = pd.to_numeric(mrow["Meta"].iloc[0], errors="coerce")
                    meta_val = float(mv) if pd.notna(mv) else np.nan

                faltan = (meta_val - achieved) if pd.notna(meta_val) else np.nan

                with grid_cols[idx % inner_grid_cols]:
                    fig = gauge_fig(
                        achieved,
                        meta_val if pd.notna(meta_val) else 0,
                        f"SUPERVISOR: {sup_name}",
                    )
                    st.plotly_chart(fig, width="stretch", key=f"meta_sup_{centro}_{sup_norm}_{idx}")

                    st.markdown(f"**Meta:** {fmt_int(meta_val) if pd.notna(meta_val) else '-'}")
                    st.markdown(f"**Alcanzado:** {fmt_int(achieved)}")
                    st.markdown(f"**FALTAN:** {fmt_int(faltan) if pd.notna(faltan) else '-'}")

    st.caption(f"Día: {datetime.now().strftime('%d/%m/%Y')}")

# ======================================================
# TAB 9: Tendencia x Ejecutivo
#   ✅ UPDATED: Sanity check now follows "Filtro por meses y semanas"
#              - If multiple months selected, Meta = sum(meta per month)
#              - Ventas = ventas inside the selected months+weeks interval
#              - Meta per month is computed using the 1st day of EACH month
#              - Antigüedad is calendar days from ingreso to 1st day of month (NOT workable days)
#
#   ✅ FIXED WARNINGS (Pylance):
#      - "ventas" is not defined
#      - "filter_month" is not defined
#      (Only for type-checking; does NOT change runtime behavior)
#
#   ✅ IMPLEMENTED:
#      - Interval team plot now colors bars: GREEN=HEALTHY, RED=RISKY
# ======================================================
with tabs[8]:
    st.markdown("## Tendencia x Ejecutivo")

    # -------------------------------
    # ✅ Pylance warnings fix (no runtime impact)
    # -------------------------------
    from typing import TYPE_CHECKING
    if TYPE_CHECKING:
        ventas: pd.DataFrame
        filter_month: str
        tabs: list

    # 1) Base context (same global filters you already apply in this tab)
    df_ctx = ventas.copy()
    if center_sel:
        df_ctx = df_ctx[df_ctx["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_ctx = df_ctx[df_ctx["Supervisor"].isin(sup_sel)]
    if sub_sel:
        df_ctx = df_ctx[df_ctx["SUBREGION"].isin(sub_sel)]

    # ======================================================
    # ✅ Filtro por meses y semanas (controla TODO este TAB)
    # ======================================================
    st.markdown("### Filtro por meses y semanas (solo esta pestaña)")

    df_mvw_ctx = df_ctx.copy()
    df_mvw_ctx = df_mvw_ctx[df_mvw_ctx["FECHA DE CAPTURA"].notna()].copy()
    df_mvw_ctx["T_DT"] = pd.to_datetime(df_mvw_ctx["FECHA DE CAPTURA"], errors="coerce")
    df_mvw_ctx = df_mvw_ctx[df_mvw_ctx["T_DT"].notna()].copy()

    if df_mvw_ctx.empty:
        st.info("No hay datos con fecha válida para aplicar el filtro de meses/semanas en esta pestaña.")
        df_ctx = df_ctx.iloc[0:0].copy()
    else:
        df_mvw_ctx["T_MonthKey"] = df_mvw_ctx["T_DT"].dt.strftime("%Y-%m")
        df_mvw_ctx["T_MonthName"] = df_mvw_ctx["T_DT"].dt.strftime("%B")
        df_mvw_ctx["T_MonthLabel"] = df_mvw_ctx["T_MonthKey"] + " (" + df_mvw_ctx["T_MonthName"] + ")"

        # Semana del mes
        month_start = df_mvw_ctx["T_DT"].dt.to_period("M").dt.to_timestamp()
        first_wd = month_start.dt.weekday
        df_mvw_ctx["T_WeekOfMonth"] = ((df_mvw_ctx["T_DT"].dt.day + first_wd - 1) // 7) + 1
        df_mvw_ctx["T_WeekLabel"] = (
            df_mvw_ctx["T_MonthLabel"] + " - Semana " + df_mvw_ctx["T_WeekOfMonth"].astype(int).astype(str)
        )

        month_map = (
            df_mvw_ctx[["T_MonthKey", "T_MonthLabel"]]
            .dropna()
            .drop_duplicates()
            .sort_values("T_MonthKey")
        )
        m_options = month_map["T_MonthLabel"].tolist()

        # Default: mes_sel + mes anterior (si existe)
        def _ym_to_monthkey_from_ym(ym_int: int) -> str:
            y = ym_int // 100
            m = ym_int % 100
            return f"{y}-{m:02d}"

        defaults = []
        try:
            cur_key = _ym_to_monthkey_from_ym(int(mes_sel)) if mes_sel else None
        except Exception:
            cur_key = None

        keys_sorted = month_map["T_MonthKey"].tolist()
        if cur_key and cur_key in keys_sorted:
            idx = keys_sorted.index(cur_key)
            defaults_keys = [keys_sorted[idx]]
            if idx - 1 >= 0:
                defaults_keys.insert(0, keys_sorted[idx - 1])
            defaults = month_map[month_map["T_MonthKey"].isin(defaults_keys)]["T_MonthLabel"].tolist()

        if not defaults:
            defaults = m_options[-2:] if len(m_options) >= 2 else m_options

        m_sel = st.multiselect(
            "Selecciona uno o más meses (Tendencia Ejecutivo)",
            options=m_options,
            default=defaults,
            key="tend_mvw_months_multi",
        )

        df_f = df_mvw_ctx.copy()
        if m_sel:
            df_f = df_f[df_f["T_MonthLabel"].isin(m_sel)].copy()

        # Semanas disponibles según meses elegidos
        w_map = (
            df_f[["T_MonthKey", "T_MonthLabel", "T_WeekOfMonth", "T_WeekLabel"]]
            .dropna()
            .drop_duplicates()
            .sort_values(["T_MonthKey", "T_WeekOfMonth"])
        )
        w_options = w_map["T_WeekLabel"].tolist()

        # ✅ NEW: When months selection changes, auto-select ALL weeks in that interval
        _months_key = tuple(sorted(m_sel)) if m_sel else tuple()
        _prev_months_key = st.session_state.get("_tend_prev_months_key", None)
        months_changed = (_prev_months_key != _months_key)
        st.session_state["_tend_prev_months_key"] = _months_key

        # ✅ Sanitize previous week selections to avoid "value not in options" errors
        prev_weeks = st.session_state.get("tend_mvw_weeks_multi", None)
        if months_changed or prev_weeks is None:
            st.session_state["tend_mvw_weeks_multi"] = w_options.copy()
        else:
            st.session_state["tend_mvw_weeks_multi"] = [w for w in prev_weeks if w in w_options]

        w_sel = st.multiselect(
            "Selecciona Semana(s) del mes (Tendencia Ejecutivo)",
            options=w_options,
            default=w_options,
            key="tend_mvw_weeks_multi",
        )


        if w_sel:
            df_f = df_f[df_f["T_WeekLabel"].isin(w_sel)].copy()

        df_ctx = df_f.copy()

    st.markdown("---")

    # ✅ Exclusion ONLY inside this TAB (options + visuals)
    df_ctx_opts = df_ctx[~df_ctx["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS)].copy()

    # ======================================================
    # ✅ Teams plot: Ventas por Supervisor + etiqueta #Ejecutivos
    # ======================================================
    st.markdown("### Equipos (Supervisores) — Ejecutivos por equipo y Ventas")

    df_team = df_ctx_opts.copy()

    if df_team.empty:
        st.info("No hay datos para visualizar equipos con los filtros actuales.")
    else:
        team_kpi = (
            df_team.groupby("Supervisor", as_index=False)
            .agg(
                Ventas=("FOLIO", "count"),
                Ejecutivos=("EJECUTIVO", "nunique"),
            )
            .sort_values("Ventas", ascending=False)
            .reset_index(drop=True)
        )

        team_kpi["Etiqueta"] = team_kpi.apply(
            lambda r: f"{int(r['Ventas']):,} ventas  |  {int(r['Ejecutivos'])} ejecutivos",
            axis=1,
        )

        dyn_h = max(380, 160 + 52 * len(team_kpi))

        fig_team = px.bar(
            team_kpi.sort_values("Ventas", ascending=True),
            x="Ventas",
            y="Supervisor",
            orientation="h",
            text="Etiqueta",
            title="Ventas por Equipo (Supervisor) y número de Ejecutivos",
            labels={"Supervisor": "Supervisor (Equipo)", "Ventas": "Ventas"},
            template=PLOTLY_TEMPLATE,
        )
        fig_team.update_traces(textposition="outside", cliponaxis=False)
        fig_team.update_layout(
            height=dyn_h,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=320, r=40, t=70, b=30),
            showlegend=False,
            xaxis=dict(zeroline=False),
            yaxis=dict(automargin=True),
        )

        st.plotly_chart(fig_team, width="stretch", key="t8_team_supervisores_overview")

        team_show = team_kpi[["Supervisor", "Ventas", "Ejecutivos"]].copy()
        team_show = add_totals_row(
            team_show,
            label_col="Supervisor",
            totals={"Ventas": int(team_show["Ventas"].sum()), "Ejecutivos": int(team_show["Ejecutivos"].sum())},
            label="TOTAL",
        )
        st.dataframe(
            style_totals_bold(team_show, label_col="Supervisor").format({"Ventas": "{:,.0f}", "Ejecutivos": "{:,.0f}"}),
            hide_index=True,
            width="stretch",
        )

    st.markdown("---")

    # ======================================================
    # 2) Supervisor filter (TAB)
    # ======================================================
    sup_opts_tab = sorted([s for s in df_ctx_opts["Supervisor"].dropna().unique().tolist()])
    sup_options = ["Todos"] + sup_opts_tab

    if "tend_ej_sup_tab" in st.session_state and st.session_state["tend_ej_sup_tab"] not in sup_options:
        st.session_state["tend_ej_sup_tab"] = "Todos"

    sup_tab_choice = st.selectbox(
        "Supervisor (Tendencia Ejecutivo)",
        options=sup_options,
        index=0,
        key="tend_ej_sup_tab",
    )

    df_ctx2 = df_ctx_opts.copy()
    if sup_tab_choice != "Todos":
        df_ctx2 = df_ctx2[df_ctx2["Supervisor"] == sup_tab_choice].copy()

    # ======================================================
    # 3) Ejecutivo filter (TAB)
    # ======================================================
    ej_opts = sorted([e for e in df_ctx2["EJECUTIVO"].dropna().unique().tolist()])
    if not ej_opts:
        st.info("No hay ejecutivos disponibles para el supervisor seleccionado con los filtros actuales.")
    else:
        if "tend_ej_sel" in st.session_state and st.session_state["tend_ej_sel"] not in ej_opts:
            st.session_state["tend_ej_sel"] = ej_opts[0]

        ej = st.selectbox("Ejecutivo", options=ej_opts, key="tend_ej_sel")
        st.markdown(f"✅ Has seleccionado: **{ej}**")

        # 4) Data for charts
        df_e = df_ctx2[df_ctx2["EJECUTIVO"] == ej].copy()

        if df_e.empty:
            st.info("Sin datos para el ejecutivo seleccionado con los filtros actuales.")
        else:
            m = (
                df_e.groupby("AñoMes", as_index=False)
                .size()
                .rename(columns={"size": "Ventas"})
                .sort_values("AñoMes")
                .reset_index(drop=True)
            )

            m["MesDT"] = pd.to_datetime(m["AñoMes"].astype(int).astype(str) + "01", format="%Y%m%d", errors="coerce")

            prom = float(m.iloc[:-1]["Ventas"].mean()) if len(m) > 1 else np.nan

            cur_ym = int(m["AñoMes"].iloc[-1])
            cur_v = int(m["Ventas"].iloc[-1])
            cur_label = mes_labels.get(cur_ym, month_key_to_name_es(cur_ym))

            c1, c2 = st.columns([0.65, 0.35], gap="large")

            with c1:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=m["MesDT"], y=m["Ventas"], mode="lines+markers", name="Ventas"))
                if not np.isnan(prom):
                    fig.add_trace(go.Scatter(x=m["MesDT"], y=[prom] * len(m), mode="lines", name="Promedio (sin mes actual)"))
                fig.update_layout(
                    title="Tendencia X Ejecutivo",
                    xaxis_title="Mes",
                    yaxis_title="Ventas",
                    height=420,
                )
                fig.update_xaxes(tickformat="%b %Y")
                apply_plotly_theme(fig)
                st.plotly_chart(fig, width="stretch", key=f"t8_tend_line_{normalize_name(ej)}")

            with c2:
                metric_card("Ventas Mes", fmt_int(cur_v), sub=cur_label)
                if not np.isnan(prom):
                    dif = cur_v - prom
                    metric_card("Diferencia", f"{dif:+.0f} vs promedio")

                figb = px.bar(
                    pd.DataFrame({"Mes": [cur_label], "Ventas": [cur_v]}),
                    x="Mes",
                    y="Ventas",
                    title="Ventas X Ejecutivo",
                    template=PLOTLY_TEMPLATE,
                )
                figb.update_layout(
                    height=300,
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=20, r=20, t=60, b=40),
                )
                st.plotly_chart(figb, width="stretch", key=f"t8_tend_bar_{normalize_name(ej)}")

            # ======================================================
            # ✅ SANITY CHECK (INTERVAL) + DESGLOSE POR EJECUTIVO (EQUIPO)
            # ======================================================
            st.markdown("---")
            st.markdown("### ✅ Sanity check — Meta mensual por antigüedad (Healthy vs Risky)")

            DEFAULT_META = 12  # default

            # ---------- Robust column picking ----------
            def _norm_txt(x: str) -> str:
                try:
                    x = str(x)
                except Exception:
                    x = ""
                x = x.strip().lower()
                x = unicodedata.normalize("NFKD", x)
                x = "".join([c for c in x if not unicodedata.combining(c)])
                x = " ".join(x.split())
                return x

            def _pick_col(df: pd.DataFrame, candidates):
                cols = list(df.columns)
                norm_cols = {_norm_txt(c): c for c in cols}

                for cand in candidates:
                    k = _norm_txt(cand)
                    if k in norm_cols:
                        return norm_cols[k]

                for cand in candidates:
                    kc = _norm_txt(cand)
                    for nk, real in norm_cols.items():
                        if kc and kc in nk:
                            return real
                return None

            name_col = _pick_col(
                empleados,
                ["Nombre Completo", "NOMBRE COMPLETO", "Nombre", "NOMBRE", "Empleado", "EMPLEADO", "Ejecutivo", "EJECUTIVO"],
            )
            ing_col = _pick_col(
                empleados,
                ["Fecha Ingreso", "FechaIngreso", "FECHA_INGRESO", "Fecha Alta", "FechaAlta", "FECHA_ALTA"],
            )
            dias_col = _pick_col(
                empleados,
                ["Dias Activos", "Días Activos", "Dias Activo", "Días Activo", "DiasActivos", "DIAS_ACTIVOS", "Dias", "DIAS"],
            )

            # ---------- Build employee records + better name matching ----------
            emp_by_norm = {}
            emp_records = []

            _STOP = {"DE", "DEL", "LA", "LAS", "LOS", "Y"}

            def _tokens(nombre: str):
                n = normalize_name(str(nombre).strip())
                toks = [t for t in n.split() if t]
                toks = [t for t in toks if t not in _STOP]
                return set(toks)

            if name_col:
                tmp = empleados[[name_col] + ([ing_col] if ing_col else []) + ([dias_col] if dias_col else [])].copy()
                tmp[name_col] = tmp[name_col].astype(str).str.strip()
                tmp["Nombre_norm"] = tmp[name_col].apply(normalize_name)

                if ing_col:
                    tmp["IngresoDT"] = pd.to_datetime(tmp[ing_col], errors="coerce")
                else:
                    tmp["IngresoDT"] = pd.NaT

                if dias_col:
                    tmp["DiasActivos"] = pd.to_numeric(tmp[dias_col], errors="coerce")
                else:
                    tmp["DiasActivos"] = np.nan

                tmp["_has_dias"] = tmp["DiasActivos"].notna().astype(int)
                tmp = tmp.sort_values(["_has_dias", "IngresoDT"], ascending=[False, True])
                tmp = tmp.drop_duplicates("Nombre_norm", keep="first").drop(columns=["_has_dias"])

                for _, r in tmp.iterrows():
                    nn = r["Nombre_norm"]
                    rec = {
                        "norm": nn,
                        "tokens": _tokens(nn),
                        "ingreso": (pd.Timestamp(r["IngresoDT"]).normalize() if pd.notna(r["IngresoDT"]) else None),
                        "dias": (float(r["DiasActivos"]) if pd.notna(r["DiasActivos"]) else None),
                    }
                    emp_by_norm[nn] = rec
                    emp_records.append(rec)

            def resolve_emp_record(nombre_ej: str):
                n = normalize_name(str(nombre_ej).strip())
                if n in emp_by_norm:
                    return emp_by_norm[n]

                t = _tokens(nombre_ej)
                if not t or not emp_records:
                    return None

                best = None
                best_score = 0.0
                for rec in emp_records:
                    inter = len(t & rec["tokens"])
                    if inter == 0:
                        continue
                    score = inter / max(len(t), len(rec["tokens"]))
                    if score > best_score:
                        best_score = score
                        best = rec

                if best and (best_score >= 0.70 or len(t & best["tokens"]) >= 3):
                    return best
                return None

            # ---------- Calendar days from ingreso -> ref_end ----------
            def calendar_days_since(ingreso_dt: pd.Timestamp, ref_end: pd.Timestamp) -> int:
                if ingreso_dt is None or pd.isna(ingreso_dt):
                    return 0
                start = pd.Timestamp(ingreso_dt).normalize()
                end = pd.Timestamp(ref_end).normalize()
                d = int((end - start).days)
                return max(d, 0)

            def antiguedad_meses_y_dias(nombre_ej: str, ref_end: pd.Timestamp):
                rec = resolve_emp_record(nombre_ej)

                if rec and rec.get("ingreso") is not None:
                    cd = calendar_days_since(rec["ingreso"], ref_end)
                    return int(cd // 30), int(cd % 30), rec.get("ingreso")

                if rec and rec.get("dias") is not None and not pd.isna(rec["dias"]):
                    d_int = int(float(rec["dias"]))
                    return int(d_int // 30), int(d_int % 30), rec.get("ingreso")

                return None, None, None

            # ✅ meta evaluated VS the 1st day of each month
            # ✅ rule: if ingreso > 1st day of month => meta = 6
            def meta_por_antiguedad(nombre_ej: str, ref_end: pd.Timestamp) -> int:
                rec = resolve_emp_record(nombre_ej)

                if rec and rec.get("ingreso") is not None:
                    ing = pd.Timestamp(rec["ingreso"]).normalize()
                    ref0 = pd.Timestamp(ref_end).normalize()

                    if ing > ref0:
                        return 6

                    cd = calendar_days_since(ing, ref0)
                    return 12 if cd > 41 else 6

                if rec and rec.get("dias") is not None and not pd.isna(rec["dias"]):
                    return 12 if float(rec["dias"]) > 41 else 6

                return int(DEFAULT_META)

            # ======================================================
            # ✅ Interval scope (controlled by months+weeks filter)
            # ======================================================
            df_scope = df_ctx2.copy()  # already filtered by month(s)+week(s) + supervisor choice
            months_in_scope = sorted(df_scope["T_MonthKey"].dropna().unique().tolist())

            if df_scope.empty or not months_in_scope:
                st.info("No hay datos suficientes en el intervalo seleccionado (meses/semanas) para el sanity check.")
            else:
                df_e_scope = df_scope[df_scope["EJECUTIVO"] == ej].copy()
                ventas_intervalo = int(df_e_scope.shape[0])

                meta_rows = []
                meta_total = 0

                for mk in months_in_scope:
                    try:
                        yy, mm = mk.split("-")
                        ms = pd.Timestamp(year=int(yy), month=int(mm), day=1).normalize()
                    except Exception:
                        continue

                    ventas_m = int(df_e_scope[df_e_scope["T_MonthKey"] == mk].shape[0])
                    meta_m = int(meta_por_antiguedad(ej, ms))  # ✅ per-month meta at 1st day
                    meta_rows.append({"Mes": mk, "Ventas": ventas_m, "Meta": meta_m, "Delta": ventas_m - meta_m})
                    meta_total += meta_m

                meta_intervalo = int(meta_total)
                estado = "HEALTHY" if ventas_intervalo >= meta_intervalo else "RISKY"
                badge = "🟢 HEALTHY" if estado == "HEALTHY" else "🟠 RISKY"
                delta = int(ventas_intervalo - meta_intervalo)

                # Antigüedad shown using the LAST month in interval (1st day of that month)
                last_mk = months_in_scope[-1]
                try:
                    yy, mm = last_mk.split("-")
                    ref_tenure = pd.Timestamp(year=int(yy), month=int(mm), day=1).normalize()
                except Exception:
                    ref_tenure = pd.Timestamp.today().normalize()

                am, ad, ing_dt = antiguedad_meses_y_dias(ej, ref_tenure)
                antig_txt = f"{am} meses {ad} días" if (am is not None and ad is not None) else "No disponible (default meta=12)"

                csc1, csc2, csc3, csc4 = st.columns(4, gap="medium")
                with csc1:
                    st.markdown("**Intervalo evaluado**")
                    st.write(f"{months_in_scope[0]} → {months_in_scope[-1]}  |  Meses: {len(months_in_scope)}")
                with csc2:
                    metric_card(
                        "Antigüedad (al 1ro del último mes)",
                        antig_txt,
                        sub=(f"Ingreso: {ing_dt:%d/%m/%Y}" if ing_dt is not None else None),
                    )
                with csc3:
                    metric_card("Ventas (intervalo)", fmt_int(ventas_intervalo))
                with csc4:
                    metric_card("Meta / Estado", f"{meta_intervalo} — {badge}", sub=f"Delta: {delta:+d}")

                # Optional breakdown per month
                if len(meta_rows) > 1:
                    st.markdown("#### Desglose por mes (según filtro)")
                    df_break = pd.DataFrame(meta_rows)
                    st.dataframe(
                        df_break.style.format({"Ventas": "{:,.0f}", "Meta": "{:,.0f}", "Delta": "{:+,.0f}"}),
                        hide_index=True,
                        width="stretch",
                    )

                # Chart sanity (Ventas vs Meta) for the interval
                bar_color = "#2ecc71" if estado == "HEALTHY" else "#e74c3c"
                meta_color = bar_color

                fig_sc = go.Figure()
                fig_sc.add_trace(
                    go.Bar(
                        x=["Ventas (intervalo)"],
                        y=[ventas_intervalo],
                        name="Ventas",
                        text=[ventas_intervalo],
                        textposition="outside",
                        marker=dict(color=bar_color),
                    )
                )
                fig_sc.add_trace(
                    go.Scatter(
                        x=["Ventas (intervalo)"],
                        y=[meta_intervalo],
                        mode="markers+text",
                        name="Meta",
                        text=[f"Meta: {meta_intervalo}"],
                        textposition="top center",
                        marker=dict(size=12, color=meta_color),
                    )
                )

                fig_sc.update_layout(
                    title=f"Sanity check — {ej} — Intervalo (meses/semanas seleccionados)",
                    yaxis_title="Cantidad",
                    height=360,
                )

                apply_plotly_theme(fig_sc)
                fig_sc.update_traces(marker_color=bar_color, selector=dict(type="bar"))
                fig_sc.update_traces(marker=dict(color=meta_color), selector=dict(type="scatter"))
                st.plotly_chart(fig_sc, width="stretch", key=f"t8_sanity_interval_{normalize_name(ej)}")

                # ---------- Desglose por ejecutivo del equipo (interval) ----------
                st.markdown("#### ✅ Desglose por Ejecutivo (equipo) — Ventas vs Meta mensual (según intervalo)")

                g = (
                    df_scope.groupby("EJECUTIVO", as_index=False)
                    .size()
                    .rename(columns={"size": "Ventas"})
                )

                # MetaIntervalo per ejecutivo = sum(meta per month in interval
                def _meta_intervalo_for_exec(exec_name: str) -> int:
                    total = 0
                    for mk2 in months_in_scope:
                        try:
                            yy2, mm2 = mk2.split("-")
                            ms2 = pd.Timestamp(year=int(yy2), month=int(mm2), day=1).normalize()
                        except Exception:
                            continue
                        total += int(meta_por_antiguedad(exec_name, ms2))
                    return int(total)

                g["MetaIntervalo"] = g["EJECUTIVO"].apply(_meta_intervalo_for_exec)

                def _antig_str(nm: str) -> str:
                    mm, dd, _ = antiguedad_meses_y_dias(nm, ref_tenure)
                    return f"{mm}m {dd}d" if (mm is not None and dd is not None) else "N/D"

                g["Antigüedad"] = g["EJECUTIVO"].apply(_antig_str)
                g["Delta"] = g["Ventas"] - g["MetaIntervalo"]
                g["Estado"] = np.where(g["Ventas"] >= np.ceil(g["MetaIntervalo"] / 2.0), "HEALTHY", "RISKY")


                def _rank_estado(x: str) -> int:
                    return 0 if x == "RISKY" else 1

                g["_rank"] = g["Estado"].apply(_rank_estado)
                g = g.sort_values(["_rank", "Delta", "Ventas"], ascending=[True, True, False]).drop(columns=["_rank"])

                # ✅ COLOR MAP: GREEN healthy, RED risky
                STATUS_COLORS = {"HEALTHY": "#2ecc71", "RISKY": "#e74c3c"}

                fig_team_exec = px.bar(
                    g.sort_values("Ventas", ascending=True),
                    x="Ventas",
                    y="EJECUTIVO",
                    orientation="h",
                    color="Estado",  # ✅ color by state
                    color_discrete_map=STATUS_COLORS,  # ✅ force green/red
                    title="Ventas por Ejecutivo (intervalo seleccionado)",
                    hover_data={"MetaIntervalo": True, "Antigüedad": True, "Delta": True, "Estado": True},
                    template=PLOTLY_TEMPLATE,
                )
                fig_team_exec.update_layout(
                    height=min(900, 140 + 30 * len(g)),
                    margin=dict(l=20, r=20, t=70, b=20),
                    legend_title_text="Estado",
                )
                apply_plotly_theme(fig_team_exec)
                st.plotly_chart(fig_team_exec, width="stretch", key=f"t8_team_exec_interval_{normalize_name(sup_tab_choice)}")

                show = g.rename(columns={"MetaIntervalo": "Meta intervalo"}).copy()
                st.dataframe(
                    show.style.format(
                        {
                            "Ventas": "{:,.0f}",
                            "Meta intervalo": "{:,.0f}",
                            "Delta": "{:+,.0f}",
                        }
                    ),
                    hide_index=True,
                    width="stretch",
                )

# ======================================================
# TAB 10: Interfaz Custom (HTML)
# ======================================================
with tabs[9]:
    st.markdown("## Sistema de Informes Custom")
    
    try:
        mes_sel_int = int(mes_sel)
    except Exception:
        today = date.today()
        mes_sel_int = today.year * 100 + today.month
        
    # ✅ 1. Determine Dynamic Interval based on Selected Month (NOT Sidebar start_dt/end_dt)
    m_start, m_end_full = month_bounds(mes_sel_int)
    today_dt = date.today()
    
    if m_start <= today_dt <= m_end_full:
        m_end = today_dt
    elif today_dt < m_start:
        m_end = m_start
    else:
        m_end = m_end_full
        
    # ✅ For En Transito traceability, use all fetched history
    history_start = prog_history["Fecha_Creacion_DT"].min() if ("prog_history" in locals() and not prog_history.empty) else m_start
    interval_start = history_start
    interval_end = end_dt

    # 2. Get dynamic strings for the period display
    current_period = mes_labels.get(mes_sel, str(mes_sel)).title()
    dias_restantes_val = float(dias_hab_restantes) if "dias_hab_restantes" in locals() else 0.0

    # ✅ Solo para abril, alinear la Interfaz Custom con 22.5 días hábiles de mes
    if int(mes_sel_int) % 100 == 4:
        dias_transcurridos_custom = float(dias_hab_transcurridos) if "dias_hab_transcurridos" in locals() else 0.0
        dias_restantes_val = max(0.0, 22.5 - dias_transcurridos_custom)
    
    if interval_start == interval_end:
        back_label = f"BACK {interval_start.strftime('%d/%m/%Y')}"
    else:
        back_label = f"BACK {interval_start.strftime('%d/%m/%Y')} → {interval_end.strftime('%d/%m/%Y')}"

    # 3. Map SQL supervisor names -> HTML IDs
    mf_norm_html = normalize_name("MARIA FERNANDA MARTINEZ BISTRAIN")
    ml_norm_html = normalize_name("MARIA LUISA MEZA GOEL")

    sup_map = {
        normalize_name("REYNA LIZZETTE MARTINEZ GARCIA"): "reyna",
        normalize_name("ALAN UZIEL SALAZAR AGUILAR"): "alan",
        normalize_name("CARLOS ALBERTO AGUILAR CANO"): "carlos",
        normalize_name("ALFREDO CABRERA PADRON"): "alfredo",
        normalize_name("JORGE MIGUEL UREÑA ZARATE"): "jorge",
        mf_norm_html: "maria",
        ml_norm_html: "maria",
    }

    # ✅ Ocultar solo en Interfaz Custom para María Luisa / bloque "maria"
    HIDE_CUSTOM_MARIA = {
        normalize_name("JESUS ABEL RODRIGUEZ ORTIZ"),
        normalize_name("AURORA EUGENIA HURTAZO SERRANO"),
        normalize_name("BELEN LOPEZ GONZALEZ"),
        normalize_name("PATRICIA PEREZ PEREZ"),
        normalize_name("JORGE MIGUEL UREÑA ZARATE"),
        normalize_name("CARMEN RIVAS GONZALEZ"),
        normalize_name("EMILIO RAFAEL CORNU AGUILAR"),
    }

    HIDE_CUSTOM_JORGE = {
        normalize_name("DIEGO GARCIA ZUÑIGA"),
        normalize_name("ERIC DE JESUS MORENO"),
        normalize_name("LUIS ALBERTO ROMANO MEJIA"),
    }


    # ---------------------------------------------------------
    # ✅ HELPERS
    # ---------------------------------------------------------
    _HTML_STOPWORDS = {"DE", "DEL", "LA", "LAS", "LOS", "Y", "DA", "DO"}

    def _html_tokens(s: str) -> set[str]:
        n = normalize_name(s)
        return {t for t in n.split() if t and t not in _HTML_STOPWORDS}

    def _resolve_sup_from_name(
        ej_norm: str,
        candidate_name: str,
        exact_sup_map: dict,
        fallback_sales_sup: dict,
        fallback_detalle_sup: dict,
        roster_records: list[dict],
    ):
        # 1) exact by normalized employee name
        if ej_norm in exact_sup_map and exact_sup_map[ej_norm] in sup_map:
            return exact_sup_map[ej_norm]

        # 2) fallback from ventas / detalle
        sup_norm = fallback_sales_sup.get(ej_norm) or fallback_detalle_sup.get(ej_norm)
        if sup_norm in sup_map:
            return sup_norm

        # 3) fuzzy by tokens against employee roster
        cand = str(candidate_name or "").strip()
        toks = _html_tokens(cand)
        if toks:
            best = None
            best_inter = 0
            best_score = 0.0

            for rec in roster_records:
                inter = len(toks & rec["tokens"])
                if inter == 0:
                    continue
                score = inter / max(len(toks), len(rec["tokens"]), 1)
                if (inter > best_inter) or (inter == best_inter and score > best_score):
                    best = rec
                    best_inter = inter
                    best_score = score

            if best and (best_score >= 0.70 or best_inter >= 3):
                return best["sup_norm"]

        return None

    # ---------------------------------------------------------
    # ✅ BUILD EMPLOYEE LOOKUP (same spirit as tránsito dashboard)
    #     Row-level supervisor assignment comes from employees,
    #     not from later post-hoc reconstruction.
    # ---------------------------------------------------------
    emp_html = empleados.copy()

    emp_html["Nombre"] = emp_html["Nombre"].astype(str).str.strip()
    emp_html["Jefe Inmediato"] = emp_html["Jefe Inmediato"].astype(str).str.strip()
    emp_html["Puesto"] = emp_html["Puesto"].astype(str).str.strip().str.upper()
    emp_html["Estatus"] = emp_html["Estatus"].astype(str).str.strip().str.upper()

    # Keep only non-supervisory sales/advisor-like rows
    emp_html = emp_html[
        (~emp_html["Puesto"].str.contains("SUPERV", na=False))
        & (~emp_html["Puesto"].str.contains("COORD", na=False))
        & (~emp_html["Puesto"].str.contains("GEREN", na=False))
        & (~emp_html["Puesto"].str.contains("JEFE", na=False))
    ].copy()

    emp_html["Nombre"] = emp_html["Nombre"].replace(
        {
            "CESAR JAHACIEL ALONSO GARCIAA": "CESAR JAHACIEL ALONSO GARCIA",
            "VICTOR BETANZO FUENTES": "VICTOR BETANZOS FUENTES",
        }
    )

    emp_html["EJ_NORM"] = emp_html["Nombre"].apply(normalize_name)
    emp_html["SUP_NORM"] = emp_html["Jefe Inmediato"].apply(normalize_name)

    # Prefer ACTIVO if duplicates exist
    emp_html["_ACTIVO_RANK"] = (emp_html["Estatus"] == "ACTIVO").astype(int)
    if "Fecha Ingreso" in emp_html.columns:
        emp_html["_ING_DT"] = pd.to_datetime(emp_html["Fecha Ingreso"], errors="coerce")
    else:
        emp_html["_ING_DT"] = pd.NaT

    emp_html = emp_html.sort_values(
        ["_ACTIVO_RANK", "_ING_DT", "Nombre"],
        ascending=[False, False, True],
    ).drop_duplicates(subset=["EJ_NORM"], keep="first")

    emp_html = emp_html[emp_html["SUP_NORM"].isin(sup_map)].copy()

    emp_exact_sup_map = emp_html.set_index("EJ_NORM")["SUP_NORM"].to_dict()
    emp_exact_name_map = emp_html.set_index("EJ_NORM")["Nombre"].to_dict()

    emp_roster_records = []
    for _, r in emp_html.iterrows():
        emp_roster_records.append(
            {
                "ej_norm": r["EJ_NORM"],
                "sup_norm": r["SUP_NORM"],
                "name": str(r["Nombre"]).strip(),
                "tokens": _html_tokens(str(r["Nombre"]).strip()),
            }
        )

    # ---------------------------------------------------------
    # ✅ FALLBACK MAPS FROM REAL SALES / DETALLE
    # ---------------------------------------------------------
    last_sup_by_ej = {}
    last_name_by_ej = {}

    if not df_base.empty:
        ventas_team = df_base.copy()
        ventas_team["EJECUTIVO"] = ventas_team["EJECUTIVO"].astype(str).str.strip()
        ventas_team["Supervisor"] = ventas_team["Supervisor"].astype(str).str.strip()
        ventas_team["EJECUTIVO_norm"] = ventas_team["EJECUTIVO"].apply(normalize_name)
        ventas_team["Supervisor_norm"] = ventas_team["Supervisor"].apply(normalize_name)

        ventas_team = ventas_team[ventas_team["Supervisor_norm"].isin(sup_map)].copy()

        if not ventas_team.empty:
            ventas_team["FECHA_DE_CAPTURA_SORT"] = pd.to_datetime(ventas_team["FECHA DE CAPTURA"], errors="coerce")
            ventas_team = ventas_team.sort_values(["FECHA_DE_CAPTURA_SORT", "EJECUTIVO"], ascending=[True, True])

            last_rows = ventas_team.drop_duplicates(subset=["EJECUTIVO_norm"], keep="last").copy()
            last_sup_by_ej = last_rows.set_index("EJECUTIVO_norm")["Supervisor_norm"].to_dict()
            last_name_by_ej = last_rows.set_index("EJECUTIVO_norm")["EJECUTIVO"].to_dict()

    sup_from_detalle = {}
    name_from_detalle = {}
    meta_from_detalle = {}

    if "df_detalle_full" in locals() and not df_detalle_full.empty:
        det_tmp = df_detalle_full.copy()
        det_tmp["Ejecutivo"] = det_tmp["Ejecutivo"].astype(str).str.strip()
        det_tmp["Supervisor"] = det_tmp["Supervisor"].astype(str).str.strip()

        det_tmp = det_tmp[
            (det_tmp["Ejecutivo"].str.upper() != "TOTAL")
            & (det_tmp["Supervisor"].str.upper() != "TOTAL")
        ].copy()

        det_tmp["EJ_NORM"] = det_tmp["Ejecutivo"].apply(normalize_name)
        det_tmp["SUP_NORM"] = det_tmp["Supervisor"].apply(normalize_name)
        det_tmp = det_tmp[det_tmp["SUP_NORM"].isin(sup_map)].copy()

        if "Ventas" in det_tmp.columns:
            det_tmp["Ventas"] = pd.to_numeric(det_tmp["Ventas"], errors="coerce").fillna(0)
            det_tmp = det_tmp.sort_values(["Ventas", "Ejecutivo"], ascending=[False, True])
        else:
            det_tmp = det_tmp.sort_values(["Ejecutivo"])

        det_tmp = det_tmp.drop_duplicates(subset=["EJ_NORM"], keep="first").copy()
        sup_from_detalle = det_tmp.set_index("EJ_NORM")["SUP_NORM"].to_dict()
        name_from_detalle = det_tmp.set_index("EJ_NORM")["Ejecutivo"].to_dict()

        if "Meta" in det_tmp.columns:
            meta_from_detalle = (
                pd.to_numeric(det_tmp["Meta"], errors="coerce")
                .fillna(0)
                .groupby(det_tmp["EJ_NORM"])
                .first()
                .to_dict()
            )

    # ---------------------------------------------------------
    # ✅ LOAD PROGRAMACION BASE
    # ---------------------------------------------------------
    prog_history = fetch_programacion_history(end_dt)

    # ✅ For En Transito traceability, use all fetched history instead of selected month only
    history_start = prog_history["Fecha_Creacion_DT"].min() if not prog_history.empty else m_start
    interval_start = history_start
    interval_end = end_dt

    # base payload
    live_data_payload = {
        "diasRestantes": dias_restantes_val,
        "periodo": current_period,
        "backLabel": back_label,
        "supData": {k: {"ventas": 0, "backFeb": 0, "entrega": 0, "prep": 0, "solic": 0, "backoff": 0, "sinventa": 0} for k in sup_map.values()},
        "agents": {k: [] for k in sup_map.values()},
        "supMetas": {},
    }

    metas_map_html = load_metas_from_csv(int(mes_sel))
    metas_sup_map_html = load_metas_supervisor_from_excel(int(mes_sel))

    for sup_norm, sup_id in sup_map.items():

        # ✅ Primero usar la meta real del supervisor del Excel del mes.
        # ✅ Solo si no existe y es Maria Luisa, usar Maria Fernanda como fallback.
        _lookup_norm = sup_norm

        if _lookup_norm in metas_sup_map_html:
            mv = metas_sup_map_html[_lookup_norm]
        elif _lookup_norm == ml_norm_html:
            mv = metas_sup_map_html.get(mf_norm_html, 0)
        else:
            mv = 0

        try:
            meta_val = int(float(mv)) if pd.notna(mv) else 0
        except Exception:
            meta_val = 0

        # ✅ Si dos supervisoras caen en el mismo bloque HTML ("maria"),
        # conservar la meta válida más alta y no sobrescribirla.
        live_data_payload["supMetas"][sup_id] = max(
            int(live_data_payload["supMetas"].get(sup_id, 0)),
            meta_val,
        )

    # ---------------------------------------------------------
    # ✅ TRUE SALES (from ventas dashboard month selection)
    # ---------------------------------------------------------
    ventas_agents_df = pd.DataFrame(columns=["EJ_NORM", "SUP_NORM", "EJECUTIVO", "ventas"])
    if not df_base.empty:
        ventas_tmp = df_base.copy()
        ventas_tmp["EJECUTIVO"] = ventas_tmp["EJECUTIVO"].astype(str).str.strip()
        ventas_tmp["Supervisor"] = ventas_tmp["Supervisor"].astype(str).str.strip()
        ventas_tmp["EJ_NORM"] = ventas_tmp["EJECUTIVO"].apply(normalize_name)
        ventas_tmp["SUP_NORM"] = ventas_tmp["Supervisor"].apply(normalize_name)

        ventas_tmp = ventas_tmp[ventas_tmp["SUP_NORM"].isin(sup_map)].copy()

        ventas_agents_df = (
            ventas_tmp.groupby(["EJ_NORM", "SUP_NORM", "EJECUTIVO"], as_index=False)
            .size()
            .rename(columns={"size": "ventas"})
            .sort_values(["ventas", "EJECUTIVO"], ascending=[False, True])
            .drop_duplicates(subset=["EJ_NORM"], keep="first")
            .copy()
        )

    # ---------------------------------------------------------
    # ✅ PROGRAMACION -> supervisor assignment at ROW LEVEL
    # ---------------------------------------------------------
    agent_records = {}

    def ensure_agent(ej_norm: str):
        if ej_norm not in agent_records:
            agent_records[ej_norm] = {
                "ej_norm": ej_norm,
                "sup_norm": None,
                "name": "",
                "meta": 0,
                "ventas": 0,
                "backFeb": 0,
                "entrega": 0,
                "prep": 0,
                "solic": 0,
                "backoff": 0,
                "sinventa": 0,
            }
        return agent_records[ej_norm]

    # ✅ Seed from full employee roster first
    # so all ejecutivos from each team appear,
    # even if they have 0 ventas and 0 programación.
    if not emp_html.empty:
        emp_seed = emp_html[emp_html["SUP_NORM"].isin(sup_map)].copy()
        emp_seed = emp_seed.sort_values(["SUP_NORM", "Nombre"], ascending=[True, True])
        emp_seed = emp_seed.drop_duplicates(subset=["EJ_NORM"], keep="first")

        for _, r in emp_seed.iterrows():
            ej_norm = r["EJ_NORM"]
            rec = ensure_agent(ej_norm)

            if r["SUP_NORM"] in sup_map:
                rec["sup_norm"] = r["SUP_NORM"]

            rec["name"] = str(r["Nombre"]).strip() or rec["name"]

            mv_seed = metas_map_html.get(ej_norm, 0)
            try:
                rec["meta"] = int(float(mv_seed)) if pd.notna(mv_seed) else int(rec["meta"] or 0)
            except Exception:
                pass

    # Seed from ventas reales after roster
    if not ventas_agents_df.empty:
        for _, r in ventas_agents_df.iterrows():
            ej_norm = r["EJ_NORM"]
            rec = ensure_agent(ej_norm)
            rec["sup_norm"] = r["SUP_NORM"] if r["SUP_NORM"] in sup_map else rec["sup_norm"]
            rec["name"] = str(r["EJECUTIVO"]).strip() or rec["name"]
            rec["ventas"] = int(r["ventas"])

    if not prog_history.empty:
        prog_assign = prog_history.copy()

        # Resolve supervisor row by row using employee roster first
        prog_assign["SUP_NORM_HTML"] = prog_assign["EJ_NORM"].map(emp_exact_sup_map)
        prog_assign["NAME_HTML"] = prog_assign["EJ_NORM"].map(emp_exact_name_map)

        # Fallback row-level where employee roster did not resolve
        unresolved_mask = ~prog_assign["SUP_NORM_HTML"].isin(sup_map)
        if unresolved_mask.any():
            prog_unres = prog_assign.loc[unresolved_mask].copy()
            resolved_sup = []
            resolved_name = []

            for _, rr in prog_unres.iterrows():
                ej_norm = rr["EJ_NORM"]
                vendor_name = str(rr.get("EJECUTIVO", "") or rr.get("Vendedor", "") or "").strip()

                sup_norm_res = _resolve_sup_from_name(
                    ej_norm=ej_norm,
                    candidate_name=vendor_name,
                    exact_sup_map=emp_exact_sup_map,
                    fallback_sales_sup=last_sup_by_ej,
                    fallback_detalle_sup=sup_from_detalle,
                    roster_records=emp_roster_records,
                )

                name_res = (
                    str(emp_exact_name_map.get(ej_norm, "")).strip()
                    or str(last_name_by_ej.get(ej_norm, "")).strip()
                    or str(name_from_detalle.get(ej_norm, "")).strip()
                    or vendor_name
                    or str(ej_norm).strip()
                )

                resolved_sup.append(sup_norm_res)
                resolved_name.append(name_res)

            prog_assign.loc[unresolved_mask, "SUP_NORM_HTML"] = resolved_sup
            prog_assign.loc[unresolved_mask, "NAME_HTML"] = resolved_name

        prog_assign["SUP_NORM_RESUELTO"] = prog_assign["SUP_NORM_HTML"]
        prog_assign["NAME_RESUELTO"] = prog_assign["NAME_HTML"]

        # ---------- Pipeline rows in selected month ----------
        # ✅ Use all history for En Transito traceability
        mask_pipeline = prog_assign["Fecha_Creacion_DT"].notna()
        df_pipeline = prog_assign.loc[mask_pipeline].copy()

        if not df_pipeline.empty:
            condlist = [
                df_pipeline["Estatus_upper"].eq("EN ENTREGA"),
                df_pipeline["Estatus_upper"].isin(["EN PREPARACION", "EN PREPARACIÓN"]),
                df_pipeline["Estatus_upper"].eq("SOLICITADO"),
                df_pipeline["Estatus_upper"].isin(["BACK OFFICE", "BACKOFFICE"]),
                (df_pipeline["Estatus_upper"] == "ENTREGADO") & df_pipeline["Venta_Vacia"],
            ]
            choicelist = ["entrega", "prep", "solic", "backoff", "sinventa"]
            df_pipeline["HTML_Cat"] = np.select(condlist, choicelist, default=None)

            df_pipeline = df_pipeline[
                df_pipeline["HTML_Cat"].notna()
                & df_pipeline["SUP_NORM_RESUELTO"].isin(sup_map)
            ].copy()

            if not df_pipeline.empty:
                pipe_group = (
                    df_pipeline.groupby(["EJ_NORM", "SUP_NORM_RESUELTO", "NAME_RESUELTO", "HTML_Cat"], as_index=False)
                    .size()
                    .rename(columns={"size": "n"})
                )

                for _, rr in pipe_group.iterrows():
                    ej_norm = rr["EJ_NORM"]
                    rec = ensure_agent(ej_norm)

                    if rr["SUP_NORM_RESUELTO"] in sup_map:
                        rec["sup_norm"] = rr["SUP_NORM_RESUELTO"]

                    if str(rr["NAME_RESUELTO"]).strip():
                        rec["name"] = str(rr["NAME_RESUELTO"]).strip()

                    cat = str(rr["HTML_Cat"]).strip()
                    n = int(rr["n"])
                    if cat == "entrega":
                        rec["entrega"] += n
                    elif cat == "prep":
                        rec["prep"] += n
                    elif cat == "solic":
                        rec["solic"] += n
                    elif cat == "backoff":
                        rec["backoff"] += n
                    elif cat == "sinventa":
                        rec["sinventa"] += n

        # ---------- True Back Office rows in selected month ----------
        bo_dt = choose_backoffice_dt_html(prog_assign, window_start=start_dt, window_end=end_dt)
        prog_assign["BO_DT"] = bo_dt
        prog_assign["BO_Fecha"] = prog_assign["BO_DT"].dt.date

        mask_bo = (
            (prog_assign["Estatus_upper"] != "CANC ERROR")
            & (prog_assign["BO_DT"].notna())
            & (prog_assign["BO_Fecha"] >= m_start)
            & (prog_assign["BO_Fecha"] <= m_end)
            & (prog_assign["SUP_NORM_RESUELTO"].isin(sup_map))
        )

        backoffice_df = prog_assign.loc[mask_bo].copy()

        if not backoffice_df.empty:
            bo_group = (
                backoffice_df.groupby(["EJ_NORM", "SUP_NORM_RESUELTO", "NAME_RESUELTO"], as_index=False)
                .size()
                .rename(columns={"size": "backFeb"})
            )

            for _, rr in bo_group.iterrows():
                ej_norm = rr["EJ_NORM"]
                rec = ensure_agent(ej_norm)

                if rr["SUP_NORM_RESUELTO"] in sup_map:
                    rec["sup_norm"] = rr["SUP_NORM_RESUELTO"]

                if str(rr["NAME_RESUELTO"]).strip():
                    rec["name"] = str(rr["NAME_RESUELTO"]).strip()

                rec["backFeb"] += int(rr["backFeb"])

    # ---------------------------------------------------------
    # ✅ Finalize agent records
    # ---------------------------------------------------------
    for ej_norm, rec in agent_records.items():
        if rec["sup_norm"] not in sup_map:
            # last fallback from ventas/detalle
            sup_norm_fallback = last_sup_by_ej.get(ej_norm) or sup_from_detalle.get(ej_norm)
            if sup_norm_fallback in sup_map:
                rec["sup_norm"] = sup_norm_fallback

        # ✅ NUEVO: todo lo de Maria Luisa se manda al bucket de Maria Fernanda
        if rec["sup_norm"] == ml_norm_html and mf_norm_html in sup_map:
            rec["sup_norm"] = mf_norm_html

        if not rec["name"]:
            rec["name"] = (
                str(emp_exact_name_map.get(ej_norm, "")).strip()
                or str(last_name_by_ej.get(ej_norm, "")).strip()
                or str(name_from_detalle.get(ej_norm, "")).strip()
                or str(ej_norm).strip()
            )

        mv = metas_map_html.get(ej_norm, meta_from_detalle.get(ej_norm, 0))
        try:
            rec["meta"] = int(float(mv)) if pd.notna(mv) else 0
        except Exception:
            rec["meta"] = 0

    # ---------------------------------------------------------
    # ✅ Push agents into payload
    # ---------------------------------------------------------
    for ej_norm, rec in agent_records.items():
        sup_norm = rec["sup_norm"]

        # ✅ NUEVO: todo lo de Maria Luisa se manda al bucket de Maria Fernanda
        if sup_norm == ml_norm_html and mf_norm_html in sup_map:
            sup_norm = mf_norm_html
            rec["sup_norm"] = sup_norm

        if sup_norm not in sup_map:
            continue

        sup_id = sup_map[sup_norm]

        # ✅ Ocultar solo en la Interfaz Custom dentro del bloque "maria"
        if (
            (sup_id == "maria" and normalize_name(rec["name"]) in HIDE_CUSTOM_MARIA)
            or
            (sup_id == "jorge" and normalize_name(rec["name"]) in HIDE_CUSTOM_JORGE)
        ):
            continue

        live_data_payload["agents"][sup_id].append(
            {
                "name": rec["name"],
                "meta": int(rec["meta"]),
                "data": {
                    "ventas": int(rec["ventas"]),
                    "backFeb": int(rec["backFeb"]),
                    "entrega": int(rec["entrega"]),
                    "prep": int(rec["prep"]),
                    "solic": int(rec["solic"]),
                    "backoff": int(rec["backoff"]),
                    "sinventa": int(rec["sinventa"]),
                },
            }
        )

    # ---------------------------------------------------------
    # 5C) Sort agents inside each supervisor
    # ---------------------------------------------------------
    for _sup_id in live_data_payload["agents"].keys():
        live_data_payload["agents"][_sup_id] = sorted(
            live_data_payload["agents"][_sup_id],
            key=lambda a: (
                -int(a.get("data", {}).get("ventas", 0) or 0),
                -(
                    int(a.get("data", {}).get("entrega", 0) or 0)
                    + int(a.get("data", {}).get("prep", 0) or 0)
                    + int(a.get("data", {}).get("solic", 0) or 0)
                    + int(a.get("data", {}).get("backoff", 0) or 0)
                    + int(a.get("data", {}).get("sinventa", 0) or 0)
                ),
                str(a.get("name", "")),
            ),
        )

    # ---------------------------------------------------------
    # 5D) Recompute supervisor totals FROM UNIQUE AGENTS
    # ---------------------------------------------------------
    for _sup_id, _agents in live_data_payload["agents"].items():
        live_data_payload["supData"][_sup_id]["ventas"] = int(sum(a["data"].get("ventas", 0) for a in _agents))
        live_data_payload["supData"][_sup_id]["backFeb"] = int(sum(a["data"].get("backFeb", 0) for a in _agents))
        live_data_payload["supData"][_sup_id]["entrega"] = int(sum(a["data"].get("entrega", 0) for a in _agents))
        live_data_payload["supData"][_sup_id]["prep"] = int(sum(a["data"].get("prep", 0) for a in _agents))
        live_data_payload["supData"][_sup_id]["solic"] = int(sum(a["data"].get("solic", 0) for a in _agents))
        live_data_payload["supData"][_sup_id]["backoff"] = int(sum(a["data"].get("backoff", 0) for a in _agents))
        live_data_payload["supData"][_sup_id]["sinventa"] = int(sum(a["data"].get("sinventa", 0) for a in _agents))

    # 6. Convert to JSON and Inject into HTML
    import os
    import json
    import streamlit.components.v1 as components
    
    json_data = json.dumps(live_data_payload)
    
    try:
        current_dir = os.path.dirname(__file__)
        html_path = os.path.join(current_dir, "dashboard.html")
        
        with open(html_path, "r", encoding="utf-8") as f:
            html_content = f.read()

        html_content = html_content.replace(
            "/*PYTHON_INJECTS_JSON_HERE*/ null",
            json_data
        )

        components.html(html_content, height=900, scrolling=True)

    except FileNotFoundError:
        st.error("No se encontró el archivo 'dashboard.html'. Asegúrate de que esté en la misma carpeta.")

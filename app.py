from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from bb_core import (
    build_month_editor,
    editor_to_store_rows,
    filter_df,
    fmt_int,
    fmt_pct,
    load_bb_store,
    make_excel_export,
    month_start,
    normalize_bb_store,
    options_for,
    read_excel_template,
    save_bb_store,
    safe_filename,
    sku_options,
)

APP_VERSION = "V6 BUILDING BLOCKS"
APP_DIR = Path("data")
BB_STORE_PATH = APP_DIR / "building_blocks.csv"

FILTER_ORDER = ["SOCIO_COMERCIAL", "MARCA_BU", "DIVISION", "FUERZA_DE_VENTA", "TERRITORIO"]

BUILDING_BLOCK_OPTIONS = [
    "",
    "Distribucion",
    "Precio / Promo",
    "Activacion",
    "Supply",
    "NPI / Lanzamiento",
    "EOL / Descontinuacion",
    "Otros",
]

CAUSAL_OPTIONS = [
    "",
    "Distribucion numerica",
    "Promocion",
    "Activacion comercial",
    "Quiebre / Fill Rate",
    "Estacionalidad",
    "Competencia",
    "Precio",
    "NPI",
    "EOL",
    "Otro",
]

st.set_page_config(page_title="Generador de Building Blocks", page_icon="📦", layout="wide")


@st.cache_data(show_spinner=False)
def cached_read_excel(file_bytes: bytes):
    return read_excel_template(BytesIO(file_bytes))


def metric(label: str, value: str):
    st.metric(label, value)


def selected_label_to_idsku(options: pd.DataFrame, label: str) -> str:
    return str(options.loc[options["LABEL"] == label, "IDSKU"].iloc[0])


def selected_label_to_articulo(options: pd.DataFrame, label: str) -> str:
    return str(options.loc[options["LABEL"] == label, "ARTICULO"].iloc[0])


def monthly_series(df: pd.DataFrame, periods: list[str]) -> pd.DataFrame:
    if df.empty or not periods:
        return pd.DataFrame(columns=["MES", "VALOR"])
    tmp = df.melt(value_vars=periods, var_name="MES", value_name="VALOR")
    tmp["MES"] = pd.to_datetime(tmp["MES"], errors="coerce")
    tmp["VALOR"] = pd.to_numeric(tmp["VALOR"], errors="coerce").fillna(0.0)
    return tmp.groupby("MES", as_index=False)["VALOR"].sum().sort_values("MES")


def monthly_fill_rate(df: pd.DataFrame, periods: list[str]) -> pd.DataFrame:
    if df.empty or not periods:
        return pd.DataFrame(columns=["MES", "FILL_RATE"])
    tmp = df.melt(value_vars=periods, var_name="MES", value_name="FILL_RATE")
    tmp["MES"] = pd.to_datetime(tmp["MES"], errors="coerce")
    tmp["FILL_RATE"] = pd.to_numeric(tmp["FILL_RATE"], errors="coerce")
    return tmp.groupby("MES", as_index=False)["FILL_RATE"].mean().dropna().sort_values("MES")


def make_long_chart(hist_ts: pd.DataFrame, editor_df: pd.DataFrame) -> pd.DataFrame:
    frames = []
    if hist_ts is not None and not hist_ts.empty:
        h = hist_ts[["MES", "VALOR"]].copy()
        h["Serie"] = "Historico ventas"
        frames.append(h)
    if editor_df is not None and not editor_df.empty:
        f = editor_df[["MES", "FCST Base", "FCST ajustado solicitado"]].copy()
        f = f.melt(
            id_vars=["MES"],
            value_vars=["FCST Base", "FCST ajustado solicitado"],
            var_name="Serie",
            value_name="VALOR",
        )
        frames.append(f)
    if not frames:
        return pd.DataFrame(columns=["MES", "Serie", "VALOR"])
    out = pd.concat(frames, ignore_index=True)
    out["MES"] = pd.to_datetime(out["MES"], errors="coerce")
    out["VALOR"] = pd.to_numeric(out["VALOR"], errors="coerce")
    return out.dropna(subset=["MES", "VALOR"])


def recalc_editor_policy(editor: pd.DataFrame) -> pd.DataFrame:
    from bb_core import policy_status

    out = editor.copy()
    out["FCST Base"] = pd.to_numeric(out["FCST Base"], errors="coerce").fillna(0.0)
    out["FCST ajustado solicitado"] = pd.to_numeric(out["FCST ajustado solicitado"], errors="coerce").fillna(out["FCST Base"])
    out["Holgura %"] = pd.to_numeric(out["Holgura %"], errors="coerce").fillna(0.0)

    estados = []
    deltas = []
    margins = []
    for _, r in out.iterrows():
        pol = policy_status(r["FCST Base"], r["FCST ajustado solicitado"], r["Holgura %"], r["Zona"])
        estados.append(pol["ESTADO_POLITICA"])
        deltas.append(pol["BB_NETO"])
        margins.append(pol["MARGEN_PERMITIDO"])
    out["BB neto"] = deltas
    out["Margen +/- permitido"] = margins
    out["Estado politica"] = estados
    return out


def create_blank_template() -> bytes:
    """Create a blank workbook compatible with read_excel_template()."""
    output = BytesIO()
    today = pd.Timestamp.today().replace(day=1)
    hist_months = pd.date_range(today - pd.DateOffset(months=24), periods=24, freq="MS")
    fcst_months = pd.date_range(today, periods=18, freq="MS")
    all_months = pd.date_range(today - pd.DateOffset(months=12), periods=30, freq="MS")

    month_cols_hist = [m.strftime("%Y-%m-%d") for m in hist_months]
    month_cols_fcst = [m.strftime("%Y-%m-%d") for m in fcst_months]
    month_cols_fill = [m.strftime("%Y-%m-%d") for m in all_months]

    sheets = {
        "1_MAESTRO_ABCXYZ": pd.DataFrame(columns=[
            "IDSKU", "SKU", "MARCA_BU", "ABC", "XYZ", "ABCXYZ", "VENTA_12M", "CV_12M", "PARTICIPACION_12M"
        ]),
        "2_MAESTRO_HOLGURAS": pd.DataFrame(columns=[
            "IDSKU", "SKU", "MARCA_BU", "ABCXYZ", "CATEGORIA_REFERENCIA", "SUBCATEGORIA_REFERENCIA", "HOLGURA_PCT", "FUENTE"
        ]),
        "3_MAESTRO_CONGELAMIENTO": pd.DataFrame(columns=[
            "IDSKU", "SKU", "MARCA_BU", "AGRUPADO", "CONGELAMIENTO_DIAS", "CONGELAMIENTO_MESES", "FUENTE"
        ]),
        "4_IMPORT_HISTORICO": pd.DataFrame(columns=[
            "IDSKU", "SOCIO_COMERCIAL", "DIVISION", "FUERZA_DE_VENTA", "TERRITORIO", "MARCA_BU", "ARTICULO", "DESCRIPTION", "HIST_YEAR", "HIST_PERIOD", "PPY", "PPC", *month_cols_hist
        ]),
        "5_IMPORT_FILL_RATE": pd.DataFrame(columns=[
            "IDSKU", "SOCIO_COMERCIAL", "DIVISION", "FUERZA_DE_VENTA", "TERRITORIO", "MARCA_BU", "ARTICULO", "VARIABLE", *month_cols_fill
        ]),
        "6_IMPORT_FCST": pd.DataFrame(columns=[
            "IDSKU", "SOCIO_COMERCIAL", "DIVISION", "FUERZA_DE_VENTA", "TERRITORIO", "MARCA_BU", "ARTICULO", "DESCRIPCION", "TIPO_DE_REGISTRO", "UNIDADES", *month_cols_fcst
        ]),
    }

    readme = pd.DataFrame({
        "Campo": [
            "Uso",
            "Llave principal",
            "Meses",
            "Holgura",
            "Congelamiento",
            "Recomendacion",
        ],
        "Descripcion": [
            "Diligencie las hojas 1 a 6 y cargue este archivo en la app.",
            "IDSKU debe ser consistente entre todas las hojas.",
            "Las columnas mensuales deben mantenerse en formato YYYY-MM-DD.",
            "HOLGURA_PCT puede diligenciarse como decimal 0.10 o porcentaje 10.",
            "CONGELAMIENTO_MESES define cuantos meses son ZEF: 1=un mes, 2=dos meses, etc.",
            "No cambie los nombres de las hojas ni los encabezados.",
        ],
    })

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        readme.to_excel(writer, sheet_name="README", index=False)
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "bg_color": "#00A99D", "font_color": "white", "border": 1})
        note_fmt = wb.add_format({"text_wrap": True, "valign": "top"})
        for sheet_name, df in {"README": readme, **sheets}.items():
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            if len(df.columns) > 0:
                ws.autofilter(0, 0, max(1, len(df)), len(df.columns) - 1)
                for i, col in enumerate(df.columns):
                    ws.write(0, i, col, header_fmt)
                    width = min(max(12, len(str(col)) + 2), 28)
                    ws.set_column(i, i, width, note_fmt if sheet_name == "README" else None)
            if sheet_name == "README":
                ws.set_column(0, 0, 22)
                ws.set_column(1, 1, 80, note_fmt)
    return output.getvalue()


if "data" not in st.session_state:
    st.session_state.data = None
if "filename" not in st.session_state:
    st.session_state.filename = ""
if "bb_store" not in st.session_state:
    st.session_state.bb_store = load_bb_store(BB_STORE_PATH)

st.title("📦 Generador de Building Blocks")
st.success(f"✅ ESTAS EN LA VERSION {APP_VERSION}. Flujo SKU → alcance → edición mensual con holguras y congelamiento.")

with st.sidebar:
    st.header("1. Plantilla")
    st.download_button(
        "Descargar plantilla Excel en blanco",
        data=create_blank_template(),
        file_name="Plantilla_Building_Blocks_Input.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    uploaded_file = st.file_uploader("Sube la plantilla Excel mensual", type=["xlsx"])
    if uploaded_file is not None:
        if st.button("Cargar / recargar plantilla", use_container_width=True):
            try:
                st.session_state.data = cached_read_excel(uploaded_file.getvalue())
                st.session_state.filename = uploaded_file.name
                st.success("Plantilla cargada correctamente.")
            except Exception as exc:
                st.error(f"No fue posible leer la plantilla: {exc}")

    st.divider()
    st.header("2. Ciclo")
    cycle_month = st.date_input("Mes abierto / ciclo", value=date.today().replace(day=1))
    cycle_month = month_start(cycle_month)
    estado_aprobacion = st.selectbox("Estado aprobación", ["Borrador", "Solicitado", "Aprobado", "Rechazado", "Escalado"])

    st.divider()
    if st.button("Recargar bitácora", use_container_width=True):
        st.session_state.bb_store = load_bb_store(BB_STORE_PATH)
        st.success("Bitácora recargada.")

if st.session_state.data is None:
    st.info("Carga la plantilla mensual desde el panel lateral para iniciar, o descarga la plantilla en blanco para diligenciar el input.")
    st.stop()

data = st.session_state.data
sheets = data["sheets"]
periods = data["periods"]
product_master = data["product_master"]
fcst = sheets["forecast"]
hist = sheets["historico"]
fill_rate = sheets["fill_rate"]

st.caption(f"Plantilla activa: {st.session_state.filename}")

sku_df = sku_options(fcst)
if sku_df.empty:
    st.warning("No se encontraron SKUs en la hoja de forecast.")
    st.stop()

st.markdown("## 1. Selecciona el SKU")
sku_label = st.selectbox("SKU / Artículo", sku_df["LABEL"].tolist(), key="sku_selector")
selected_idsku = selected_label_to_idsku(sku_df, sku_label)
selected_articulo = selected_label_to_articulo(sku_df, sku_label)

sku_fcst = fcst[fcst["IDSKU"].astype(str) == str(selected_idsku)].copy()
sku_hist = hist[hist["IDSKU"].astype(str) == str(selected_idsku)].copy()
sku_fr = fill_rate[fill_rate["IDSKU"].astype(str) == str(selected_idsku)].copy()

st.markdown("## 2. Define el alcance del ajuste")
filter_state: dict[str, str] = {}
working = sku_fcst.copy()
cols = st.columns(len(FILTER_ORDER))
for idx, col in enumerate(FILTER_ORDER):
    with cols[idx]:
        opts = options_for(working, col)
        label = col.replace("_", " ").title()
        selected = st.selectbox(label, opts, key=f"filter_{col}")
        filter_state[col] = selected
        if selected != "Todos" and col in working.columns:
            working = working[working[col].astype(str) == str(selected)]

fcst_scope = filter_df(sku_fcst, filter_state)
hist_scope = filter_df(sku_hist, filter_state)
fr_scope = filter_df(sku_fr, filter_state)

product_row = product_master[product_master["IDSKU"].astype(str) == str(selected_idsku)].head(1)
product = product_row.iloc[0].to_dict() if not product_row.empty else {}
abcxyz = product.get("ABCXYZ", "")
holgura = product.get("HOLGURA_PCT", 0)
freeze_months = product.get("CONGELAMIENTO_MESES", 0)
venta_12m = product.get("VENTA_12M", 0)
cv_12m = product.get("CV_12M", 0)
agrupado = product.get("AGRUPADO", "")
marca_bu = product.get("MARCA_BU", "")

editor_base = build_month_editor(data, selected_idsku, filter_state, cycle_month, st.session_state.bb_store)

st.markdown("## 3. Información clave del SKU")
k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1:
    metric("ABCXYZ", str(abcxyz))
with k2:
    metric("Holgura", fmt_pct(holgura))
with k3:
    metric("Congelamiento", f"{fmt_int(freeze_months)} mes(es)")
with k4:
    metric("Ventas 12M", fmt_int(venta_12m))
with k5:
    metric("CV 12M", f"{float(cv_12m):.2f}" if pd.notna(cv_12m) else "")
with k6:
    metric("Forecast base", fmt_int(editor_base["FCST Base"].sum()) if not editor_base.empty else "0")

info = pd.DataFrame([
    ["SKU", selected_articulo],
    ["Marca BU", marca_bu],
    ["Agrupado política", agrupado],
    ["Regla", "ZEF: cantidad de meses igual al congelamiento. ZC: mes siguiente a ZEF. ZF: meses posteriores. Mes cerrado: meses anteriores al ciclo."],
])
st.dataframe(info, hide_index=True, use_container_width=True, column_config={0: "Campo", 1: "Valor"})

hist_ts = monthly_series(hist_scope, periods["historico"])
fr_ts = monthly_fill_rate(fr_scope, periods["fill_rate"])
chart_data = make_long_chart(hist_ts, editor_base)

st.markdown("## 4. Visualización")
c1, c2 = st.columns([2, 1])
with c1:
    st.markdown("#### Histórico, forecast base y forecast ajustado")
    if chart_data.empty:
        st.info("No hay datos suficientes para graficar.")
    else:
        chart = (
            alt.Chart(chart_data)
            .mark_line(point=True)
            .encode(
                x=alt.X("MES:T", title="Mes"),
                y=alt.Y("VALOR:Q", title="Unidades"),
                color=alt.Color("Serie:N", title="Serie"),
                tooltip=[alt.Tooltip("MES:T", title="Mes"), alt.Tooltip("Serie:N"), alt.Tooltip("VALOR:Q", format=",.0f")],
            )
            .properties(height=320)
        )
        st.altair_chart(chart, use_container_width=True)

with c2:
    st.markdown("#### Fill Rate")
    if fr_ts.empty:
        st.info("No hay fill rate para este alcance.")
    else:
        fr_data = fr_ts.copy()
        fr_data["FILL_RATE"] = pd.to_numeric(fr_data["FILL_RATE"], errors="coerce")
        chart_fr = (
            alt.Chart(fr_data.dropna())
            .mark_line(point=True)
            .encode(
                x=alt.X("MES:T", title="Mes"),
                y=alt.Y("FILL_RATE:Q", title="Fill Rate", axis=alt.Axis(format="%")),
                tooltip=[alt.Tooltip("MES:T", title="Mes"), alt.Tooltip("FILL_RATE:Q", format=".1%")],
            )
            .properties(height=320)
        )
        st.altair_chart(chart_fr, use_container_width=True)

st.markdown("## 5. Edita building blocks por mes")
st.caption("Edita el FCST ajustado solicitado. La herramienta calcula BB neto contra FCST Base y valida holgura/zona automáticamente.")

if editor_base.empty:
    st.warning("Este SKU no tiene forecast base para el alcance seleccionado.")
    st.stop()

visible_cols = [
    "Activo", "SKU", "ABCXYZ", "Socio Comercial", "Marca BU", "Division / Canal", "Fuerza de Ventas", "Territorio",
    "Mes", "FCST Base", "Holgura %", "Margen +/- permitido", "FCST ajustado solicitado", "BB neto",
    "Building Block", "Causal", "Comentario", "Zona", "Estado politica", "Congelamiento meses"
]

# Ensure visible columns exist for backward compatibility.
for col in visible_cols:
    if col not in editor_base.columns:
        editor_base[col] = "" if col not in ["Congelamiento meses"] else freeze_months

edited_visible = st.data_editor(
    editor_base[visible_cols],
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    column_config={
        "Activo": st.column_config.SelectboxColumn(options=["Si", "No"], required=True),
        "FCST Base": st.column_config.NumberColumn(format="%.0f", disabled=True),
        "Holgura %": st.column_config.NumberColumn(format="%.1f%%", disabled=True),
        "Margen +/- permitido": st.column_config.NumberColumn(format="%.0f", disabled=True),
        "FCST ajustado solicitado": st.column_config.NumberColumn(format="%.0f", step=1),
        "BB neto": st.column_config.NumberColumn(format="%.0f", disabled=True),
        "Building Block": st.column_config.SelectboxColumn(options=BUILDING_BLOCK_OPTIONS),
        "Causal": st.column_config.SelectboxColumn(options=CAUSAL_OPTIONS),
    },
    disabled=[
        "SKU", "ABCXYZ", "Socio Comercial", "Marca BU", "Division / Canal", "Fuerza de Ventas", "Territorio",
        "Mes", "FCST Base", "Holgura %", "Margen +/- permitido", "BB neto", "Zona", "Estado politica", "Congelamiento meses",
    ],
)

edited = editor_base.copy()
for col in edited_visible.columns:
    edited[col] = edited_visible[col].values
edited = recalc_editor_policy(edited)

st.markdown("#### Validación automática")
st.dataframe(edited[visible_cols], use_container_width=True, hide_index=True)

b1, b2, b3 = st.columns([1, 1, 1])
with b1:
    if st.button("Guardar ajustes del SKU", type="primary", use_container_width=True):
        existing = normalize_bb_store(st.session_state.bb_store)
        from bb_core import scope_filter_mask
        remaining = existing[~scope_filter_mask(existing, selected_idsku, filter_state)].copy() if not existing.empty else existing
        new_rows = editor_to_store_rows(edited, selected_idsku, filter_state, "", "", estado_aprobacion, "")
        combined = pd.concat([remaining, new_rows], ignore_index=True)
        st.session_state.bb_store = normalize_bb_store(combined)
        save_bb_store(st.session_state.bb_store, BB_STORE_PATH)
        st.success("Ajustes guardados correctamente.")
        st.rerun()

with b2:
    if st.button("Limpiar ajustes de este SKU/alcance", use_container_width=True):
        existing = normalize_bb_store(st.session_state.bb_store)
        from bb_core import scope_filter_mask
        st.session_state.bb_store = existing[~scope_filter_mask(existing, selected_idsku, filter_state)].copy() if not existing.empty else existing
        save_bb_store(st.session_state.bb_store, BB_STORE_PATH)
        st.success("Ajustes eliminados para este SKU/alcance.")
        st.rerun()

with b3:
    export = make_excel_export(st.session_state.bb_store, edited)
    st.download_button(
        "Descargar soporte Excel",
        data=export,
        file_name=f"{safe_filename('soporte_BB_' + selected_idsku)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with st.expander("Ver bitácora completa"):
    st.dataframe(normalize_bb_store(st.session_state.bb_store), use_container_width=True, hide_index=True)

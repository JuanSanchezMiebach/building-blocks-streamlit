from __future__ import annotations

import io
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import numpy as np
import pandas as pd

EXPECTED_SHEETS = {
    "abcxyz": "1_MAESTRO_ABCXYZ",
    "holguras": "2_MAESTRO_HOLGURAS",
    "congelamiento": "3_MAESTRO_CONGELAMIENTO",
    "historico": "4_IMPORT_HISTORICO",
    "fill_rate": "5_IMPORT_FILL_RATE",
    "forecast": "6_IMPORT_FCST",
}

DIMENSION_COLS = ["SOCIO_COMERCIAL", "DIVISION", "FUERZA_DE_VENTA", "TERRITORIO", "MARCA_BU"]
DATE_STR_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+.*)?$")

BB_COLUMNS = [
    "ID_BB", "CREADO_EN", "ACTUALIZADO_EN", "ACTIVO", "IDSKU", "ARTICULO",
    "SOCIO_COMERCIAL", "DIVISION", "FUERZA_DE_VENTA", "TERRITORIO", "MARCA_BU",
    "MES_IMPACTO", "FCST_BASE", "FCST_AJUSTADO_SOLICITADO", "BB_NETO",
    "BUILDING_BLOCK", "CAUSAL", "COMENTARIO", "ABCXYZ", "HOLGURA_PCT",
    "MARGEN_PERMITIDO", "ZONA", "ESTADO_POLITICA", "CONGELAMIENTO_DIAS",
    "CONGELAMIENTO_MESES", "GERENTE_MARCA", "EVIDENCIA", "ESTADO_APROBACION", "APROBADOR"
]


def strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", str(value))
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize_column_name(col: Any) -> str:
    if isinstance(col, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(col).strftime("%Y-%m-%d")
    text = str(col).strip()
    if DATE_STR_RE.match(text):
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            return pd.Timestamp(parsed).strftime("%Y-%m-%d")
    text = strip_accents(text).upper().strip()
    text = text.replace("/", "_")
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Z0-9_]+", "", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def standardize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy().dropna(axis=0, how="all").dropna(axis=1, how="all")
    df.columns = [normalize_column_name(c) for c in df.columns]
    return df


def period_columns(df: pd.DataFrame) -> List[str]:
    return sorted([c for c in df.columns if isinstance(c, str) and DATE_STR_RE.match(c)], key=lambda x: pd.Timestamp(x))


def first_9_digits(value: Any) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    match = re.search(r"\d{9}", text)
    if match:
        return match.group(0)
    return text[:9] if text else None


def ensure_idsku(df: pd.DataFrame, source_candidates: Iterable[str]) -> pd.DataFrame:
    df = df.copy()
    if "IDSKU" not in df.columns:
        df["IDSKU"] = np.nan
    source_col = next((c for c in source_candidates if c in df.columns), None)
    if source_col:
        extracted = df[source_col].apply(first_9_digits)
        invalid = df["IDSKU"].isna() | df["IDSKU"].astype(str).str.lower().isin(["nan", "none", ""])
        df.loc[invalid, "IDSKU"] = extracted[invalid]
    df["IDSKU"] = df["IDSKU"].astype(str).str.strip()
    df.loc[df["IDSKU"].str.lower().isin(["nan", "none", ""]), "IDSKU"] = np.nan
    return df


def numeric_series(values: pd.Series) -> pd.Series:
    return pd.to_numeric(values, errors="coerce").fillna(0.0)


def normalize_percent(values: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(values, errors="coerce")
    return numeric.where(numeric <= 1, numeric / 100.0)


def read_excel_template(file_obj: Any) -> Dict[str, Any]:
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    xls = pd.ExcelFile(file_obj, engine="openpyxl")
    missing = [s for s in EXPECTED_SHEETS.values() if s not in xls.sheet_names]
    if missing:
        raise ValueError("Faltan hojas esperadas: " + ", ".join(missing))
    sheets: Dict[str, pd.DataFrame] = {}
    periods: Dict[str, List[str]] = {}
    for key, sheet in EXPECTED_SHEETS.items():
        df = standardize_dataframe(pd.read_excel(xls, sheet_name=sheet, dtype=object))
        df = ensure_idsku(df, ["SKU", "ARTICULO"])
        pcols = period_columns(df)
        for c in pcols:
            df[c] = numeric_series(df[c])
        periods[key] = pcols
        sheets[key] = df
    product_master = build_product_master(sheets)
    return {"sheets": sheets, "periods": periods, "product_master": product_master}


def build_product_master(sheets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    abc = sheets["abcxyz"].copy()
    hol = sheets["holguras"].copy()
    con = sheets["congelamiento"].copy()
    abc_cols = [c for c in ["IDSKU", "SKU", "MARCA_BU", "ABC", "XYZ", "ABCXYZ", "VENTA_12M", "CV_12M", "PARTICIPACION_12M"] if c in abc.columns]
    hol_cols = [c for c in ["IDSKU", "HOLGURA_PCT", "CATEGORIA_REFERENCIA", "SUBCATEGORIA_REFERENCIA"] if c in hol.columns]
    con_cols = [c for c in ["IDSKU", "AGRUPADO", "CONGELAMIENTO_DIAS", "CONGELAMIENTO_MESES"] if c in con.columns]
    master = abc[abc_cols].drop_duplicates("IDSKU")
    master = master.merge(hol[hol_cols].drop_duplicates("IDSKU"), on="IDSKU", how="left")
    master = master.merge(con[con_cols].drop_duplicates("IDSKU"), on="IDSKU", how="left")
    master["HOLGURA_PCT"] = normalize_percent(master.get("HOLGURA_PCT", pd.Series(index=master.index, dtype=float)))
    master["CONGELAMIENTO_DIAS"] = pd.to_numeric(master.get("CONGELAMIENTO_DIAS", pd.Series(index=master.index, dtype=float)), errors="coerce")
    master["CONGELAMIENTO_MESES"] = pd.to_numeric(master.get("CONGELAMIENTO_MESES", pd.Series(index=master.index, dtype=float)), errors="coerce")
    master["CONGELAMIENTO_MESES"] = master["CONGELAMIENTO_MESES"].fillna(np.ceil(master["CONGELAMIENTO_DIAS"] / 30.0)).fillna(0).astype(int)
    return master


def melt_periods(df: pd.DataFrame, periods: List[str], id_vars: Optional[List[str]], value_name: str) -> pd.DataFrame:
    if df.empty or not periods:
        return pd.DataFrame(columns=(id_vars or []) + ["MES", value_name])
    id_vars = [c for c in (id_vars or []) if c in df.columns]
    long = df.melt(id_vars=id_vars, value_vars=periods, var_name="MES", value_name=value_name)
    long["MES"] = pd.to_datetime(long["MES"], errors="coerce")
    long[value_name] = pd.to_numeric(long[value_name], errors="coerce").fillna(0.0)
    return long


def filter_df(df: pd.DataFrame, filters: Dict[str, str]) -> pd.DataFrame:
    out = df.copy()
    for col, val in filters.items():
        if val and val != "Todos" and col in out.columns:
            out = out[out[col].astype(str) == str(val)]
    return out


def options_for(df: pd.DataFrame, col: str) -> List[str]:
    if col not in df.columns:
        return ["Todos"]
    vals = sorted(df[col].dropna().astype(str).unique().tolist())
    return ["Todos"] + vals


def sku_options(fcst: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in ["IDSKU", "ARTICULO", "MARCA_BU"] if c in fcst.columns]
    out = fcst[cols].dropna(subset=["IDSKU"]).drop_duplicates("IDSKU").copy()
    if "ARTICULO" not in out.columns:
        out["ARTICULO"] = out["IDSKU"]
    out["LABEL"] = out["IDSKU"].astype(str) + " | " + out["ARTICULO"].astype(str)
    return out.sort_values("LABEL")


def month_start(value: Any) -> pd.Timestamp:
    ts = pd.Timestamp(value)
    return pd.Timestamp(ts.year, ts.month, 1)


def months_between(start: Any, end: Any) -> int:
    s = month_start(start)
    e = month_start(end)
    return (e.year - s.year) * 12 + (e.month - s.month)


def zone_for_period(period: Any, cycle_month: Any, freeze_months: int) -> str:
    offset = months_between(cycle_month, period)
    freeze_months = max(int(freeze_months or 0), 0)
    if offset < 0:
        return "Mes cerrado"
    # ZEF es proporcional al congelamiento en meses:
    # 30 dias = 1 mes en ZEF; 60 dias = 2 meses en ZEF; etc.
    # Por eso se usa offset < freeze_months, no offset <= freeze_months.
    if offset < freeze_months:
        return "ZEF"
    if offset == freeze_months:
        return "ZC"
    return "ZF"


def policy_status(fcst_base: float, fcst_adjusted: float, holgura_pct: float, zona: str) -> Dict[str, Any]:
    base = float(0 if pd.isna(fcst_base) else fcst_base)
    adjusted = float(0 if pd.isna(fcst_adjusted) else fcst_adjusted)
    delta = adjusted - base
    holgura = float(0 if pd.isna(holgura_pct) else holgura_pct)
    margen = abs(base) * holgura
    if zona == "Mes cerrado":
        estado = "Mes cerrado" if abs(delta) < 1e-9 else "Mes cerrado / no editable"
    elif zona in ["ZEF", "ZC"]:
        estado = "Permitido dentro de holgura" if abs(delta) <= margen + 1e-9 else "Excede holgura / escalar"
    else:
        estado = "Permitido - ZF"
    return {"BB_NETO": delta, "MARGEN_PERMITIDO": margen if zona in ["ZEF", "ZC"] else np.nan, "ESTADO_POLITICA": estado}


def empty_bb_store() -> pd.DataFrame:
    return pd.DataFrame(columns=BB_COLUMNS)


def normalize_bb_store(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return empty_bb_store()
    df = df.copy()
    for c in BB_COLUMNS:
        if c not in df.columns:
            df[c] = np.nan
    df = df[BB_COLUMNS]
    for c in ["FCST_BASE", "FCST_AJUSTADO_SOLICITADO", "BB_NETO", "HOLGURA_PCT", "MARGEN_PERMITIDO", "CONGELAMIENTO_DIAS", "CONGELAMIENTO_MESES"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in BB_COLUMNS:
        if c not in ["FCST_BASE", "FCST_AJUSTADO_SOLICITADO", "BB_NETO", "HOLGURA_PCT", "MARGEN_PERMITIDO", "CONGELAMIENTO_DIAS", "CONGELAMIENTO_MESES"]:
            df[c] = df[c].fillna("").astype(str)
    return df


def load_bb_store(path: Path) -> pd.DataFrame:
    if path.exists():
        return normalize_bb_store(pd.read_csv(path, dtype=object))
    return empty_bb_store()


def save_bb_store(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    normalize_bb_store(df).to_csv(path, index=False, encoding="utf-8-sig")


def scope_filter_mask(bb: pd.DataFrame, idsku: str, filters: Dict[str, str]) -> pd.Series:
    mask = bb["IDSKU"].astype(str) == str(idsku)
    for c in ["SOCIO_COMERCIAL", "DIVISION", "FUERZA_DE_VENTA", "TERRITORIO", "MARCA_BU"]:
        val = filters.get(c, "Todos")
        if val != "Todos":
            mask &= bb[c].astype(str) == str(val)
        else:
            mask &= bb[c].fillna("").astype(str).isin(["", "Todos"])
    return mask


def build_month_editor(data: Dict[str, Any], idsku: str, filters: Dict[str, str], cycle_month: Any, bb_store: pd.DataFrame) -> pd.DataFrame:
    fcst = data["sheets"]["forecast"]
    pcols = data["periods"]["forecast"]
    master = data["product_master"]
    product = master[master["IDSKU"].astype(str) == str(idsku)].head(1)
    info = product.iloc[0].to_dict() if not product.empty else {}

    sku_fcst = fcst[fcst["IDSKU"].astype(str) == str(idsku)].copy()
    filtered = filter_df(sku_fcst, filters)
    long = melt_periods(filtered, pcols, id_vars=[], value_name="FCST_BASE")
    base = long.groupby("MES", as_index=False)["FCST_BASE"].sum().sort_values("MES")

    holgura = float(0 if pd.isna(info.get("HOLGURA_PCT", np.nan)) else info.get("HOLGURA_PCT"))
    freeze_months = int(info.get("CONGELAMIENTO_MESES", 0) or 0)
    freeze_days = info.get("CONGELAMIENTO_DIAS", np.nan)
    abcxyz = info.get("ABCXYZ", "")
    articulo = info.get("SKU", "")
    if not articulo and not sku_fcst.empty and "ARTICULO" in sku_fcst.columns:
        articulo = sku_fcst["ARTICULO"].dropna().astype(str).iloc[0]

    rows = []
    bb = normalize_bb_store(bb_store)
    scope_bb = bb[scope_filter_mask(bb, idsku, filters)].copy() if not bb.empty else empty_bb_store()
    existing = {}
    if not scope_bb.empty:
        for _, r in scope_bb.iterrows():
            existing[pd.Timestamp(r["MES_IMPACTO"]).strftime("%Y-%m-%d")] = r.to_dict()

    for _, r in base.iterrows():
        mes = pd.Timestamp(r["MES"])
        mes_key = mes.strftime("%Y-%m-%d")
        fcst_base = float(r["FCST_BASE"])
        ex = existing.get(mes_key, {})
        fcst_adj = pd.to_numeric(ex.get("FCST_AJUSTADO_SOLICITADO", np.nan), errors="coerce")
        if pd.isna(fcst_adj):
            fcst_adj = fcst_base
        zona = zone_for_period(mes, cycle_month, freeze_months)
        pol = policy_status(fcst_base, fcst_adj, holgura, zona)
        rows.append({
            "Activo": ex.get("ACTIVO", "Si") or "Si",
            "SKU": articulo,
            "ABCXYZ": abcxyz,
            "Socio Comercial": filters.get("SOCIO_COMERCIAL", "Todos"),
            "Division / Canal": filters.get("DIVISION", "Todos"),
            "Fuerza de Ventas": filters.get("FUERZA_DE_VENTA", "Todos"),
            "Territorio": filters.get("TERRITORIO", "Todos"),
            "Marca BU": filters.get("MARCA_BU", "Todos"),
            "MES": mes,
            "Mes": mes.strftime("%Y-%m"),
            "FCST Base": fcst_base,
            "Holgura %": holgura,
            "Margen +/- permitido": pol["MARGEN_PERMITIDO"],
            "FCST ajustado solicitado": fcst_adj,
            "BB neto": pol["BB_NETO"],
            "Building Block": ex.get("BUILDING_BLOCK", ""),
            "Causal": ex.get("CAUSAL", ""),
            "Comentario": ex.get("COMENTARIO", ""),
            "Zona": zona,
            "Estado politica": pol["ESTADO_POLITICA"],
            "Congelamiento dias": freeze_days,
            "Congelamiento meses": freeze_months,
        })
    return pd.DataFrame(rows)


def editor_to_store_rows(editor: pd.DataFrame, idsku: str, filters: Dict[str, str], gerente: str, evidencia: str, estado_aprobacion: str, aprobador: str) -> pd.DataFrame:
    rows = []
    now = datetime.now().isoformat(timespec="seconds")
    for _, r in editor.iterrows():
        fcst_base = float(pd.to_numeric(r.get("FCST Base"), errors="coerce") or 0)
        fcst_adj = float(pd.to_numeric(r.get("FCST ajustado solicitado"), errors="coerce") or 0)
        bb_neto = fcst_adj - fcst_base
        has_meta = any(str(r.get(c, "")).strip() for c in ["Building Block", "Causal", "Comentario"])
        if abs(bb_neto) < 1e-9 and not has_meta:
            continue
        mes = pd.Timestamp(r["MES"]).strftime("%Y-%m-%d")
        row = {
            "ID_BB": f"{idsku}_{mes}_{abs(hash(str(filters))) % 1000000}",
            "CREADO_EN": now,
            "ACTUALIZADO_EN": now,
            "ACTIVO": r.get("Activo", "Si"),
            "IDSKU": idsku,
            "ARTICULO": r.get("SKU", ""),
            "SOCIO_COMERCIAL": filters.get("SOCIO_COMERCIAL", "Todos"),
            "DIVISION": filters.get("DIVISION", "Todos"),
            "FUERZA_DE_VENTA": filters.get("FUERZA_DE_VENTA", "Todos"),
            "TERRITORIO": filters.get("TERRITORIO", "Todos"),
            "MARCA_BU": filters.get("MARCA_BU", "Todos"),
            "MES_IMPACTO": mes,
            "FCST_BASE": fcst_base,
            "FCST_AJUSTADO_SOLICITADO": fcst_adj,
            "BB_NETO": bb_neto,
            "BUILDING_BLOCK": r.get("Building Block", ""),
            "CAUSAL": r.get("Causal", ""),
            "COMENTARIO": r.get("Comentario", ""),
            "ABCXYZ": r.get("ABCXYZ", ""),
            "HOLGURA_PCT": r.get("Holgura %", np.nan),
            "MARGEN_PERMITIDO": r.get("Margen +/- permitido", np.nan),
            "ZONA": r.get("Zona", ""),
            "ESTADO_POLITICA": r.get("Estado politica", ""),
            "CONGELAMIENTO_DIAS": r.get("Congelamiento dias", np.nan),
            "CONGELAMIENTO_MESES": r.get("Congelamiento meses", np.nan),
            "GERENTE_MARCA": gerente,
            "EVIDENCIA": evidencia,
            "ESTADO_APROBACION": estado_aprobacion,
            "APROBADOR": aprobador,
        }
        rows.append(row)
    if not rows:
        return empty_bb_store()
    return normalize_bb_store(pd.DataFrame(rows))


def make_excel_export(bb_store: pd.DataFrame, editor: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        normalize_bb_store(bb_store).to_excel(writer, sheet_name="Bitacora_BB", index=False)
        editor.drop(columns=["MES"], errors="ignore").to_excel(writer, sheet_name="Ajustes_SKU", index=False)
        for sheet, df in {"Bitacora_BB": normalize_bb_store(bb_store), "Ajustes_SKU": editor.drop(columns=["MES"], errors="ignore")}.items():
            ws = writer.sheets[sheet]
            ws.freeze_panes(1, 0)
            if len(df.columns) > 0:
                ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
                for i, col in enumerate(df.columns):
                    ws.set_column(i, i, min(max(12, len(str(col)) + 2), 35))
    return output.getvalue()


def fmt_int(v: Any) -> str:
    try:
        if pd.isna(v):
            return ""
        return f"{float(v):,.0f}".replace(",", ".")
    except Exception:
        return ""


def fmt_pct(v: Any) -> str:
    try:
        if pd.isna(v):
            return ""
        return f"{float(v):.1%}"
    except Exception:
        return ""


def safe_filename(text: str) -> str:
    text = strip_accents(text)
    return re.sub(r"[^A-Za-z0-9_-]+", "_", text).strip("_") or "export"

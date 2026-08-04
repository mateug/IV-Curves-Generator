import re
import pandas as pd

UNIT_FACTORS = {
    "I": {
        "A": 1.0,
        "mA": 1e-3,
        "uA": 1e-6,
        "μA": 1e-6,
        "nA": 1e-9,
        "pA": 1e-12,
    },
    "V": {
        "V": 1.0,
        "mV": 1e-3,
        "uV": 1e-6,
        "μV": 1e-6,
    },
}


def detect_header_indices(df_preview):
    """Busca en las primeras filas la cabecera que contiene columnas de I y V."""
    for idx, row in df_preview.iterrows():
        row_str = [str(x).strip().upper() if pd.notna(x) else "" for x in row]

        curr_i = []
        curr_v = []
        for col_idx, val in enumerate(row_str):
            is_i = False
            if val in ["I", "INTENSIDAD", "CURRENT"]:
                is_i = True
            elif "I (" in val or "INTENSIDAD (" in val or "CURRENT (" in val:
                is_i = True
            elif val.startswith("I_") or val.startswith("CURRENT_"):
                is_i = True

            is_v = False
            if val in ["V", "VOLTAJE", "VOLTAGE"]:
                is_v = True
            elif "V (" in val or "VOLTAJE (" in val or "VOLTAGE (" in val:
                is_v = True
            elif val.startswith("V_") or val.startswith("VOLTAGE_"):
                is_v = True

            if is_i:
                curr_i.append(col_idx)
            if is_v:
                curr_v.append(col_idx)

        if curr_i and curr_v:
            return {
                "header_row": idx,
                "i_cols_idx": curr_i,
                "v_cols_idx": curr_v,
            }

    return {"header_row": None, "i_cols_idx": [], "v_cols_idx": []}


def get_unit_from_str(s, var_type):
    """Busca unidades en paréntesis ej: (mA), (V)."""
    match = re.search(r"\((.*?)\)", s)
    if match:
        u = match.group(1).strip()
        if u in UNIT_FACTORS[var_type]:
            return u
    return "A" if var_type == "I" else "V"


def prepare_iv_series(df, v_col, i_col, source_i_unit, source_v_unit, target_i_unit, target_v_unit, swap=False):
    """Convierte y alinea las series de I y V para la gráfica IV."""
    base_i_factor = UNIT_FACTORS["I"].get(source_i_unit, 1.0)
    base_v_factor = UNIT_FACTORS["V"].get(source_v_unit, 1.0)

    target_i_factor = UNIT_FACTORS["I"][target_i_unit]
    target_v_factor = UNIT_FACTORS["V"][target_v_unit]

    mult_i = base_i_factor / target_i_factor
    mult_v = base_v_factor / target_v_factor

    i_vals = pd.to_numeric(df[i_col], errors="coerce").dropna() * mult_i
    v_vals = pd.to_numeric(df[v_col], errors="coerce").dropna() * mult_v

    min_len = min(len(i_vals), len(v_vals))
    i_vals = i_vals.iloc[:min_len].reset_index(drop=True)
    v_vals = v_vals.iloc[:min_len].reset_index(drop=True)

    x_vals = v_vals if not swap else i_vals
    y_vals = i_vals if not swap else v_vals
    return x_vals, y_vals


def get_clean_column_name(col_name):
    if not isinstance(col_name, str):
        return str(col_name)
    return "".join(c for c in col_name.lower() if c.isalnum())


def prepare_summary_dataframe(path):
    """Lee la hoja resumen y la normaliza para el gráfico Isc."""
    try:
        df = pd.read_excel(path, sheet_name="resumen")
    except Exception:
        try:
            df = pd.read_excel(path, sheet_name=0)
        except Exception:
            return None

    fecha_col = None
    isc_col = None
    for col in df.columns:
        col_clean = get_clean_column_name(col)
        if "fecha" in col_clean and "inicio" in col_clean:
            fecha_col = col
        if col_clean == "isc" or "isc" in col_clean:
            isc_col = col

    if fecha_col is None or isc_col is None:
        return None

    df = df.rename(columns={fecha_col: "Fecha y hora inicio", isc_col: "Isc"})
    df["Fecha y hora inicio"] = pd.to_datetime(df["Fecha y hora inicio"], errors="coerce")
    df["Isc"] = pd.to_numeric(df["Isc"], errors="coerce")
    return df.dropna(subset=["Fecha y hora inicio", "Isc"])


def load_excel_file_info(path, filename):
    """Lee un archivo Excel y devuelve los metadatos necesarios para la UI y los gráficos."""
    df_preview = pd.read_excel(path, sheet_name="datos_IV", header=None, nrows=50)

    detection = detect_header_indices(df_preview)
    header_row = detection["header_row"]
    i_cols_idx = detection["i_cols_idx"]
    v_cols_idx = detection["v_cols_idx"]

    status = "✅"
    i_unit = "-"
    v_unit = "-"
    i_col_name = None
    v_col_name = None
    summary_df = None

    if header_row is None:
        status = "❌"
    else:
        row_raw = list(df_preview.iloc[header_row])
        i_col_name = row_raw[i_cols_idx[0]]
        v_col_name = row_raw[v_cols_idx[0]]

        if len(i_cols_idx) > 1 or len(v_cols_idx) > 1:
            status = "⚠️"

        i_unit = get_unit_from_str(str(i_col_name), "I")
        v_unit = get_unit_from_str(str(v_col_name), "V")
        summary_df = prepare_summary_dataframe(path)

    return {
        "filename": filename,
        "path": path,
        "status": status,
        "header_row": header_row,
        "i_col": i_col_name,
        "v_col": v_col_name,
        "i_unit": i_unit,
        "v_unit": v_unit,
        "summary_df": summary_df,
    }

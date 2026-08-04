import pandas as pd
import matplotlib.dates as mdates


def build_style_map():
    return {
        "Línea continua (-)": "-",
        "Línea rayada (--)": "--",
        "Puntos (.)": ".",
        "Puntos y línea (.-)": ".-",
    }


def build_combined_iv(selected_items, same_voltage_series):
    combined_columns = []
    top_headers = []
    bottom_headers = []
    data_dict = {}

    unique_voltages = []
    voltage_keys = []

    for item in selected_items:
        try:
            df = pd.read_excel(item['path'], sheet_name='datos_IV', header=item['header_row'])
            v_values = pd.to_numeric(df[item['v_col']], errors='coerce').dropna().reset_index(drop=True)
            i_values = pd.to_numeric(df[item['i_col']], errors='coerce').dropna().reset_index(drop=True)
            length = min(len(v_values), len(i_values))
            v_values = v_values.iloc[:length]
            i_values = i_values.iloc[:length]
        except Exception:
            continue

        matched_index = None
        for idx, existing_v in enumerate(unique_voltages):
            if same_voltage_series(existing_v, v_values):
                matched_index = idx
                break

        if matched_index is None:
            unique_voltages.append(v_values)
            key = f'V_{len(unique_voltages)}'
            voltage_keys.append(key)
            combined_columns.append(key)
            top_headers.append(item['filename'])
            bottom_headers.append('V')
            data_dict[key] = v_values

            key_i = f'I_{len(unique_voltages)}_{item["filename"]}'
            combined_columns.append(key_i)
            top_headers.append(item['filename'])
            bottom_headers.append('I')
            data_dict[key_i] = i_values
        else:
            key_i = f'I_{matched_index + 1}_{item["filename"]}'
            combined_columns.append(key_i)
            top_headers.append(item['filename'])
            bottom_headers.append('I')
            data_dict[key_i] = i_values

    max_len = max((len(col) for col in data_dict.values()), default=0)
    for key, series in data_dict.items():
        if len(series) < max_len:
            data_dict[key] = series.reindex(range(max_len))

    if not combined_columns:
        return pd.DataFrame()

    df_combined = pd.DataFrame({key: data_dict[key] for key in combined_columns})
    df_combined.columns = pd.MultiIndex.from_arrays([top_headers, bottom_headers])
    return df_combined


def build_combined_summary(excel_data, selected_files):
    rows = []
    for item in excel_data:
        if item['filename'] not in selected_files or item['status'] not in ['✅', '⚠️']:
            continue
        summary_df = item.get('summary_df')
        if summary_df is None or summary_df.empty:
            continue
        temp = summary_df.copy()
        temp['Archivo'] = item['filename']
        rows.append(temp)

    if not rows:
        return pd.DataFrame(columns=['Archivo', 'Fecha y hora inicio', 'Isc'])

    return pd.concat(rows, ignore_index=True)


def flatten_multiindex_columns(df):
    if not isinstance(df.columns, pd.MultiIndex):
        return df

    new_cols = []
    for a, b in df.columns:
        new_cols.append(f"{a}_{b}")
    df = df.copy()
    df.columns = new_cols
    return df


def prepare_isc_plot_data(combined_summary):
    if combined_summary is None or combined_summary.empty:
        return []

    records = []
    for filename, group in combined_summary.groupby('Archivo'):
        series_dt = pd.to_datetime(group['Fecha y hora inicio'], dayfirst=True, errors='coerce')
        series_y = pd.to_numeric(group['Isc'], errors='coerce')
        mask = series_dt.notna() & series_y.notna()
        if not mask.any():
            continue
        records.append((filename, series_dt[mask], series_y[mask]))
    return records

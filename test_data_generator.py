import os
import pandas as pd
import numpy as np


def write_two_sheet_excel(path, datos_iv_df, resumen_df):
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        datos_iv_df.to_excel(writer, sheet_name='datos_IV', index=False)
        resumen_df.to_excel(writer, sheet_name='resumen', index=False)


def create_iv_dataframe(columns, num_rows=100, generate_curve=False, curve_multiplier=1.0):
    data = {}
    if generate_curve:
        v = np.linspace(-1, 1, num_rows)
        i = (1e-6 * (np.exp(v * 10) - 1) + np.random.normal(0, 1e-7, num_rows)) * curve_multiplier
        for col in columns:
            if 'V' in col.upper():
                data[col] = v
            elif 'I' in col.upper():
                data[col] = i
            else:
                data[col] = np.random.rand(num_rows)
    else:
        for col in columns:
            data[col] = np.random.rand(num_rows)
    return pd.DataFrame(data)


def create_summary_dataframe(start_date, isc_values):
    fechas = pd.date_range(start=start_date, periods=len(isc_values), freq='D')
    return pd.DataFrame({
        'Fecha y hora inicio': fechas.strftime('%Y-%m-%d %H:%M:%S'),
        'Isc': isc_values
    })


def create_metadata_file(path):
    num_rows = 50
    v = np.linspace(-1, 1, num_rows)
    i = (1e-6 * (np.exp(v * 10) - 1) + np.random.normal(0, 1e-7, num_rows)) * 0.8

    rows = [
        ['INFORME DE ENSAYO DE DIODO - PROBETA #42'],
        ['Fecha del ensayo:', '2026-07-05', 'Operario:', 'Antigravity'],
        ['Temperatura ambiente:', 25.4, 'Humedad relativa (%):', 45],
        [],
        ['Voltaje (V)', 'Intensidad (mA)', 'Columna Extra Inútil']
    ]
    for idx in range(num_rows):
        rows.append([v[idx], i[idx] * 1000, np.random.rand()])
    rows.append(['Nota final:', 'Este ensayo tiene algunas imperfecciones en los extremos.'])
    rows.append(['FIN DEL ARCHIVO'])

    df_meta = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df_meta.to_excel(writer, sheet_name='datos_IV', index=False, header=False)
        pd.DataFrame({
            'Fecha y hora inicio': ['2026-07-05 12:00:00'],
            'Isc': [30.504]
        }).to_excel(writer, sheet_name='resumen', index=False)


def main():
    folder = 'datos'
    if not os.path.exists(folder):
        os.makedirs(folder)

    # 1. Archivo válido normal (V y I en Amperios)
    df1 = create_iv_dataframe(['V (V)', 'I (A)', 'Time (s)'], generate_curve=True, curve_multiplier=1.0)
    resumen1 = create_summary_dataframe('2026-07-01 10:00:00', [30.5])
    write_two_sheet_excel(os.path.join(folder, 'valido_A_V.xlsx'), df1, resumen1)

    # 2. Archivo válido con uA y mV y columnas repetidas
    df2 = create_iv_dataframe(['V (mV)', 'V secundario (mV)', 'I (uA)', 'I secundario (uA)'], generate_curve=True, curve_multiplier=0.5)
    resumen2 = create_summary_dataframe('2026-07-02 10:00:00', [12.3])
    write_two_sheet_excel(os.path.join(folder, 'valido_uA_mV.xlsx'), df2, resumen2)

    # 3. Archivo sin columnas válidas (Cruz roja)
    df3 = create_iv_dataframe(['Time (s)', 'Temperature (C)', 'Pressure (Pa)'], generate_curve=False)
    resumen3 = create_summary_dataframe('2026-07-03 10:00:00', [0.0])
    write_two_sheet_excel(os.path.join(folder, 'invalido.xlsx'), df3, resumen3)

    # 4. Archivo con múltiples columnas posibles (Triángulo aviso)
    df4 = create_iv_dataframe(['V (V)', 'V_sec (V)', 'I (A)', 'I_leak (A)'], generate_curve=True, curve_multiplier=2.0)
    resumen4 = create_summary_dataframe('2026-07-04 10:00:00', [53.1])
    write_two_sheet_excel(os.path.join(folder, 'multiples_IV.xlsx'), df4, resumen4)

    # 5. Archivo con columnas Intensidad y Voltaje sin unidades
    df5 = create_iv_dataframe(['Voltaje', 'Intensidad'], generate_curve=True, curve_multiplier=1.5)
    resumen5 = create_summary_dataframe('2026-07-05 10:00:00', [18.7])
    write_two_sheet_excel(os.path.join(folder, 'valido_sin_unidades.xlsx'), df5, resumen5)

    # 6. Archivo con metadatos iniciales y cabecera no en la primera fila
    create_metadata_file(os.path.join(folder, 'valido_con_metadatos.xlsx'))

    print("Archivos de prueba generados exitosamente en la carpeta './datos'.")


if __name__ == '__main__':
    main()

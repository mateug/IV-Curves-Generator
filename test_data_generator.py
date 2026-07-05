import os
import pandas as pd
import numpy as np

def create_mock_excel(filename, columns, num_rows=100, generate_curve=False, curve_multiplier=1.0):
    data = {}
    if generate_curve:
        # Generar una curva IV típica de un diodo o panel solar (con ruido)
        v = np.linspace(-1, 1, num_rows)
        i = (1e-6 * (np.exp(v * 10) - 1) + np.random.normal(0, 1e-7, num_rows)) * curve_multiplier
        
        for col in columns:
            if 'V' in col:
                data[col] = v
            elif 'I' in col:
                data[col] = i
            else:
                data[col] = np.random.rand(num_rows)
    else:
        for col in columns:
            data[col] = np.random.rand(num_rows)
            
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

def main():
    folder = "datos"
    if not os.path.exists(folder):
        os.makedirs(folder)

    # 1. Archivo válido normal (V y I en Amperios)
    create_mock_excel(f"{folder}/valido_A_V.xlsx", ["V (V)", "I (A)", "Time (s)"], generate_curve=True, curve_multiplier=1.0)
    
    # 2. Archivo válido con uA y mV
    create_mock_excel(f"{folder}/valido_uA_mV.xlsx", ["V (mV)", "I (uA)"], generate_curve=True, curve_multiplier=0.5)
    
    # 3. Archivo sin columnas válidas (Cruz roja)
    create_mock_excel(f"{folder}/invalido.xlsx", ["Time (s)", "Temperature (C)", "Pressure (Pa)"])
    
    # 4. Archivo con múltiples columnas posibles (Triángulo aviso)
    create_mock_excel(f"{folder}/multiples_IV.xlsx", ["V (V)", "V_sec (V)", "I (A)", "I_leak (A)"], generate_curve=True, curve_multiplier=2.0)
    
    # 5. Archivo con columnas Intensidad y Voltaje sin unidades (asumiremos por defecto V y A en el script principal o preguntaremos)
    create_mock_excel(f"{folder}/valido_sin_unidades.xlsx", ["Voltaje", "Intensidad"], generate_curve=True, curve_multiplier=1.5)

    # 6. Archivo "formato raro" con metadatos iniciales, filas vacías, y notas al final
    # Generar datos de curva
    num_rows = 50
    v = np.linspace(-1, 1, num_rows)
    i = (1e-6 * (np.exp(v * 10) - 1) + np.random.normal(0, 1e-7, num_rows)) * 0.8
    
    # Crear filas como listas
    rows = [
        ["INFORME DE ENSAYO DE DIODO - PROBETA #42"],
        ["Fecha del ensayo:", "2026-07-05", "Operario:", "Antigravity"],
        ["Temperatura ambiente:", 25.4, "Humedad relativa (%):", 45],
        [], # fila vacía
        ["Voltaje (V)", "Intensidad (mA)", "Columna Extra Inútil"] # Cabecera real en fila 4 (0-indexed)
    ]
    for idx in range(num_rows):
        rows.append([v[idx], i[idx] * 1000, np.random.rand()]) # i convertida a mA
        
    rows.append(["Nota final:", "Este ensayo tiene algunas imperfecciones en los extremos."])
    rows.append(["FIN DEL ARCHIVO"])
    
    df_meta = pd.DataFrame(rows)
    df_meta.to_excel(f"{folder}/valido_con_metadatos.xlsx", index=False, header=False)

    print("Archivos de prueba generados exitosamente en la carpeta './datos'.")

if __name__ == "__main__":
    main()

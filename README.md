# Generador de Curvas IV

Herramienta con interfaz gráfica en Python diseñada para buscar, leer y visualizar archivos Excel que contengan datos de curvas Voltaje-Intensidad (IV).

## Características Principales

- **Selección y Escaneo Dinámico:** Elige una carpeta y el programa detectará automáticamente los archivos Excel, buscando columnas de Intensidad (I) y Voltaje (V). Con un botón para actualizar al instante si añades nuevos archivos.
- **Detección de Unidades y Conversión:** Detecta de forma inteligente las unidades indicadas entre paréntesis (ej. `(mV)`, `(μA)`) en las cabeceras de Excel y realiza las conversiones pertinentes a la unidad que prefieras visualizar en la gráfica. (Soporta `A, mA, uA, μA, nA, pA` y `V, mV, uV, μV`).
- **Personalización del Gráfico:** 
  - Ajuste de los límites del gráfico y reinicio automático.
  - Inversión de ejes (ascendente/descendente).
  - Intercambio de variables en los ejes (X por Y).
  - Cambio de estilo de la curva (continua, rayada, puntos, o puntos con línea).
- **Interactividad Intuitiva:** Información detallada mediante "hover". Al pasar el ratón por encima de una curva, se despliega una pequeña etiqueta que muestra a qué archivo pertenece, las columnas que usa y el valor exacto en el eje X e Y.
- **Exportación:** Exporta la gráfica generada con un solo clic a formato `.png` (con autoincremento para no sobrescribir) o elige tú el nombre y la extensión.

## Requisitos de Sistema e Instalación

El proyecto está construido sobre Python 3. Se requiere un entorno virtual con las siguientes dependencias:
- `pandas`
- `matplotlib`
- `openpyxl`

### Instalación (Manual)
```bash
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

1. **Ejecutable:** Puedes lanzar el programa haciendo doble clic en el archivo **`run.bat`**. Éste se encarga de activar automáticamente el entorno virtual y lanzar la interfaz gráfica.
2. Si prefieres la consola (estando dentro del entorno virtual), ejecuta:
```bash
python main.py
```
3. Generador de Pruebas: Dispones de un script `test_data_generator.py` que genera automáticamente 5 archivos Excel diferentes en la carpeta `./datos` para probar la herramienta en distintas situaciones (múltiples columnas, ausencia de datos, conversiones, etc).

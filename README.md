# Generador de Curvas IV

Herramienta con interfaz gráfica en Python diseñada para buscar, leer y visualizar archivos Excel que contengan datos de curvas Voltaje-Intensidad (IV).

## Estructura del Proyecto

El código principal ahora está organizado en la carpeta `src/`.
- `src/main.py`: entrada principal de la aplicación.
- `src/ui.py`: construcción de la interfaz Tkinter.
- `src/excel_loader.py`: lectura y normalización de hojas Excel.
- `src/graphing.py`: lógica de combinación de datos y preparación de gráficos.
- `src/exporting.py`: exportación dedicada a Excel.
- `src/test_data_generator.py`: script de datos de prueba.

La raíz del proyecto mantiene `run.bat` y `main.spec` para arranque y empaquetado, y `main.py` en la raíz ya no es el punto de entrada recomendado.

## Características Principales

- **Selección y Escaneo Dinámico:** Elige una carpeta y el programa detectará automáticamente los archivos Excel. Escanea dinámicamente las primeras 50 filas de cada archivo para localizar la cabecera real (tolerando títulos, filas vacías o metadatos iniciales) y localiza las columnas de Intensidad (I) y Voltaje (V).
- **Barra de Progreso en Escaneo:** Escanea carpetas grandes mostrando progreso visible mientras procesa cada archivo.
- **Selección múltiple más cómoda:** Usa `Ctrl + clic` para seleccionar o deseleccionar varios archivos en la lista.
- **Detección de Unidades y Conversión:** Detecta unidades en las cabeceras de Excel (por ejemplo `(mV)`, `(μA)`) y convierte los datos a las unidades deseadas en el gráfico.
- **Personalización del Gráfico:** 
  - Ajuste de límites del gráfico.
  - Reinicio de límites a valores automáticos.
  - Inversión de ejes X/Y.
  - Intercambio de variables en los ejes.
  - Cambio de estilo de la curva (continua, rayada, puntos, puntos con línea).
- **Interactividad Intuitiva:** Tooltip en hover sobre cada curva con información del archivo, columnas usadas y valores exactos.
- **Exportación separada y dedicada:** Exporta datos combinados a un Excel con hojas `datos_IV` y `resumen` mediante el módulo `src/exporting.py`.
- **Exportación de gráfico:** Guarda la gráfica generada como `.png`, `.pdf` u otros formatos compatibles.

## Requisitos de Sistema e Instalación

El proyecto está construido sobre Python 3. Se recomienda usar un entorno virtual con estas dependencias:
- `pandas`
- `matplotlib`
- `openpyxl`

### Instalación (Manual)
```bash
python -m venv venv
source venv/bin/activate   # Linux / macOS
# o
.\venv\Scripts\activate  # Windows
# si PowerShell lo bloquea por política de ejecución:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Uso

1. Ejecutar desde la raíz con el entorno activado:
```bash
python src/main.py
```

2. Alternativamente, en Windows puedes usar `run.bat`, que activa el entorno virtual y lanza la aplicación:
```bat
run.bat
```

3. El script `src/test_data_generator.py` crea archivos Excel de prueba en `./datos` para validar el comportamiento de la aplicación.

## Empaquetado a ejecutable

Este proyecto incluye una especificación de PyInstaller en `main.spec` que ya apunta a `src/main.py`. Para crear el ejecutable por primera vez:

```bash
python -m pip install pyinstaller
pyinstaller main.spec
```

El ejecutable resultante se generará en la carpeta `dist/`.

### Actualizar el ejecutable tras cambios

Cada vez que modifiques el código, vuelve a ejecutar PyInstaller con el mismo spec para regenerar el exe:

```bash
pyinstaller main.spec
```

Si quieres asegurarte de que la compilación es limpia, elimina primero las carpetas de construcción antiguas:

```bash
rm -rf build dist __pycache__
pyinstaller main.spec
```

> Nota: `main.spec` se usa para empaquetar la aplicación con la entrada actual en `src/main.py`. Mantén `src/` como la raíz del código fuente y actualiza el exe cada vez que cambies módulos dentro de `src/` o las dependencias en `requirements.txt`.

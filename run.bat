@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\activate.bat" (
    echo Creando entorno virtual por primera vez...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo Instalando dependencias...
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate.bat
)

python src\main.py

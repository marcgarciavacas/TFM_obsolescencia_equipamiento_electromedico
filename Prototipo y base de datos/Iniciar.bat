@echo off
REM Trabajar desde la carpeta donde esta este archivo.
cd /d "%~dp0"

echo Abriendo la aplicacion en el navegador...
echo Manten esta ventana abierta mientras utilizas la aplicacion.
py -3 -m streamlit run app.py
if errorlevel 1 (
    echo El inicio ha terminado con un error. Revisa el mensaje anterior.
    echo Si faltan librerias, ejecuta primero Instalar.bat.
    pause
    exit /b 1
)

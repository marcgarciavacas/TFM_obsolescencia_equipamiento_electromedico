@echo off
REM Trabajar desde la carpeta donde esta este archivo.
cd /d "%~dp0"

echo Instalando las librerias del proyecto...
py -3 -m pip install -r requirements.txt
if errorlevel 1 (
    echo No se ha completado la instalacion. Revisa el mensaje anterior, Python y la conexion a internet.
    pause
    exit /b 1
)

echo Instalacion completada. Ya puedes abrir Iniciar.bat.
pause

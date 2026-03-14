@echo off
CHCP 65001 > NUL
title Ejecutar Scripts de Python - Access Audit

:menu
cls
echo ===================================================
echo     Herramientas de Auditoria de Access (Python)
echo ===================================================
echo 1. Abrir Interfaz Grafica (access_audit_gui.py)
echo 2. Ejecutar Comparacion CLI modo ejemplo
echo 3. Salir
echo ===================================================
set /p opcion="Selecciona una opcion (1-3): "

if "%opcion%"=="1" goto gui
if "%opcion%"=="2" goto cli
if "%opcion%"=="3" goto fin

goto menu

:gui
cls
echo Iniciando access_audit_gui.py ...
python Python\access_audit_gui.py
echo.
pause
goto menu

:cli
cls
echo Ejecutando access_audit_compare.py con los archivos de prueba A.accdb y B.accdb ...
if not exist "A.accdb" (
    echo No se encuentra A.accdb. Asegurate de que los archivos de prueba esten en esta misma carpeta.
    pause
    goto menu
)
if not exist "B.accdb" (
    echo No se encuentra B.accdb. Asegurate de que los archivos de prueba esten en esta misma carpeta.
    pause
    goto menu
)
echo Comando: python Python\access_audit_compare.py --before "A.accdb" --after "B.accdb" --output "audit-report-python-A-vs-B.json"
echo.
python Python\access_audit_compare.py --before "A.accdb" --after "B.accdb" --output "audit-report-python-A-vs-B.json"
echo.
echo Proceso finalizado. Revisa el archivo generado: audit-report-python-A-vs-B.json
pause
goto menu

:fin
exit

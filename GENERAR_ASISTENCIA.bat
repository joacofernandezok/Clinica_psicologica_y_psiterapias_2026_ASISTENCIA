@echo off
title Generador Asistencia Clinica - UNC
color 05
echo ============================================================
echo      INICIANDO GENERADOR DE ASISTENCIA DINAMICA
echo ============================================================
echo.
echo [1/2] Verificando entorno local de Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [X] Error: Python no esta instalado en este equipo o no se agrego al PATH.
    pause
    exit
)
echo [2/2] Ejecutando motor algoritmico de asistencia...
python asistencia.py
if %errorlevel% neq 0 (
    echo [X] Error durante la ejecucion de asistencia.py. Verifique las librerias.
    pause
    exit
)
echo [DONE] Proceso finalizado correctamente. Presione una tecla para salir.
pause >nul

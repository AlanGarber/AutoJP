@echo off
rem Doble click para generar el reporte de hoy.
rem Para otro día: ejecutar.bat -Fecha 2026-11-12
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0reporte_no_respondidos.ps1" %*
pause

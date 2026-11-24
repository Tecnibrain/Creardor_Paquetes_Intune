@echo off
REM ============================================================
REM  Fabrica de Paquetes PSADT para Intune - Bancolombia
REM  Doble clic para abrir la herramienta.
REM  Usar solo en el equipo/VM de empaquetamiento.
REM ============================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Nuevo-Paquete.ps1"
if errorlevel 1 pause

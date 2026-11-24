@echo off
REM ============================================================
REM  Fabrica de Paquetes PSADT para Intune - VERSION WEB
REM  Doble clic: abre la Fabrica en el navegador (Edge).
REM  Solo funciona en este equipo (127.0.0.1). No cierres
REM  esta ventana mientras uses la pagina.
REM ============================================================
title Fabrica de Paquetes - version web (no cerrar)
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Nuevo-Paquete.ps1" -ModoWeb
if errorlevel 1 pause

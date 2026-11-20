@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo Python est introuvable. Installez Python depuis python.org puis relancez.
  pause
  exit /b 1
)
echo Collecte des taux OANDA en XOF (mois precedent par defaut)
echo Pour un autre mois : lancer_collecte.bat --mois 2026-01
echo.
py collecte_taux_xof.py %*
echo.
pause

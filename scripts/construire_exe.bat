@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo Python est introuvable. Installez Python depuis python.org puis relancez.
  pause
  exit /b 1
)
echo Installation des outils de fabrication...
py -m pip install --upgrade pyinstaller requests pandas openpyxl keyring
if errorlevel 1 goto erreur
echo.
echo Fabrication de collecte_taux_xof.exe (quelques minutes)...
py -m PyInstaller --onefile --console --noconfirm --name collecte_taux_xof --collect-submodules keyring --distpath "%~dp0." --workpath "%TEMP%\build_exe" --specpath "%TEMP%\build_exe" "%~dp0collecte_taux_xof.py"
if errorlevel 1 goto erreur
echo.
echo Termine : %~dp0collecte_taux_xof.exe
echo Double-cliquez dessus pour lancer la collecte du mois precedent.
pause
exit /b 0
:erreur
echo.
echo La fabrication a echoue. Copiez le message ci-dessus et envoyez-le moi.
pause
exit /b 1

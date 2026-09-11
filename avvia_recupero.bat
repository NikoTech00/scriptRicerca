@echo off
setlocal
cd /d "%~dp0"

echo Recupero incrementale di tutte le fonti configurate
echo Cartella progetto: %CD%
echo.

python .\scriptMedici.py "input\scriptMedici.xlsx" ^
  --massivo ^
  --max-documents 1000 ^
  --max-http-requests 2000 ^
  --mass-checkpoint-every 100 ^
  --output "output\risultato_massivo_recuperato.xlsx"

set "RUN_EXIT=%ERRORLEVEL%"
echo.
if not "%RUN_EXIT%"=="0" (
  echo Esecuzione terminata con errore %RUN_EXIT%.
) else (
  echo Esecuzione completata correttamente.
)
echo Premi un tasto per chiudere questa finestra.
pause >nul
exit /b %RUN_EXIT%

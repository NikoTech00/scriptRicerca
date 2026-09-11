@echo off
setlocal
cd /d "%~dp0"

echo Recupero istituzionale: Gemelli, IEO e MultiMedica
echo Cartella progetto: %CD%
echo.

python .\scriptMedici.py "input\scriptMedici.xlsx" ^
  --massivo ^
  --source-catalog ".\fonti_recupero_istituzionali.json" ^
  --refresh-sources ^
  --max-documents 1000 ^
  --max-http-requests 1000 ^
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

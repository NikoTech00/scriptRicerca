@echo off
setlocal
cd /d "%~dp0"

echo Recupero istituzionale: Policlinico Campus Bio-Medico di Roma
echo Cartella progetto: %CD%
echo.

python .\scriptMedici.py "input\scriptMedici.xlsx" ^
  --massivo ^
  --max-documents 500 ^
  --max-http-requests 1000 ^
  --mass-checkpoint-every 100 ^
  --output "output\risultato_massivo_recuperato.xlsx"

set "campus_exit=%ERRORLEVEL%"
echo.
if not "%campus_exit%"=="0" (
  echo Esecuzione terminata con errore %campus_exit%.
) else (
  echo Esecuzione completata correttamente.
)
echo Premi un tasto per chiudere questa finestra.
pause >nul
exit /b %campus_exit%

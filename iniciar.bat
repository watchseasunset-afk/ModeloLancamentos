@echo off
cd /d C:\ModLancamentos
echo Lancamentos Laterais - App Local
echo Abre o browser em http://localhost:5000
echo Ctrl+C para parar
echo.
start "" http://localhost:5000
python app.py
pause

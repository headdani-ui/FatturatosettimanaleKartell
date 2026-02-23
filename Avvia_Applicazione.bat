@echo off
title Avvio Applicazione Kartell
echo Avvio in corso dell'elaboratore CSV Kartell...
echo.

:: Verifica se streamlit è installato
where streamlit >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERRORE] Streamlit non sembra essere installato. 
    echo Assicurati di aver installato i requisiti con: pip install -r requirements.txt
    pause
    exit /b
)

:: Avvio dell'app
streamlit run app.py

pause

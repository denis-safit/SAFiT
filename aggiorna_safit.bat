@echo off
title Aggiornamento Portale SAFIT
color 0B
cls

echo ============================================================
echo           CARICAMENTO DATI SAFIT SU GITHUB
echo ============================================================
echo.

:: 1. Entra nella cartella corretta
cd /d "C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files"

echo [1/3] Sincronizzazione rami...
git pull origin main --quiet

echo.
echo [2/3] Preparazione file Excel...
git add .
git commit -m "Aggiornamento dati %date% %time%" --quiet

echo.
echo [3/3] Invio dati al portale online...
git push origin main

echo.
echo ============================================================
echo    DATI INVIATI! APERTURA PORTALE IN CORSO...
echo ============================================================

:: USA QUESTO INDIRIZZO SPECIFICO (IL TUO PORTALE PUBBLICO)
start chrome.exe "https://qey2qqomzpzjmuxb8mfm5h.streamlit.app"

echo.
echo Operazione completata. 
timeout /t 5
exit
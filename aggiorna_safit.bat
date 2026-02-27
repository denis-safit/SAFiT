@echo off
title Aggiornamento Portale SAFIT
color 0B
cls

echo ============================================================
echo            CARICAMENTO DATI SAFIT SU GITHUB
echo ============================================================
echo.

:: 1. Entra nella cartella corretta
cd /d "C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files"

echo [1/4] Controllo aggiornamenti remoti...
git pull origin main

echo.
echo [2/4] Forzatura scansione file modificati...
:: Questo comando forza Git a ricontrollare i timestamp dei file
git update-index --refresh > nul

echo.
echo [3/4] Preparazione file Excel...
git add --all
:: Rimosso --quiet per vedere cosa sta salvando effettivamente
git commit -m "Aggiornamento dati %date% %time%"

echo.
echo [4/4] Invio dati al portale online...
git push origin main

echo.
echo ============================================================
echo    DATI INVIATI! APERTURA PORTALE IN CORSO...
echo ============================================================

:: Apertura portale pubblico
start chrome.exe "https://qey2qqomzpzjmuxb8mfm5h.streamlit.app"

echo.
echo Operazione completata. 
timeout /t 5
exit
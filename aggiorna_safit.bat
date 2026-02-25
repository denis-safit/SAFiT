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
:: Il comando sotto "impacchetta" le modifiche
git commit -m "Aggiornamento dati %date% %time%" --quiet

echo.
echo [3/3] Invio dati al portale online...
:: Questo invia i dati su GitHub e aggiorna Streamlit
git push origin main

echo.
echo ============================================================
echo              OPERAZIONE COMPLETATA!
echo.
echo Il portale Streamlit si aggiornera' a breve.
echo Puoi chiudere questa finestra o premere un tasto.
echo ============================================================
pause
exit
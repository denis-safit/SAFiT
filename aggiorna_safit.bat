@echo off
title Aggiornamento Portale SAFIT - MODO FORZATO
color 0B
cls

echo ============================================================
echo            CARICAMENTO DATI SAFIT SU GITHUB
echo ============================================================
echo.

:: 1. Entra nella cartella corretta
cd /d "C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files"

echo [1/5] Congelamento dati locali...
:: Prima di tutto diciamo a Git che i tuoi file nuovi sono quelli buoni
git add --all
git commit -m "Aggiornamento dati locale %date% %time%"

echo.
echo [2/5] Allineamento con il Cloud (Senza sovrascrivere)...
:: Usiamo 'rebase' invece di 'pull' semplice per mettere i tuoi dati "sopra" quelli vecchi
git pull origin main --rebase

echo.
echo [3/5] Risoluzione conflitti automatica...
:: Se Git ha dubbi, gli diciamo di tenere i TUOI file (ours)
git checkout --ours .
git add --all
:: Se il rebase è in corso lo finiamo, altrimenti andiamo avanti
git rebase --continue 2>nul

echo.
echo [4/5] Invio dati finale...
git push origin main --force

echo.
echo ============================================================
echo    DATI INVIATI! IL TUO PC HA SOVRASCRITTO IL CLOUD.
echo ============================================================

:: Apertura portale pubblico
start chrome.exe "https://qey2qqomzpzjmuxb8mfm5h.streamlit.app"

echo.
echo Operazione completata. 
timeout /t 5
exit
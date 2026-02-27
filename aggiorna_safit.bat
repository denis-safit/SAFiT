@echo off
title Aggiornamento Portale SAFIT - AUTOMAZIONE TOTALE
color 0B
cls

echo ============================================================
echo            CARICAMENTO DATI SAFIT SU GITHUB
echo ============================================================
echo.

:: 1. DEFINIZIONE PERCORSI (Verifica che il DB PONTE sia sul Desktop)
set DIR_GIT=C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files
set DB_PONTE=%USERPROFILE%\Desktop\Estrattore_Safit.accdb

cd /d "%DIR_GIT%"

echo [1/7] Estrazione dati da ACCESS (Tramite file Ponte)...
:: Apre Access, lancia la macro e aspetta che finisca
start /wait "" "msaccess.exe" "%DB_PONTE%" /x Vai_Safit

echo.
echo [2/7] Aggiornamento Pivot ARCA (Refresh Silenzioso)...
:: Lancia lo script VBScript che abbiamo creato per la Pivot
cscript //nologo refresh_arca.vbs

echo.
echo [3/7] Congelamento dati locali...
git add --all
git commit -m "Auto-Update: %date% %time%"

echo.
echo [4/7] Allineamento con il Cloud (Rebase)...
git pull origin main --rebase

echo.
echo [5/7] Risoluzione conflitti automatica (Keep Ours)...
git checkout --ours .
git add --all
git rebase --continue 2>nul

echo.
echo [6/7] Invio dati finale al Portale...
git push origin main --force

echo.
echo ============================================================
echo    AGGIORNAMENTO COMPLETATO! APERTURA PORTALE...
echo ============================================================

:: Apre il portale per controllare il risultato
start chrome.exe "https://qey2qqomzpzjmuxb8mfm5h.streamlit.app"

echo.
echo Il sistema si chiudera tra 10 secondi.
timeout /t 10
exit
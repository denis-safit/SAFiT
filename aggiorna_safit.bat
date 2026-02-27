@echo off
title Aggiornamento Portale SAFIT - AUTOMAZIONE TOTALE
color 0B
cls

:: 1. DEFINIZIONE PERCORSI (Tutto nella stessa cartella)
set DIR_GIT=C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files
:: Ora il database è cercato direttamente nella cartella git-files
set DB_PONTE=%DIR_GIT%\Estrattore_Safit.accdb

cd /d "%DIR_GIT%"

echo [1/7] Estrazione dati da ACCESS (File Ponte in locale)...
if exist "%DB_PONTE%" (
    start /wait "" "msaccess.exe" "%DB_PONTE%" /x Vai_Safit
) else (
    echo ERRORE: File %DB_PONTE% non trovato!
    pause
    exit
)

echo.
echo [2/7] Aggiornamento Pivot ARCA...
cscript //nologo refresh_arca.vbs

echo.
echo [3/7] Congelamento dati locali...
git add --all
git commit -m "Auto-Update: %date% %time%"

echo.
echo [4/7] Allineamento con il Cloud (Rebase)...
git pull origin main --rebase

echo.
echo [5/7] Risoluzione conflitti (Keep Ours)...
git checkout --ours .
git add --all
git rebase --continue 2>nul

echo.
echo [6/7] Invio dati finale al Portale...
git push origin main --force

echo.
echo ============================================================
echo    AGGIORNAMENTO COMPLETATO!
echo ============================================================
timeout /t 10
exit
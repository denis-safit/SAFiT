@echo off
title Aggiornamento Portale SAFIT - DEFINITIVO
color 0B
cls

:: 1. IMPOSTAZIONE PERCORSI
set DIR_GIT=C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files
set DB_PONTE=%DIR_GIT%\Estrattore_Safit.accdb
set EXCEL_FILE=%DIR_GIT%\Avanzamento_access.xlsx

cd /d "%DIR_GIT%"

echo ============================================================
echo            FASE 1: ESTRAZIONE DATI ACCESS
echo ============================================================

echo [1/7] Pulizia file Excel precedente...
if exist "%EXCEL_FILE%" del /f /q "%EXCEL_FILE%"

echo [2/7] Lancio Access (Macro Vai_Safit)...
:: Lancio Access in background
start "" "msaccess.exe" "%DB_PONTE%" /x Vai_Safit

echo In attesa che Access generi il nuovo file...
:loop
timeout /t 2 >nul
if not exist "%EXCEL_FILE%" (
    echo ...sto ancora lavorando... attendere...
    goto loop
)
echo [OK] File Excel creato con successo!

echo.
echo ============================================================
echo            FASE 2: AGGIORNAMENTO ARCA E GIT
echo ============================================================

echo [3/7] Aggiornamento Pivot ARCA...
if exist refresh_arca.vbs (
    cscript //nologo refresh_arca.vbs
) else (
    echo Salto Pivot: refresh_arca.vbs non trovato.
)

echo [4/7] Preparazione file per GitHub...
git add --all
git commit -m "Auto-Update: %date% %time%"

echo [5/7] Sincronizzazione con il Cloud (Rebase)...
:: Questo risolve l'errore "unstaged changes" e allinea i dati
git pull origin main --rebase

echo [6/7] Controllo conflitti...
git checkout --ours .
git add --all
git rebase --continue 2>nul

echo [7/7] Invio finale al Portale Online...
git push origin main --force

echo.
echo ============================================================
echo    AGGIORNAMENTO COMPLETATO! IL PORTALE E' ONLINE.
echo ============================================================
timeout /t 10
exit
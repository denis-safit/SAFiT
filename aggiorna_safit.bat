@echo off
title Aggiornamento Portale SAFIT - DEFINITIVO + APERTURA WEB
color 0B
cls

set DIR_GIT=C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files
set DB_PONTE=%DIR_GIT%\Estrattore_Safit.accdb
set EXCEL_FILE=%DIR_GIT%\Avanzamento_access.xlsx

cd /d "%DIR_GIT%"

echo ============================================================
echo            FASE 1: ESTRAZIONE DATI ACCESS V3.8
echo ============================================================

echo [1/8] Pulizia file Excel precedente...
if exist "%EXCEL_FILE%" del /f /q "%EXCEL_FILE%"

echo [2/8] Lancio Access (Macro Vai_Safit)...
start "" "msaccess.exe" "%DB_PONTE%" /x Vai_Safit

echo In attesa che Access generi il nuovo file...
:loop
C:\Windows\System32\timeout.exe /t 2 >nul
if not exist "%EXCEL_FILE%" (
    echo ...sto ancora lavorando... attendere...
    goto loop
)
echo [OK] File Excel creato con successo!

echo.
echo ============================================================
echo            FASE 2: AGGIORNAMENTO ARCA E GIT
echo ============================================================

echo [3/8] Aggiornamento Pivot ARCA...
del /f /q "%TEMP%\arca_done.tmp" 2>nul
start "" /min cmd /c "%DIR_GIT%\run_arca.bat"
python "%DIR_GIT%\spinner_arca.py" "%TEMP%\arca_done.tmp"

echo [4/8] Preparazione file per GitHub...
git add --all
git commit -m "Auto-Update: %date% %time%"

echo [5/8] Sincronizzazione con il Cloud (Rebase)...
git pull origin main --rebase

echo [6/8] Controllo conflitti...
git checkout --ours .
git add --all
git rebase --continue 2>nul

echo [7/8] Invio finale al Portale Online...
git push origin main --force

echo.
echo ============================================================
echo    FASE 3: APERTURA PORTALE SAFIT
echo ============================================================
echo [8/8] Apertura Browser...
start chrome.exe "https://qey2qqomzpzjmuxb8mfm5h.streamlit.app"

echo.
echo ============================================================
echo    AGGIORNAMENTO COMPLETATO!
echo ============================================================
C:\Windows\System32\timeout.exe /t 2 >nul





pause 




start C:\Users\Venezian.Denis\Desktop\python_access_sql\Agente_arca_mail\avvia_agente.bat

exit
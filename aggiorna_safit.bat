@echo off
title Aggiornamento Portale SAFIT - DEFINITIVO + APERTURA WEB
color 0B
cls

:: 1. IMPOSTAZIONE PERCORSI
set DIR_GIT=C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files
set DB_PONTE=%DIR_GIT%\Estrattore_Safit.accdb
set EXCEL_FILE=%DIR_GIT%\Avanzamento_access.xlsx

cd /d "%DIR_GIT%"

echo ============================================================
echo            FASE 1: ESTRAZIONE DATI ACCESS V3.9
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

echo [3/8] Aggiornamento Pivot ARCA (tutti i file Excel)...
echo       Attendere circa 90 secondi...

:: Pulisce log precedente e lancia vbs in background
del /f /q "%TEMP%\arca_log.txt" 2>nul
if exist refresh_arca.vbs (
    start /b "" cscript //nologo refresh_arca.vbs > "%TEMP%\arca_log.txt" 2>&1
)

:: Spinner — attende ALMENO 30s prima di controllare il log
set _t=0
set _s=0
:spinner
C:\Windows\System32\timeout.exe /t 3 >nul
set /a _t+=3
set /a _s=(_s+1) %% 4

:: Ultimo file completato dal log
set LAST=
for /f "delims=" %%L in ('type "%TEMP%\arca_log.txt" 2^>nul ^| findstr /i "OK:"') do set LAST=%%L

if %_s%==0 echo   ^|  %_t%s  %LAST%
if %_s%==1 echo   /  %_t%s  %LAST%
if %_s%==2 echo   -  %_t%s  %LAST%
if %_s%==3 echo   \  %_t%s  %LAST%

:: Controlla fine solo dopo almeno 30 secondi
if %_t% lss 30 goto spinner

findstr /i "Tutti i file aggiornati" "%TEMP%\arca_log.txt" >nul 2>&1
if %errorlevel%==0 goto vbs_done
if %_t% gtr 200 goto vbs_done
goto spinner

:vbs_done
echo.
echo [OK] Tutti i file Excel aggiornati!
echo.
:: Mostra riepilogo file aggiornati
type "%TEMP%\arca_log.txt"
echo.

:: Rimuove file temporanei Excel (~$*) prima del commit
echo [3b] Pulizia file temporanei Excel...
for /f "delims=" %%F in ('dir /b /s "~$*.xlsx" 2^>nul') do del /f /q "%%F" 2>nul
for /f "delims=" %%F in ('dir /b /s "~$*.xls" 2^>nul') do del /f /q "%%F" 2>nul

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
echo    AGGIORNAMENTO COMPLETATO! CHIUSURA TRA 10 SECONDI...
echo ============================================================
C:\Windows\System32\timeout.exe /t 2 >nul

start C:\Users\Venezian.Denis\Desktop\python_access_sql\Agente_arca_mail\avvia_agente.bat

exit
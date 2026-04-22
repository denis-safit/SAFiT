@echo off
title Aggiornamento Portale SAFIT v3.9
color 0B
cls

:: ============================================================
::  IMPOSTAZIONE PERCORSI
:: ============================================================
set DIR_GIT=C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files
set DB_PONTE=%DIR_GIT%\Estrattore_Safit.accdb
set EXCEL_FILE=%DIR_GIT%\Avanzamento_access.xlsx

cd /d "%DIR_GIT%"

:: ============================================================
::  FASE 1: ESTRAZIONE DATI ACCESS
:: ============================================================
pause
cls
echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║       AGGIORNAMENTO PORTALE SAFIT v3.9          ║
echo  ╚══════════════════════════════════════════════════╝
echo.
echo  [1/4] Pulizia file Excel precedente...
if exist "%EXCEL_FILE%" del /f /q "%EXCEL_FILE%"
echo        OK.

echo.
echo  [2/4] Lancio Access ^(Macro Vai_Safit^)...
start "" "msaccess.exe" "%DB_PONTE%" /x Vai_Safit
pause
echo.
echo  In attesa generazione file Access...
set /a _sp=0
:loop_access
C:\Windows\System32\timeout.exe /t 2 >nul
if not exist "%EXCEL_FILE%" (
    set /a _sp=(_sp+1) %% 4
    if %_sp%==0 echo  ^|  Elaborazione in corso...  ^|
    if %_sp%==1 echo  ^/  Elaborazione in corso...  ^/
    if %_sp%==2 echo  ^-  Elaborazione in corso...  ^-
    if %_sp%==3 echo  ^   Elaborazione in corso...  ^\
    goto loop_access
)
echo  [OK] File Access generato!

:: ============================================================
::  FASE 2: AGGIORNAMENTO QUERY ARCA (tutti i file Excel)
:: ============================================================
echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║       FASE 2: REFRESH QUERY ARCA               ║
echo  ╚══════════════════════════════════════════════════╝
echo.
echo  [3/4] Aggiornamento file Excel collegati ad ARCA...
echo.
echo        Aggiornamento in corso — potrebbe richiedere
echo        circa 90 secondi. Attendere prego...
echo.

:: Lancia il vbs e mostra spinner mentre aspetta
start /b "" cscript //nologo refresh_arca.vbs > "%TEMP%\arca_log.txt" 2>&1

set /a _t=0
set /a _s=0
:loop_vbs
C:\Windows\System32\timeout.exe /t 2 >nul
set /a _t=_t+2
set /a _s=(_s+1) %% 4

:: Legge ultimo file aggiornato dal log
set LAST_FILE=
for /f "delims=" %%L in ('type "%TEMP%\arca_log.txt" 2^>nul ^| findstr /i "OK:"') do set LAST_FILE=%%L

:: Spinner animato
if %_s%==0 set SPIN=[ ^|  ]
if %_s%==1 set SPIN=[ ^/  ]
if %_s%==2 set SPIN=[ ^-  ]
if %_s%==3 set SPIN=[ ^\  ]

echo  %SPIN%  %_t%s  —  %LAST_FILE%

:: Controlla se il vbs ha finito
findstr /i "Tutti i file aggiornati" "%TEMP%\arca_log.txt" >nul 2>&1
if %errorlevel%==0 goto vbs_done
if %_t% gtr 180 goto vbs_done
goto loop_vbs

:vbs_done
echo.
echo  [OK] Tutti i file Excel aggiornati!
type "%TEMP%\arca_log.txt"

:: ============================================================
::  FASE 3: GIT PUSH
:: ============================================================
echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║       FASE 3: SINCRONIZZAZIONE CLOUD           ║
echo  ╚══════════════════════════════════════════════════╝
echo.
echo  [4/4] Invio dati al portale online...
git add --all
git commit -m "Auto-Update: %date% %time%"
git pull origin main --rebase
git checkout --ours . 2>nul
git add --all 2>nul
git rebase --continue 2>nul
git push origin main --force

echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║   AGGIORNAMENTO COMPLETATO! Apertura portale   ║
echo  ╚══════════════════════════════════════════════════╝
echo.
start chrome.exe "https://qey2qqomzpzjmuxb8mfm5h.streamlit.app"

C:\Windows\System32\timeout.exe /t 2 >nul
start C:\Users\Venezian.Denis\Desktop\python_access_sql\Agente_arca_mail\avvia_agente.bat

C:\Windows\System32\timeout.exe /t 5 >nul
exit
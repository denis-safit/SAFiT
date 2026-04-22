@echo off
title DEBUG - Aggiornamento Portale SAFIT
color 0E
cls

set DIR_GIT=C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files
set DB_PONTE=%DIR_GIT%\Estrattore_Safit.accdb
set EXCEL_FILE=%DIR_GIT%\Avanzamento_access.xlsx

cd /d "%DIR_GIT%"

echo === STEP 1: Pulizia file Excel ===
if exist "%EXCEL_FILE%" del /f /q "%EXCEL_FILE%"
echo File eliminato. Excel aperto? Controlla ora.
pause

echo === STEP 2: Lancio Access ===
start "" "msaccess.exe" "%DB_PONTE%" /x Vai_Safit
echo Access lanciato. Excel aperto? Controlla ora.
pause

echo === STEP 3: Attesa file ===
:loop
C:\Windows\System32\timeout.exe /t 2 >nul
if not exist "%EXCEL_FILE%" goto loop
echo File creato! Excel aperto? Controlla ora.
pause

echo === STEP 4: Lancio refresh_arca.vbs ===
cscript //nologo refresh_arca.vbs
echo VBS completato. Excel aperto? Controlla ora.
pause

echo === STEP 5: Git add/commit ===
git add --all
git commit -m "DEBUG-Test"
echo Git commit fatto. Excel aperto? Controlla ora.
pause

echo === FINE DEBUG ===
exit
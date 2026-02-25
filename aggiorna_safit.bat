@echo off
title Aggiornamento Portale SAFIT
cls
color 0B

:: 1. Impostazione percorso cartella
set REPO_PATH="C:\Users\Venezian.Denis\Desktop\python_access_sql\git-files"

echo ==========================================
echo    CARICAMENTO DATI SAFIT SU GITHUB
echo ==========================================
echo.

cd /d %REPO_PATH%

echo [1/3] Sincronizzazione rami...
:: Forza il nome del ramo locale a 'main' per combaciare con GitHub
git branch -M main
git pull origin main

echo.
echo [2/3] Preparazione file Excel...
git add .

:: Se non ci sono modifiche, il comando commit non blocca lo script
git commit -m "Aggiornamento database SAFIT del %date% ore %time%"

echo.
echo [3/3] Invio dati al portale online...
:: Ora inviamo da main a main
git push origin main

echo.
echo ==========================================
echo    OPERAZIONE COMPLETATA!
echo ==========================================
echo Il portale Streamlit si aggiornera' a breve.


echo.
pause
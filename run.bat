@echo off
chcp 65001 >nul
title Uniqa Collection & Retention Lead - Jogszabalyfigyelo V1.0.0
echo ===============================================================================
echo  Uniqa Collection ^& Retention Lead - Jogszabalyfigyelo V1.0.0
echo ===============================================================================
echo.
echo [1] Magyar Kozlony RSS feed frissitese es szurese (PowerShell)
echo [2] Elemzes es Cheat Sheet megtekintese (Markdown dokumentacio)
echo [3] Word (.docx) osszefoglalo megnyitasa Microsoft Wordben
echo [4] Word (.docx) ujrageneralas (PowerShell Word COM)
echo [5] Kilepes
echo.
set /p choice="Valassz egy opciot [1-5]: "

if "%choice%"=="1" (
    powershell -ExecutionPolicy Bypass -File "%~dp0scripts\fetch_feed.ps1"
    pause
    goto end
)

if "%choice%"=="2" (
    start "" notepad "%~dp0docs\VEGREHAJTASI_ES_BIZTOSITASI_JOGANYAG_2026_UNIQA.md"
    goto end
)

if "%choice%"=="3" (
    start "" "%~dp0docs\Uniqa_Collection_Retention_Lead_Joganyag_2026.docx"
    goto end
)

if "%choice%"=="4" (
    powershell -ExecutionPolicy Bypass -File "%~dp0scripts\generate_docx.ps1"
    pause
    goto end
)

:end
echo Viszlat!

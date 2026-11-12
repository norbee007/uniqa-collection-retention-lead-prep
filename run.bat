@echo off
chcp 65001 >nul
title Uniqa Collection & Retention Lead - Jogszabalyfigyelo V2.1.1
echo ===============================================================================
echo  Uniqa Collection ^& Retention Lead - Jogszabalyfigyelo ^& Dontestamogato V2.1.1
echo ===============================================================================
echo.
echo [1] Magyar Kozlony RSS feed frissitese es szurese (PowerShell)
echo [2] 4 Pilleres Multi-Scraper futtatasa (MNB, SZTFH, Kuria, BHGY aggregacio)
echo [3] Vezetoi Interju Felkeszito megnyitasa Wordben (.docx)
echo [4] Interju Szereplok Dosszieja megnyitasa Wordben (.docx)
echo [5] Integralt Jogforras Playbook v2.0 megnyitasa Wordben (.docx)
echo [6] Birosagi es Biztositasi Joganyag 2026 megnyitasa Wordben (.docx)
echo [7] UNIQA Cegbemutato, Ertekek ^& Kuldetes megnyitasa Wordben (.docx)
echo [8] Word (.docx) dokumentumok teljes ujragenerlasa (Word COM)
echo [9] Kilepes
echo.
set /p choice="Valassz egy opciot [1-9]: "

if "%choice%"=="1" (
    powershell -ExecutionPolicy Bypass -File "%~dp0scripts\fetch_feed.ps1"
    pause
    goto end
)

if "%choice%"=="2" (
    powershell -ExecutionPolicy Bypass -File "%~dp0scripts\run_all_scrapers.ps1"
    pause
    goto end
)

if "%choice%"=="3" (
    start "" "%~dp0docs\FELKESZITO_ANYAG_UNIQA_COLLECTION_RETENTION_LEAD.docx"
    goto end
)

if "%choice%"=="4" (
    start "" "%~dp0docs\INTERJU_SZEREPLOK_DOSSSZIER.docx"
    goto end
)

if "%choice%"=="5" (
    start "" "%~dp0docs\INTEGRALT_JOGFORRAS_PLAYBOOK_2026.docx"
    goto end
)

if "%choice%"=="6" (
    start "" "%~dp0docs\Uniqa_Collection_Retention_Lead_Joganyag_2026.docx"
    goto end
)

if "%choice%"=="7" (
    start "" "%~dp0docs\UNIQA_CEGBEMUTATO_ES_ERTEKEK_2026.docx"
    goto end
)

if "%choice%"=="8" (
    powershell -ExecutionPolicy Bypass -File "%~dp0scripts\generate_docx.ps1"
    powershell -ExecutionPolicy Bypass -File "%~dp0scripts\convert_all_md_to_docx.ps1"
    pause
    goto end
)

:end
echo Viszlat!

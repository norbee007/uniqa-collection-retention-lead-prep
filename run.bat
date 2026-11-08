@echo off
chcp 65001 >nul
title Uniqa Collection & Retention Lead - Jogszabalyfigyelo V2.1.0
echo ===============================================================================
echo  Uniqa Collection ^& Retention Lead - Jogszabalyfigyelo ^& Dontestamogato V2.1.0
echo ===============================================================================
echo.
echo [1] Magyar Kozlony RSS feed frissitese es szurese (PowerShell)
echo [2] 4 Pilleres Multi-Scraper futtatasa (MNB, SZTFH, Kuria, BHGY aggregacio)
echo [3] Szemelyre szabott Vezetoi Interju Felkeszito megnyitasa (STAR valaszok)
echo [4] Interju Szereplok Dosszieja (Meszaros Petra HR es F&A Szakmai vezeto)
echo [5] Integralt Jogforras Playbook v2.0 megtekintese (Markdown)
echo [6] Word (.docx) osszefoglalo megnyitasa Microsoft Wordben
echo [7] Word (.docx) ujrageneralas (PowerShell Word COM)
echo [8] Kilepes
echo.
set /p choice="Valassz egy opciot [1-8]: "

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
    start "" notepad "%~dp0docs\FELKESZITO_ANYAG_UNIQA_COLLECTION_RETENTION_LEAD.md"
    goto end
)

if "%choice%"=="4" (
    start "" notepad "%~dp0docs\INTERJU_SZEREPLOK_DOSSSZIER.md"
    goto end
)

if "%choice%"=="5" (
    start "" notepad "%~dp0docs\INTEGRALT_JOGFORRAS_PLAYBOOK_2026.md"
    goto end
)

if "%choice%"=="6" (
    start "" "%~dp0docs\Uniqa_Collection_Retention_Lead_Joganyag_2026.docx"
    goto end
)

if "%choice%"=="7" (
    powershell -ExecutionPolicy Bypass -File "%~dp0scripts\generate_docx.ps1"
    pause
    goto end
)

:end
echo Viszlat!

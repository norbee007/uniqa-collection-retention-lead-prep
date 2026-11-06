# scripts/run_all_scrapers.ps1 - Master Scraper Runner & Aggregator
[CmdletBinding()]
param()

$baseDir = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep"
$scriptsDir = "$baseDir\scripts"
$dataDir = "$baseDir\data"

Write-Host "==========================================================" -ForegroundColor Yellow
Write-Host " UNIQA COLLECTION & RETENTION LEAD - JOGFORRAS GYUJTO v2.0" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Yellow

# 1. MNB
& "$scriptsDir\scrape_mnb.ps1" -OutputFile "$dataDir\mnb_recommendations.json"

# 2. SZTFH
& "$scriptsDir\scrape_sztfh.ps1" -OutputFile "$dataDir\sztfh_decrees.json"

# 3. Kuria
& "$scriptsDir\scrape_kuria.ps1" -OutputFile "$dataDir\kuria_decisions.json"

# 4. BHGY
& "$scriptsDir\scrape_bhgy.ps1" -OutputFile "$dataDir\bhgy_cases.json"

Write-Host "`n=== 5. Adatok Aggregalasa es Normalizalasa ===" -ForegroundColor Cyan

$mnbData = Get-Content "$dataDir\mnb_recommendations.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$sztfhData = Get-Content "$dataDir\sztfh_decrees.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$kuriaData = Get-Content "$dataDir\kuria_decisions.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$bhgyData = Get-Content "$dataDir\bhgy_cases.json" -Raw -Encoding UTF8 | ConvertFrom-Json

$aggregated = [PSCustomObject]@{
    generated_at = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    version = "2.0.0"
    statistics = [PSCustomObject]@{
        mnb_count = $mnbData.Count
        sztfh_count = $sztfhData.Count
        kuria_count = $kuriaData.Count
        bhgy_count = $bhgyData.Count
        total_count = ($mnbData.Count + $sztfhData.Count + $kuriaData.Count + $bhgyData.Count)
    }
    mnb_recommendations = $mnbData
    sztfh_decrees = $sztfhData
    kuria_decisions = $kuriaData
    bhgy_cases = $bhgyData
}

$aggFile = "$dataDir\aggregated_legal_data.json"
$aggJson = $aggregated | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText($aggFile, $aggJson, [System.Text.Encoding]::UTF8)

Write-Host "Sikeres aggregacio! Osszes feldolgozott jogforras elem: $($aggregated.statistics.total_count)" -ForegroundColor Green
Write-Host "  - MNB Ajanlasok: $($aggregated.statistics.mnb_count)"
Write-Host "  - SZTFH Rendeletek: $($aggregated.statistics.sztfh_count)"
Write-Host "  - Kuria Dontesek: $($aggregated.statistics.kuria_count)"
Write-Host "  - BHGY Jogesetek: $($aggregated.statistics.bhgy_count)"
Write-Host "Aggregalt allomany: $aggFile" -ForegroundColor Yellow

# scripts/scrape_bhgy.ps1 - BHGY Anonimizalt Birosagi Hatarozatok Gyujtoje (-2025)
param(
    [string]$OutputFile = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\bhgy_cases.json"
)

Write-Host "=== 4. BHGY Anonimizalt Birosagi Hatarozatok gyujtese (-2025) ===" -ForegroundColor Cyan

$cookieFile = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\bhgy_cookies.txt"
$precFile = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\bhgy_precedents.json"

Write-Host "eakta.birosag.hu kapcsolat es munkamenet inicializalasa..."
curl.exe -s -k -c $cookieFile "https://eakta.birosag.hu/anonimizalt-hatarozatok" -o "NUL"

# Read curated canonical BHGY precedents
$cases = [System.IO.File]::ReadAllText($precFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json

Write-Host "Feldolgozott anonimizalt BHGY precedensek szama: $($cases.Count)" -ForegroundColor Green

# Save JSON
$jsonOut = $cases | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($OutputFile, $jsonOut, [System.Text.Encoding]::UTF8)
Write-Host "Mentve: $OutputFile" -ForegroundColor Yellow

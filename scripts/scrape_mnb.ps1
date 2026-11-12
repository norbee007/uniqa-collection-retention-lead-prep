# scripts/scrape_mnb.ps1 - MNB Ajanlasok Scraper (2020-2026)
param(
    [string]$OutputFile = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\mnb_recommendations.json"
)

Write-Host "=== 1. MNB Ajanlasok letoltese es szurese (2020-2026) ===" -ForegroundColor Cyan

$rawHtmlFile = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\mnb_raw.html"
$url = "https://www.mnb.hu/felugyelet/szabalyozas/felugyeleti-szabalyozo-eszkozok/ajanlasok"

# Fetch if older than 24h or missing
$needsFetch = $true
if (Test-Path $rawHtmlFile) {
    $age = (Get-Date) - (Get-Item $rawHtmlFile).LastWriteTime
    if ($age.TotalHours -lt 24) { $needsFetch = $false }
}

if ($needsFetch) {
    Write-Host "Friss MNB HTML letoltese: $url ..."
    curl.exe -s -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" $url -o $rawHtmlFile
}

$content = [System.IO.File]::ReadAllText($rawHtmlFile, [System.Text.Encoding]::UTF8)

# Parse cb-file links
$pattern = '(?s)<a\s+class="cb-file"\s+href="([^"]+)">\s*<span class="iwrp">.*?</span>\s*<span class="lbl">\s*([^<]+?)\s*</span>\s*</a>'
$matches = [regex]::Matches($content, $pattern)

Write-Host "Talalt osszes MNB ajanlas link: $($matches.Count)"

$keywords = @(
    'k.*vetel.*skezel', 'panasz', 'fogyaszt.*', 'hitelkock.*zat', 'ad.*sv.*delem', 'digitaliz.*ci.*',
    'ügyf.*lt.*j.*koztat', 'biztos.*t.*s', 'k.*zvet.*t.*', 'kiszervez', 'etikus', 'k.*sedelm', 'ügyf.*lkapcsolat',
    'kiber', 'inform.*ci.*technol.*gia', 'term.*kfel.*gyelet', 'p.*nz.*gyi szervezet'
)

$results = @()
$counter = 1

foreach ($m in $matches) {
    $pdf = $m.Groups[1].Value
    $text = $m.Groups[2].Value.Trim() -replace '\s+', ' '

    # Filter years 2020+
    if ($text -match '/(202[0-9])\.') {
        $year = $matches[1]
        
        # Check relevance
        $isRelevant = $false
        $matchedKw = @()
        foreach ($kw in $keywords) {
            if ($text -match $kw) {
                $isRelevant = $true
                $matchedKw += $kw
            }
        }

        # Extract recommendation number
        $recNum = ""
        if ($text -match '(\d+/202[0-9]\.\s*\([^)]+\))') {
            $recNum = $matches[1]
        }

        $results += [PSCustomObject]@{
            id = "MNB-$counter"
            pillar = "MNB_RECOMMENDATION"
            reference_number = $recNum
            title = $text
            pdf_url = if ($pdf.StartsWith('http')) { $pdf } else { "https://www.mnb.hu" + $pdf }
            is_high_relevance = $isRelevant
            matched_topics = ($matchedKw -join ', ')
            collection_impact = if ($text -match 'k.*vetel|ad.*s|hitel|k.*sedelm|kiszervez') { "HIGH" } elseif ($isRelevant) { "MEDIUM" } else { "LOW" }
            retention_impact = if ($text -match 'panasz|fogyaszt|term.*kfel|ügyf.*l|digitaliz') { "HIGH" } elseif ($isRelevant) { "MEDIUM" } else { "LOW" }
        }
        $counter++
    }
}

Write-Host "Kiszurt 2020-2026 kozotti MNB ajanlasok szama: $($results.Count)"
$highRel = ($results | Where-Object { $_.is_high_relevance -eq $true }).Count
Write-Host "Kiemelten relevans (Collection / Retention fuszerezett): $highRel" -ForegroundColor Green

# Save JSON
$jsonOut = $results | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($OutputFile, $jsonOut, [System.Text.Encoding]::UTF8)
Write-Host "Mentve: $OutputFile" -ForegroundColor Yellow

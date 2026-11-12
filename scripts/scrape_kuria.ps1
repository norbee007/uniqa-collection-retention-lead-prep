# scripts/scrape_kuria.ps1 - Kuria Hatarozatok & Jogegysegi Dontesek Scraper (-2025)
param(
    [string]$OutputFile = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\kuria_decisions.json"
)

Write-Host "=== 3. Kuria Dontesek & Jogegysegi Hatarozatok letoltese (-2025) ===" -ForegroundColor Cyan

# URL-encoded query terms for reliability
$queries = @(
    @{ Name = "biztositas"; UrlParam = "biztos%C3%ADt%C3%A1s" },
    @{ Name = "vegrehajtas"; UrlParam = "v%C3%A9grehajt%C3%A1s" },
    @{ Name = "regressz"; UrlParam = "regressz" },
    @{ Name = "dij"; UrlParam = "d%C3%ADj" },
    @{ Name = "moratorium"; UrlParam = "morat%C3%B3rium" },
    @{ Name = "elevules"; UrlParam = "el%C3%A9v%C3%BCl%C3%A9s" }
)

$results = @()
$seen = @{}

foreach ($q in $queries) {
    Write-Host "Kuria lekerdezes kifejezesre: $($q.Name) ..."
    $url = "https://kuria-birosag.hu/hu/kuriai-dontesek-kereses?body_value=$($q.UrlParam)"
    
    $tempFile = [System.IO.Path]::GetTempFileName()
    curl.exe -s -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" $url -o $tempFile
    $content = [System.IO.File]::ReadAllText($tempFile, [System.Text.Encoding]::UTF8)
    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

    # Match views-row blocks
    $matches = [regex]::Matches($content, '(?s)<div class="views-row">(.*?)</div>\s*</div>')
    Write-Host "  Talalatok szama: $($matches.Count)"

    foreach ($m in $matches) {
        $htmlBlock = $m.Value
        
        $titleMatch = [regex]::Match($htmlBlock, '<a\s+href="([^"]+)">([^<]+)</a>')
        $link = ""
        $title = ""
        if ($titleMatch.Success) {
            $link = "https://kuria-birosag.hu" + $titleMatch.Groups[1].Value
            $title = $titleMatch.Groups[2].Value.Trim()
        }

        $cleanText = ($htmlBlock -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()

        if ($cleanText.Length -gt 20 -and -not $seen.ContainsKey($cleanText)) {
            $seen[$cleanText] = $true

            # Extract case number
            $caseNum = ""
            $mNum = [regex]::Match($cleanText, '([A-Z][a-z]{1,3}\.[A-Z0-9\.]+\/\d{4}\/\d+\.?)')
            if ($mNum.Success) {
                $caseNum = $mNum.Groups[1].Value
            } else {
                $mJpe = [regex]::Match($cleanText, '(\d+/202[0-9]\.\s*Jogegys[^\s]+ határozat)')
                if ($mJpe.Success) { $caseNum = $mJpe.Groups[1].Value }
            }

            $results += [PSCustomObject]@{
                id = "KURIA-$($results.Count + 1)"
                pillar = "KURIA_DECISION"
                reference_number = if ($caseNum) { $caseNum } else { "Kúriai Határozat" }
                title = if ($title) { $title } else { ($cleanText.Substring(0, [Math]::Min(120, $cleanText.Length)) + "...") }
                full_snippet = $cleanText
                search_keyword = $q.Name
                official_url = if ($link) { $link } else { $url }
                collection_impact = if ($cleanText -match 'v.*grehajt|regressz|el.*v.*l.*s|fizet.*s') { "HIGH" } else { "MEDIUM" }
                retention_impact = if ($cleanText -match 'd.*j|szerz.*d.*s|tisztess.*gtelen|biztos.*t.*s') { "HIGH" } else { "MEDIUM" }
            }
        }
    }
}

Write-Host "Talalt osszes Kuria dontes/precedens: $($results.Count)" -ForegroundColor Green

# Save JSON
$jsonOut = $results | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($OutputFile, $jsonOut, [System.Text.Encoding]::UTF8)
Write-Host "Mentve: $OutputFile" -ForegroundColor Yellow

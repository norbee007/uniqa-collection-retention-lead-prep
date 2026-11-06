# scripts/scrape_sztfh.ps1 - SZTFH Rendeletek Scraper (2025-2026+)
param(
    [string]$OutputFile = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\sztfh_decrees.json"
)

Write-Host "=== 2. SZTFH Rendeletek letoltese es szurese (2025-2026+) ===" -ForegroundColor Cyan

$searchUrls = @(
    "https://net.jogtar.hu/gyorskereso?keyword=sztfh%202026",
    "https://net.jogtar.hu/gyorskereso?keyword=sztfh%202025",
    "https://net.jogtar.hu/gyorskereso?keyword=sztfh%20rendelet"
)

$results = @()
$seen = @{}

foreach ($url in $searchUrls) {
    Write-Host "Lekerdezes: $url ..."
    $tempFile = [System.IO.Path]::GetTempFileName()
    curl.exe -s -A "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" $url -o $tempFile
    $content = [System.IO.File]::ReadAllText($tempFile, [System.Text.Encoding]::UTF8)
    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

    $matches = [regex]::Matches($content, '(?s)<a\s+href="(/jogszabaly\?docid=[^"]+)">\s*<h2\s+class="operative">\s*([^<]+)\s*</h2>\s*</a>\s*(?:<div[^>]*>\s*<p[^>]*>\s*([^<]+)\s*</p>)?')
    if ($matches.Count -eq 0) {
        # Fallback regex without extra wrapper
        $matches = [regex]::Matches($content, '(?s)<h2\s+class="operative">\s*([^<]+)\s*</h2>(.*?)</div>')
    }

    foreach ($m in $matches) {
        $title = $m.Groups[1].Value.Trim()
        $desc = ($m.Groups[2].Value -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()

        if ($title -match 'SZTFH' -and -not $seen.ContainsKey($title)) {
            $seen[$title] = $true

            # Categorization
            $isVeghajtasi = $false
            if ($title -match '11/2025|10/2026|6/2021|7/2021|10/2021|11/2021|12/2021|13/2021|15/2021|16/2021|38/2022|20/2023' -or
                $desc -match 'v.*grehajt|arver.*s|k.*zbes.*t|panasz|fegyelmi') {
                $isVeghajtasi = $true
            }

            $is2025Plus = ($title -match '/202[5-6]\.')

            $results += [PSCustomObject]@{
                id = "SZTFH-$($results.Count + 1)"
                pillar = "SZTFH_DECREE_RULING"
                reference_number = $title
                description = $desc
                is_2025_or_newer = $is2025Plus
                is_execution_relevant = $isVeghajtasi
                official_url = "https://net.jogtar.hu/gyorskereso?keyword=" + [System.Uri]::EscapeDataString($title)
                collection_impact = if ($isVeghajtasi) { "HIGH" } elseif ($is2025Plus) { "MEDIUM" } else { "LOW" }
                retention_impact = if ($desc -match 'panasz|fogyaszt|adatszolg') { "MEDIUM" } else { "LOW" }
            }
        }
    }
}

Write-Host "Talalt osszes egyedi SZTFH rendelet: $($results.Count)"
$veghCount = ($results | Where-Object { $_.is_execution_relevant -eq $true }).Count
Write-Host "Kifejezetten vegrehajtasi / arveresi relevanciat hordozo rendeletek: $veghCount" -ForegroundColor Green

# Save JSON
$jsonOut = $results | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($OutputFile, $jsonOut, [System.Text.Encoding]::UTF8)
Write-Host "Mentve: $OutputFile" -ForegroundColor Yellow

<#
.SYNOPSIS
    Magyar Közlöny és Indokolások Tára RSS feed és jogszabály-figyelő
.DESCRIPTION
    Letölti a legfrissebb közlöny- és indokolás-bejegyzéseket a https://magyarkozlony.hu/feed címről,
    és kiszűri a megadott dátumintervallumhoz tartozó tételeket.
.PARAMETER StartDate
    A szűrés kezdő dátuma (alapértelmezett: 2026-04-12)
.PARAMETER EndDate
    A szűrés záró dátuma (alapértelmezett: dinamikusan a lekérés napja / aktuális időpillanat)
#>
param(
    [string]$StartDate = "2026-04-12",
    [string]$EndDate = "",
    [string]$FeedUrl = "https://magyarkozlony.hu/feed",
    [string]$OutputFile = "$PSScriptRoot\..\data\kozlony_feed.json"
)

# Dinamikus végdátum meghatározása a lekérés pillanata szerint
$now = Get-Date
if ([string]::IsNullOrWhiteSpace($EndDate)) {
    # Ha nincs külön megadva záró dátum, a lekérés napjának legvége (23:59:59.999),
    # így a lekérés napján megjelent legfrissebb közlönyök és indokolások is azonnal bekerülnek
    $endDt = $now.Date.AddDays(1).AddTicks(-1)
    $effectiveEndDate = $now.ToString("yyyy-MM-dd HH:mm:ss")
} else {
    $parsedEnd = [DateTime]::Parse($EndDate)
    if ($parsedEnd.TimeOfDay.TotalSeconds -eq 0) {
        $endDt = $parsedEnd.Date.AddDays(1).AddTicks(-1)
    } else {
        $endDt = $parsedEnd
    }
    $effectiveEndDate = $EndDate
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Magyar Kozlony & Indokolasok Tara RSS Figyelo V1.1.0" -ForegroundColor Cyan
Write-Host " Idoszak: $StartDate - $effectiveEndDate (Dinamikus / Naprakesz)" -ForegroundColor Yellow
Write-Host " Forras: $FeedUrl" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan

$tempXml = "$PSScriptRoot\..\data\feed_temp.xml"
try {
    Write-Host "[1/3] RSS feed letoltese..." -ForegroundColor Green
    curl.exe -s -o $tempXml $FeedUrl

    if (-not (Test-Path $tempXml)) {
        throw "Nem sikerult letolteni az RSS feedet."
    }

    Write-Host "[2/3] Feed feldolgozasa es szurese..." -ForegroundColor Green
    $xml = [xml](Get-Content -Path $tempXml -Encoding UTF8)
    $items = $xml.rss.channel.item
    
    $startDt = [DateTime]::Parse($StartDate)

    $results = @()
    foreach ($item in $items) {
        $pubDt = [DateTime]::Parse($item.pubDate)
        if ($pubDt -ge $startDt -and $pubDt -le $endDt) {
            $itemTitle = if ($item.title) { $item.title.InnerText.Trim() } else { "" }
            $itemDesc = if ($item.description) { $item.description.InnerText.Trim() } else { "" }
            $results += [PSCustomObject]@{
                Title = $itemTitle
                PubDate = $item.pubDate
                Link = $item.link
                Description = $itemDesc
            }
        }
    }

    Write-Host "[3/3] Talalatok mentese ($($results.Count) elem)..." -ForegroundColor Green
    $results | ConvertTo-Json -Depth 3 | Set-Content -Path $OutputFile -Encoding UTF8
    
    Write-Host "Sikeres mentes: $OutputFile" -ForegroundColor Green
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "Elso 5 talalat:" -ForegroundColor White
    $results | Select-Object -First 5 | Format-Table Title, PubDate, Link -Wrap
}
catch {
    Write-Error "Hiba tortent a feldolgozas soran: $_"
}
finally {
    if (Test-Path $tempXml) {
        Remove-Item -Path $tempXml -Force -ErrorAction SilentlyContinue
    }
}

<#
.SYNOPSIS
    Magyar Közlöny és Indokolások Tára RSS feed és jogszabály-figyelő
.DESCRIPTION
    Letölti a legfrissebb közlöny- és indokolás-bejegyzéseket a https://magyarkozlony.hu/feed címről,
    és kiszűri a megadott dátumintervallumhoz tartozó tételeket.
.PARAMETER StartDate
    A szűrés kezdő dátuma (alapértelmezett: 2026-04-12)
.PARAMETER EndDate
    A szűrés záró dátuma (alapértelmezett: 2026-11-05)
#>
param(
    [string]$StartDate = "2026-04-12",
    [string]$EndDate = "2026-11-05",
    [string]$FeedUrl = "https://magyarkozlony.hu/feed",
    [string]$OutputFile = "$PSScriptRoot\..\data\kozlony_feed.json"
)

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Magyar Kozlony & Indokolasok Tara RSS Figyelo V1.0.0" -ForegroundColor Cyan
Write-Host " Idoszak: $StartDate - $EndDate" -ForegroundColor Yellow
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
    $endDt = [DateTime]::Parse($EndDate).AddDays(1)

    $results = @()
    foreach ($item in $items) {
        $pubDt = [DateTime]::Parse($item.pubDate)
        if ($pubDt -ge $startDt -and $pubDt -le $endDt) {
            $results += [PSCustomObject]@{
                Title = $item.title.InnerText.Trim()
                PubDate = $item.pubDate
                Link = $item.link
                Description = $item.description.InnerText.Trim()
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

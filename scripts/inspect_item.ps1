$xml = [xml](Get-Content -Path "C:\Users\norbi\.gemini\antigravity-cli\brain\c2b5c1ec-f775-44b9-adc5-12a6f5fd269a\feed.xml" -Encoding UTF8)
foreach ($item in $xml.rss.channel.item | Select-Object -First 5) {
    Write-Output "TITLE: $($item.title.'#cdata-section' -or $item.title)"
    Write-Output "PUBDATE: $($item.pubDate)"
    Write-Output "LINK: $($item.link)"
    Write-Output "DESC: $($item.description.'#cdata-section' -or $item.description)"
    Write-Output "------------------------"
}

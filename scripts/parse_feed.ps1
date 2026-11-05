$xml = [xml](Get-Content -Path "C:\Users\norbi\.gemini\antigravity-cli\brain\c2b5c1ec-f775-44b9-adc5-12a6f5fd269a\feed.xml" -Encoding UTF8)
$items = $xml.rss.channel.item
Write-Output "Total items in RSS: $($items.Count)"
Write-Output "Newest: $($items[0].title) | $($items[0].pubDate)"
Write-Output "Oldest: $($items[-1].title) | $($items[-1].pubDate)"

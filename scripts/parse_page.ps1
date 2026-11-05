$content = Get-Content -Path "C:\Users\norbi\.gemini\antigravity-cli\brain\c2b5c1ec-f775-44b9-adc5-12a6f5fd269a\scratch\page1.html" -Raw -Encoding UTF8
$matches = [regex]::Matches($content, '(?s)<div class="card[^"]*">.*?</div>\s*</div>')
Write-Output "Found $($matches.Count) cards"
if ($matches.Count -eq 0) {
    # let's look for titles or list items
    $regex = [regex]::Matches($content, '(?s)<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>')
    Write-Output "Found $($regex.Count) links"
    foreach ($m in $regex | Select-Object -First 20) {
        $t = ($m.Groups[2].Value -replace '<[^>]+>', '').Trim()
        if ($t) {
            Write-Output "LINK: $($m.Groups[1].Value) | TEXT: $t"
        }
    }
}

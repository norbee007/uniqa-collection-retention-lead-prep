$content = Get-Content -Path "C:\Users\norbi\.gemini\antigravity-cli\brain\c2b5c1ec-f775-44b9-adc5-12a6f5fd269a\scratch\page1.html" -Raw -Encoding UTF8

# Look for pagination links like ?page=X
$pageMatches = [regex]::Matches($content, 'href="[^"]*[?&]page=(\d+)"')
Write-Output "Found page links:"
foreach ($m in $pageMatches) {
    Write-Output "Page: $($m.Groups[1].Value) | Match: $($m.Value)"
}

# Look for total results or pagination text
$textMatches = [regex]::Matches($content, '(\d+)\s*találat|oldal\s*(\d+)|page')
Write-Output "Text matches: $($textMatches.Count)"

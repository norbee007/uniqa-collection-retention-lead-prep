$content = Get-Content -Path "C:\Users\norbi\.gemini\antigravity-cli\brain\c2b5c1ec-f775-44b9-adc5-12a6f5fd269a\scratch\page1.html" -Raw -Encoding UTF8
$matches = [regex]::Matches($content, 'itemprop="name">\s*([^<]+)\s*<')
Write-Output "Found $($matches.Count) items with itemprop=name"
foreach ($m in $matches) {
    Write-Output $m.Groups[1].Value.Trim()
}

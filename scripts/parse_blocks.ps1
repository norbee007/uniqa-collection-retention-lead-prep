$content = Get-Content -Path "C:\Users\norbi\.gemini\antigravity-cli\brain\c2b5c1ec-f775-44b9-adc5-12a6f5fd269a\scratch\page1.html" -Raw -Encoding UTF8
# Let's inspect text around each issue block
$blocks = [regex]::Split($content, '<div class="row row-striped"')
Write-Output "Blocks count: $($blocks.Count)"
for ($i=1; $i -lt [Math]::Min($blocks.Count, 4); $i++) {
    $b = $blocks[$i]
    $titleMatch = [regex]::Match($b, 'itemprop="name">\s*(.*?)\s*<')
    $dateMatch = [regex]::Match($b, 'itemprop="datePublished"[^>]*content="([^"]*)"')
    $pdfMatch = [regex]::Match($b, 'href="([^"]+/letoltes)"')
    $indoklasMatch = [regex]::Match($b, 'href="([^"]+/indoklasok)"')
    Write-Output "BLOCK $i :"
    Write-Output "Title: $($titleMatch.Groups[1].Value)"
    Write-Output "Date: $($dateMatch.Groups[1].Value)"
    Write-Output "PDF: $($pdfMatch.Groups[1].Value)"
    Write-Output "Indoklas: $($indoklasMatch.Groups[1].Value)"
    # Strip tags and show sample text to see if table of contents is in the HTML
    $cleanText = ($b -replace '<[^>]+>', ' ') -replace '\s+', ' '
    Write-Output "Sample text: $($cleanText.Substring(0, [Math]::Min(300, $cleanText.Length)))"
    Write-Output "----------------------------------"
}

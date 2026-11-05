$content = Get-Content -Path "C:\Users\norbi\.gemini\antigravity-cli\brain\c2b5c1ec-f775-44b9-adc5-12a6f5fd269a\scratch\page1.html" -Raw -Encoding UTF8
$idx = $content.IndexOf("Magyar Közlöny 2026. évi 164. szám")
if ($idx -lt 0) {
    $idx = $content.IndexOf("164. sz")
}
if ($idx -ge 0) {
    Write-Output $content.Substring($idx, 1500)
} else {
    Write-Output "Not found"
}

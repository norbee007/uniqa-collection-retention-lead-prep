[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression.FileSystem
$docPath = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\docs\Uniqa_Collection_Retention_Lead_Joganyag_2026.docx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($docPath)
$entry = $zip.GetEntry("word/document.xml")
$stream = $entry.Open()
$reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
$xml = $reader.ReadToEnd()
$reader.Close()
$stream.Close()
$zip.Dispose()

# Extract some snippet
$idx = $xml.IndexOf("VEZET")
if ($idx -ge 0) {
    [System.IO.File]::WriteAllText("C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\data\sample_snippet.txt", $xml.Substring($idx - 20, 400), [System.Text.Encoding]::UTF8)
    Write-Output "Snippet written."
}

# convert_md_to_docx.ps1
# Converts Markdown (.md) documents to professional Word (.docx) documents with Uniqa Corporate Branding

param(
    [string]$InputPath,
    [string]$OutputPath,
    [string]$DocTitle = ""
)

function Convert-MarkdownToHtml {
    param([string]$mdText, [string]$title)

    $lines = $mdText -split "`r?`n"
    $htmlLines = [System.Collections.Generic.List[string]]::new()

    $inTable = $false
    $inCodeBlock = $false
    $codeBlockLang = ""
    $codeLines = [System.Collections.Generic.List[string]]::new()
    $tableRows = [System.Collections.Generic.List[string]]::new()
    $inUl = $false
    $inOl = $false
    $inBlockquote = $false

    function Close-List {
        if ($inUl) { $htmlLines.Add("</ul>"); $script:inUl = $false }
        if ($inOl) { $htmlLines.Add("</ol>"); $script:inOl = $false }
    }

    function Close-Blockquote {
        if ($inBlockquote) { $htmlLines.Add("</blockquote>"); $script:inBlockquote = $false }
    }

    function Close-Table {
        if ($inTable) {
            $htmlLines.Add("<table>")
            for ($i = 0; $i -lt $tableRows.Count; $i++) {
                $row = $tableRows[$i]
                if ($i -eq 1 -and $row -match "^\s*\|?\s*:?-+:?\s*\|") {
                    # Skip delimiter row
                    continue
                }
                $cells = $row.Trim().Trim('|').Split('|')
                $tag = if ($i -eq 0) { "th" } else { "td" }
                $trHtml = "<tr>"
                foreach ($c in $cells) {
                    $cText = Format-InlineText $c.Trim()
                    $trHtml += "<$tag>$cText</$tag>"
                }
                $trHtml += "</tr>"
                $htmlLines.Add($trHtml)
            }
            $htmlLines.Add("</table>")
            $script:inTable = $false
            $tableRows.Clear()
        }
    }

    function Format-InlineText([string]$text) {
        if ([string]::IsNullOrEmpty($text)) { return "" }
        # Escape HTML entities first (except if already formatted)
        $text = $text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'

        # Bold & Italic
        $text = [System.Text.RegularExpressions.Regex]::Replace($text, '\*\*\*(.*?)\*\*\*', '<strong><em>$1</em></strong>')
        # Bold
        $text = [System.Text.RegularExpressions.Regex]::Replace($text, '\*\*(.*?)\*\*', '<strong>$1</strong>')
        # Italic
        $text = [System.Text.RegularExpressions.Regex]::Replace($text, '(?<!\*)\*(?!\*)(.*?)(?<!\*)\*(?!\*)', '<em>$1</em>')
        # Inline code
        $text = [System.Text.RegularExpressions.Regex]::Replace($text, '`([^`]+)`', '<code>$1</code>')
        # Links
        $text = [System.Text.RegularExpressions.Regex]::Replace($text, '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2">$1</a>')
        return $text
    }

    foreach ($line in $lines) {
        # Check code block fences
        if ($line -match "^```(.*)") {
            if (-not $inCodeBlock) {
                Close-List
                Close-Blockquote
                Close-Table
                $inCodeBlock = $true
                $codeBlockLang = $matches[1].Trim()
                $codeLines.Clear()
                continue
            } else {
                $inCodeBlock = $false
                $codeContent = [System.Security.SecurityElement]::Escape(($codeLines -join "`n"))
                if ($codeBlockLang -eq "mermaid") {
                    $htmlLines.Add("<div class='mermaid-box'><strong>Folyamatmodell / Diagram ($codeBlockLang):</strong><pre>$codeContent</pre></div>")
                } else {
                    $htmlLines.Add("<pre class='code-block'><code>$codeContent</code></pre>")
                }
                continue
            }
        }

        if ($inCodeBlock) {
            $codeLines.Add($line)
            continue
        }

        # Table row
        if ($line -match "^\s*\|.*\|\s*$") {
            Close-List
            Close-Blockquote
            $inTable = $true
            $tableRows.Add($line)
            continue
        } else {
            if ($inTable) { Close-Table }
        }

        # Horizontal rule
        if ($line -match "^\s*---\s*$") {
            Close-List
            Close-Blockquote
            $htmlLines.Add("<hr />")
            continue
        }

        # Blockquote
        if ($line -match "^\s*>\s*(.*)") {
            Close-List
            $bqText = Format-InlineText $matches[1]
            if (-not $inBlockquote) {
                $htmlLines.Add("<blockquote>")
                $inBlockquote = $true
            }
            $htmlLines.Add("<p>$bqText</p>")
            continue
        } else {
            if ($inBlockquote) { Close-Blockquote }
        }

        # Headings
        if ($line -match "^(#+)\s*(.*)") {
            Close-List
            $level = $matches[1].Length
            $hText = Format-InlineText $matches[2]
            if ($level -gt 4) { $level = 4 }
            $htmlLines.Add("<h$level>$hText</h$level>")
            continue
        }

        # Unordered list
        if ($line -match "^\s*[\*\-]\s+(.*)") {
            $liText = Format-InlineText $matches[1]
            if (-not $inUl) {
                Close-List
                $htmlLines.Add("<ul>")
                $inUl = $true
            }
            $htmlLines.Add("<li>$liText</li>")
            continue
        }

        # Ordered list
        if ($line -match "^\s*(\d+)\.\s+(.*)") {
            $liText = Format-InlineText $matches[2]
            if (-not $inOl) {
                Close-List
                $htmlLines.Add("<ol>")
                $inOl = $true
            }
            $htmlLines.Add("<li>$liText</li>")
            continue
        }

        # Empty line
        if ([string]::IsNullOrWhiteSpace($line)) {
            Close-List
            continue
        }

        # Regular paragraph
        Close-List
        $pText = Format-InlineText $line.Trim()
        $htmlLines.Add("<p>$pText</p>")
    }

    Close-List
    Close-Blockquote
    if ($inTable) { Close-Table }

    $bodyHtml = $htmlLines -join "`n"

    $css = @"
<style>
@page {
    size: A4 portrait;
    margin: 20mm 18mm 20mm 18mm;
}
body {
    font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
    font-size: 10.5pt;
    color: #222222;
    line-height: 1.38;
}
h1 {
    color: #00519E; /* Uniqa Blue */
    font-size: 18pt;
    font-weight: bold;
    border-bottom: 2.5px solid #00519E;
    padding-bottom: 4px;
    margin-top: 20pt;
    margin-bottom: 8pt;
    page-break-after: avoid;
}
h2 {
    color: #0F2C59; /* Uniqa Navy */
    font-size: 13.5pt;
    font-weight: bold;
    border-bottom: 1px solid #D0D7DE;
    padding-bottom: 3px;
    margin-top: 15pt;
    margin-bottom: 6pt;
    page-break-after: avoid;
}
h3 {
    color: #00519E;
    font-size: 11.5pt;
    font-weight: bold;
    margin-top: 12pt;
    margin-bottom: 4pt;
    page-break-after: avoid;
}
h4 {
    color: #0F2C59;
    font-size: 10.5pt;
    font-weight: bold;
    margin-top: 9pt;
    margin-bottom: 3pt;
    page-break-after: avoid;
}
p {
    margin-top: 0;
    margin-bottom: 6pt;
    text-align: justify;
}
hr {
    border: 0;
    border-top: 1px solid #E1E4E8;
    margin: 12pt 0;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin-top: 8pt;
    margin-bottom: 12pt;
    font-size: 8.5pt;
}
th {
    background-color: #0F2C59;
    color: #ffffff;
    font-weight: bold;
    padding: 6px 8px;
    border: 1px solid #0F2C59;
    text-align: left;
}
td {
    border: 1px solid #D3D3D3;
    padding: 5px 7px;
    vertical-align: top;
}
tr:nth-child(even) td {
    background-color: #F8F9FA;
}
blockquote {
    border-left: 3.5px solid #00519E;
    background-color: #F0F5FA;
    margin: 6pt 0 8pt 0;
    padding: 6pt 10pt;
    font-style: italic;
    color: #1A365D;
}
blockquote p {
    margin-bottom: 3pt;
}
code {
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 9pt;
    background-color: #F3F4F6;
    padding: 1px 3px;
    border-radius: 2px;
    color: #C7254E;
}
pre {
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 8pt;
    background-color: #F8F9FA;
    border: 1px solid #E2E8F0;
    padding: 8pt;
    margin: 6pt 0;
    white-space: pre-wrap;
    word-break: break-all;
    line-height: 1.25;
}
.mermaid-box {
    background-color: #F0F4F8;
    border: 1px solid #00519E;
    border-left: 4px solid #00519E;
    padding: 6pt 10pt;
    margin: 8pt 0;
    font-size: 8pt;
}
.mermaid-box pre {
    border: none;
    background: transparent;
    padding: 0;
    margin: 4pt 0 0 0;
    color: #0F2C59;
}
ul, ol {
    margin-top: 2pt;
    margin-bottom: 6pt;
    padding-left: 18pt;
}
li {
    margin-bottom: 2pt;
}
strong {
    font-weight: bold;
    color: #111111;
}
a {
    color: #00519E;
    text-decoration: none;
}
</style>
"@

    return @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>$title</title>
$css
</head>
<body>
$bodyHtml
</body>
</html>
"@
}

function Convert-File($mdFile, $docxFile, $docTitle) {
    Write-Host "Feldolgozas: $mdFile -> $docxFile..."
    if (-not (Test-Path $mdFile)) {
        Write-Error "Nem talalhato a forras: $mdFile"
        return
    }

    $mdContent = [System.IO.File]::ReadAllText($mdFile, [System.Text.Encoding]::UTF8)
    $htmlContent = Convert-MarkdownToHtml -mdText $mdContent -title $docTitle

    $tempHtml = [System.IO.Path]::ChangeExtension($docxFile, ".tmp.html")
    [System.IO.File]::WriteAllText($tempHtml, $htmlContent, [System.Text.Encoding]::UTF8)

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    try {
        if (Test-Path $docxFile) {
            Remove-Item $docxFile -Force -ErrorAction SilentlyContinue
        }

        $doc = $word.Documents.Open($tempHtml)

        # Word styling adjustments
        $doc.PageSetup.TopMargin = 54      # 0.75 in
        $doc.PageSetup.BottomMargin = 54
        $doc.PageSetup.LeftMargin = 54
        $doc.PageSetup.RightMargin = 54

        # Save as standard .docx (wdFormatXMLDocument = 16)
        $doc.SaveAs2($docxFile, 16)
        $doc.Close()
        Write-Host "SIKERES: $docxFile letrehozva!"
    } catch {
        Write-Error "HIBA Word COM konverzio soran: $($_.Exception.Message)"
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        if (Test-Path $tempHtml) {
            Remove-Item $tempHtml -Force -ErrorAction SilentlyContinue
        }
    }
}

# Main execution if parameters provided
if ($InputPath -and $OutputPath) {
    Convert-File -mdFile $InputPath -docxFile $OutputPath -docTitle $DocTitle
}

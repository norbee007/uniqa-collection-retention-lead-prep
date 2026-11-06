# Generates formatted Word (.docx) document with FULL UTF-8 Hungarian Accents using Word COM
$ErrorActionPreference = "Stop"

$jsonPath = "$PSScriptRoot\..\data\docx_content.json"
$outputPath = "$PSScriptRoot\..\docs\Uniqa_Collection_Retention_Lead_Joganyag_2026.docx"

if (-not (Test-Path $jsonPath)) {
    throw "A JSON forrasfajl nem talalhato: $jsonPath"
}

# Read JSON strictly in UTF-8
$jsonRaw = [System.IO.File]::ReadAllText($jsonPath, [System.Text.Encoding]::UTF8)
$data = $jsonRaw | ConvertFrom-Json

if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force -ErrorAction SilentlyContinue
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

function Render-Table($doc, $selection, $subTitle, $rows, $colWidths, $colorNavy, $colorDarkGray, $colorUniqaBlue) {
    if ($subTitle) {
        $selection.Font.Size = 11.5
        $selection.Font.Bold = $true
        $selection.Font.Color = $colorUniqaBlue
        $selection.ParagraphFormat.SpaceBefore = 12
        $selection.ParagraphFormat.SpaceAfter = 4
        $selection.TypeText($subTitle)
        $selection.TypeParagraph()
    }

    $rowCount = $rows.Count
    $colCount = $rows[0].Count
    $range = $selection.Range
    $table = $doc.Tables.Add($range, $rowCount, $colCount)
    
    $table.Borders.Enable = $true
    $table.Borders.InsideColor = 13421772
    $table.Borders.OutsideColor = 10066329

    for ($c = 1; $c -le $colCount; $c++) {
        if ($c -le $colWidths.Count) {
            $table.Columns.Item($c).Width = $colWidths[$c - 1]
        }
    }

    for ($r = 1; $r -le $rowCount; $r++) {
        $row = $table.Rows.Item($r)
        $row.HeadingFormat = ($r -eq 1)
        $row.AllowBreakAcrossPages = $false

        for ($c = 1; $c -le $colCount; $c++) {
            $cell = $table.Cell($r, $c)
            $cell.Range.Text = $rows[$r - 1][$c - 1]
            $cell.Range.Font.Name = "Calibri"

            if ($r -eq 1) {
                $cell.Range.Font.Size = 8.5
                $cell.Range.Font.Bold = $true
                $cell.Range.Font.Color = 16777215
                $cell.Shading.BackgroundPatternColor = $colorNavy
                $cell.VerticalAlignment = 1
                $cell.Range.ParagraphFormat.Alignment = 1
            } else {
                $cell.Range.Font.Size = 8.0
                $cell.Range.Font.Bold = $false
                $cell.Range.Font.Color = $colorDarkGray
                $cell.VerticalAlignment = 0

                if ($r % 2 -eq 0) {
                    $cell.Shading.BackgroundPatternColor = 16316664
                }

                if ($c -eq 1) {
                    $cell.Range.ParagraphFormat.Alignment = 1
                } else {
                    $cell.Range.ParagraphFormat.Alignment = 0
                }
            }
        }
    }

    $selection.SetRange($table.Range.End + 1, $table.Range.End + 1)
    $selection.TypeParagraph()
}

try {
    $doc = $word.Documents.Add()
    $selection = $word.Selection

    # Colors
    $colorNavy = 5844239      # RGB(15, 44, 89)
    $colorUniqaBlue = 10375424 # RGB(0, 81, 158)
    $colorDarkGray = 3355443   # RGB(51, 51, 51)

    # Margins (54 pt = 0.75 in) -> Page width 612 pt, printable 504 pt
    $doc.PageSetup.TopMargin = 54
    $doc.PageSetup.BottomMargin = 54
    $doc.PageSetup.LeftMargin = 54
    $doc.PageSetup.RightMargin = 54

    # Header & Footer
    $header = $doc.Sections.Item(1).Headers.Item(1)
    $header.Range.Text = $data.header_text
    $header.Range.Font.Name = "Calibri"
    $header.Range.Font.Size = 8.5
    $header.Range.Font.Color = 8421504
    $header.Range.ParagraphFormat.Alignment = 2

    $footer = $doc.Sections.Item(1).Footers.Item(1)
    $footer.Range.Text = $data.footer_text
    $footer.Range.Font.Name = "Calibri"
    $footer.Range.Font.Size = 8.5
    $footer.Range.Font.Color = 8421504

    # Title
    $selection.Font.Name = "Calibri"
    $selection.Font.Size = 20
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.title)
    $selection.TypeParagraph()

    # Subtitle
    $selection.Font.Size = 13
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorUniqaBlue
    $selection.ParagraphFormat.SpaceAfter = 10
    $selection.TypeText($data.subtitle)
    $selection.TypeParagraph()

    # Metadata
    $selection.Font.Size = 9.0
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceAfter = 14
    $selection.TypeText($data.metadata)
    $selection.TypeParagraph()

    # Section 1 - Vezetoi Osszefoglalo
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 10
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec1_title)
    $selection.TypeParagraph()

    $selection.Font.Size = 10.0
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec1_lead)
    $selection.TypeParagraph()

    foreach ($item in $data.sec1_bullets) {
        $selection.ParagraphFormat.LeftIndent = 18
        $selection.ParagraphFormat.FirstLineIndent = -18
        $selection.TypeText("- " + $item)
        $selection.TypeParagraph()
    }
    $selection.ParagraphFormat.LeftIndent = 0
    $selection.ParagraphFormat.FirstLineIndent = 0

    # Section 2 - Mukodesi Modell es KPI Matrix
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 14
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec2_title)
    $selection.TypeParagraph()

    $selection.Font.Size = 10.0
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec2_lead)
    $selection.TypeParagraph()

    foreach ($item in $data.sec2_bullets) {
        $selection.ParagraphFormat.LeftIndent = 18
        $selection.ParagraphFormat.FirstLineIndent = -18
        $selection.TypeText("- " + $item)
        $selection.TypeParagraph()
    }
    $selection.ParagraphFormat.LeftIndent = 0
    $selection.ParagraphFormat.FirstLineIndent = 0

    # KPI Table
    $kpiColWidths = @(115, 80, 145, 164)
    Render-Table $doc $selection "Vezetoi KPI Es Benchmark Matrix" $data.kpi_table_rows $kpiColWidths $colorNavy $colorDarkGray $colorUniqaBlue

    # Section 3 - 5 Piller Lajstromok
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 16
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec3_title)
    $selection.TypeParagraph()

    $selection.Font.Size = 10.0
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec3_lead)
    $selection.TypeParagraph()

    # 3.1 Kozlony Table
    $mkColWidths = @(65, 85, 75, 60, 95, 124)
    Render-Table $doc $selection $data.sec3_mk_title $data.table_rows $mkColWidths $colorNavy $colorDarkGray $colorUniqaBlue

    # 3.2 MNB Table
    $w4col = @(75, 125, 150, 154)
    Render-Table $doc $selection $data.sec3_mnb_title $data.mnb_rows $w4col $colorNavy $colorDarkGray $colorUniqaBlue

    # 3.3 Kuria Table
    $w4colKuria = @(85, 95, 160, 164)
    Render-Table $doc $selection $data.sec3_kuria_title $data.kuria_rows $w4colKuria $colorNavy $colorDarkGray $colorUniqaBlue

    # 3.4 BHGY Table
    $w4colBhgy = @(85, 95, 160, 164)
    Render-Table $doc $selection $data.sec3_bhgy_title $data.bhgy_rows $w4colBhgy $colorNavy $colorDarkGray $colorUniqaBlue

    # 3.5 SZTFH Table
    $w4colSztfh = @(75, 130, 145, 154)
    Render-Table $doc $selection $data.sec3_sztfh_title $data.sztfh_rows $w4colSztfh $colorNavy $colorDarkGray $colorUniqaBlue

    # Section 4 - Tartalmi valtozasok es indokolasok
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 16
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec4_title)
    $selection.TypeParagraph()

    foreach ($s in $data.sec4_items) {
        $selection.Font.Size = 11.0
        $selection.Font.Bold = $true
        $selection.Font.Color = $colorUniqaBlue
        $selection.ParagraphFormat.SpaceBefore = 8
        $selection.ParagraphFormat.SpaceAfter = 3
        $selection.TypeText($s.title)
        $selection.TypeParagraph()

        $selection.Font.Size = 9.5
        $selection.Font.Bold = $false
        $selection.Font.Color = $colorDarkGray
        $selection.ParagraphFormat.SpaceBefore = 0
        $selection.ParagraphFormat.SpaceAfter = 6
        $selection.TypeText($s.body)
        $selection.TypeParagraph()
    }

    # Section 5 - Interju strategia & Cheat Sheet
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 16
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec5_title)
    $selection.TypeParagraph()

    $selection.Font.Size = 10.0
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($data.sec5_lead)
    $selection.TypeParagraph()

    foreach ($qa in $data.sec5_items) {
        $selection.Font.Size = 10.5
        $selection.Font.Bold = $true
        $selection.Font.Color = $colorUniqaBlue
        $selection.ParagraphFormat.SpaceBefore = 8
        $selection.ParagraphFormat.SpaceAfter = 3
        $selection.TypeText($qa.q)
        $selection.TypeParagraph()

        $selection.Font.Size = 9.0
        $selection.Font.Bold = $false
        $selection.Font.Color = $colorDarkGray
        $selection.ParagraphFormat.SpaceBefore = 0
        $selection.ParagraphFormat.SpaceAfter = 6
        $selection.TypeText($qa.a)
        $selection.TypeParagraph()
    }

    # Save
    $doc.SaveAs([ref]$outputPath, [ref]16)
    Write-Output "DOCX_CLEANED_SUCCESSFULLY: $outputPath"
}
catch {
    Write-Error "DOCX_GENERATION_FAILED: $_"
}
finally {
    if ($doc) {
        $doc.Close([ref]$false)
    }
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

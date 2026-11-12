# convert_all_md_to_docx.ps1
# Batch converts all markdown guides and playbooks to .docx using convert_md_to_docx.ps1

$docsDir = "$PSScriptRoot\..\docs"
$converter = "$PSScriptRoot\convert_md_to_docx.ps1"

$filesToConvert = @(
    @{
        Input = "$docsDir\FELKESZITO_ANYAG_UNIQA_COLLECTION_RETENTION_LEAD.md"
        Output = "$docsDir\FELKESZITO_ANYAG_UNIQA_COLLECTION_RETENTION_LEAD.docx"
        Title = "UNIQA Collection & Retention Lead - Vezetoi Felkeszito Anyag"
    },
    @{
        Input = "$docsDir\INTERJU_SZEREPLOK_DOSSSZIER.md"
        Output = "$docsDir\INTERJU_SZEREPLOK_DOSSSZIER.docx"
        Title = "UNIQA Interju Szereplok es Donteshozok Szakmai Dosszieja"
    },
    @{
        Input = "$docsDir\INTEGRALT_JOGFORRAS_PLAYBOOK_2026.md"
        Output = "$docsDir\INTEGRALT_JOGFORRAS_PLAYBOOK_2026.docx"
        Title = "Integralt Jogforras es Vezetoi Dontestamogato Playbook (v2.0)"
    },
    @{
        Input = "$docsDir\VEGREHAJTASI_ES_BIZTOSITASI_JOGANYAG_2026_UNIQA.md"
        Output = "$docsDir\VEGREHAJTASI_ES_BIZTOSITASI_JOGANYAG_2026_UNIQA.docx"
        Title = "Birosagi Vegrehajtasi es Biztositasi Jogszabalyvaltozasok 2026"
    },
    @{
        Input = "$docsDir\UNIQA_CEGBEMUTATO_ES_ERTEKEK_2026.md"
        Output = "$docsDir\UNIQA_CEGBEMUTATO_ES_ERTEKEK_2026.docx"
        Title = "UNIQA Biztosito - Vallalati Ertekek es Cegbemutato 2026"
    }
)

Write-Host "================================================================="
Write-Host "  Markdown -> Word (.docx) Batch Konverzio Inditasa (Word COM)   "
Write-Host "================================================================="

foreach ($item in $filesToConvert) {
    if (Test-Path $item.Input) {
        & $converter -InputPath $item.Input -OutputPath $item.Output -DocTitle $item.Title
    } else {
        Write-Warning "Nem talalhato a forrasfajl: $($item.Input)"
    }
}

Write-Host "`nA konverzio befejezodott! Ellenorzes:"
Get-ChildItem -Path $docsDir -Filter "*.docx" | Select-Object Name, Length, LastWriteTime

# Set strict error preference
$ErrorActionPreference = "Stop"

$outputPath = "C:\Users\norbi\.gemini\PROJECTS\uniqa-collection-retention-lead-prep\docs\Uniqa_Collection_Retention_Lead_Joganyag_2026.docx"

if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force -ErrorAction SilentlyContinue
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Add()
    $selection = $word.Selection

    # Colors
    $colorNavy = 5844239      # RGB(15, 44, 89)
    $colorUniqaBlue = 10375424 # RGB(0, 81, 158)
    $colorDarkGray = 3355443   # RGB(51, 51, 51)

    # Margins (54 pt = 0.75 in)
    $doc.PageSetup.TopMargin = 54
    $doc.PageSetup.BottomMargin = 54
    $doc.PageSetup.LeftMargin = 54
    $doc.PageSetup.RightMargin = 54

    # Header & Footer
    $header = $doc.Sections.Item(1).Headers.Item(1)
    $header.Range.Text = "UNIQA BIZTOSITO - COLLECTION & RETENTION LEAD INTERJU FELKESZITO (2026)"
    $header.Range.Font.Name = "Calibri"
    $header.Range.Font.Size = 8.5
    $header.Range.Font.Color = 8421504
    $header.Range.ParagraphFormat.Alignment = 2

    $footer = $doc.Sections.Item(1).Footers.Item(1)
    $footer.Range.Text = "Magyar Kozlony & Vegrehajtasi Joganyag Valtozasok (2026.04.12 - 2026.11.05) | Szakmai Felkeszito"
    $footer.Range.Font.Name = "Calibri"
    $footer.Range.Font.Size = 8.5
    $footer.Range.Font.Color = 8421504

    # Title
    $selection.Font.Name = "Calibri"
    $selection.Font.Size = 20
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText('BIROSAGI VEGREHAJTASI ES BIZTOSITASI SZEKTORALIS JOGSZABALYVALTOZASOK')
    $selection.TypeParagraph()

    # Subtitle
    $selection.Font.Size = 13
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorUniqaBlue
    $selection.ParagraphFormat.SpaceAfter = 12
    $selection.TypeText('Vezetoi Dontestamogato Elemzes es Interju-Strategia (2026.04.12 - 2026.11.05)')
    $selection.TypeParagraph()

    # Metadata
    $selection.Font.Size = 9.5
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceAfter = 16
    $selection.TypeText('Celpozicio: Uniqa Biztosito - Collection & Retention Lead | Datum: 2026. november 5. | Forrasok: Magyar Kozlony, Indokolasok Tara, MNB Rendeletek, Vht., Bit., Gfbt.')
    $selection.TypeParagraph()

    # Section 1
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 10
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText('1. VEZETOI OSSZEFOGLALO (EXECUTIVE SUMMARY)')
    $selection.TypeParagraph()

    $selection.Font.Size = 10.5
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText('A 2026. aprilis 12. es 2026. november 5. kozotti idoszakban a magyar jogalkoto es az MNB tobb olyan melyrehato jogszabaly-modositast leptetett eletbe, amelyek kozvetlenul meghatarozzak az Uniqa Collection & Retention teruleteinek mukodesi modelljet:')
    $selection.TypeParagraph()

    $bList = @(
        'Paradigmatikus vegrehajtasi reform (1158/2026. Korm. hat.): A Kormany elrendelte a birosagi vegrehajtas nonprofit, birosagi szervezeten beluli megvalositasanak elokesziteset. Ez az onallo vegrehajtoi (MBVK) modell allamositasat es az eljarasok atmeneti lassulasat hozza, felertekelve az Uniqa peren kivuli korai (soft) behajtasat.',
        'Kenyszer-vegrehajtasi moratorium (2026. evi XVI. tv.): A tisztessegtelen fogyasztoi szerzodesek felulvizsgalata erdekeben a devizahiteles torvenyek hatalya ala tartozo ugyekben azonnali vegrehajtasi stop es kotelezo perfuggesztes lepett eletbe. Ez erinti a hitelfedezeti elet- es vagyonbiztositasok dijfizeteset es a visskereseti jogervenyesitest.',
        'Uj vegrehajtas alol mentes juttatas a Vht.-ban (2026. evi XXXIII. tv.): A Vht. 92. par. uj e) pontja 100%-ban mentesiti a vegrehajtas alol az allami iskolakezdesi tamogatast (100 000 Ft/gyermek). Hatosagi atutalasi megbizasnal a bankok ezt kotelesek levonasmentesen hagyni.',
        'Uj MNB panaszkezelesi keretrendszer (16/2026. MNB rend.): 2026. julius 18-tol felvaltja a 66/2021. rendeletet. Szigorubba valt az elektronikus kapcsolattartas (ugyfelportal, app), kotelezove valt a honlapi kozzetetel es a 30 napos hatarido a behajtasi vitak kezeleseben.',
        'Uj szuperhatosag (NVVH) felallitasa (2026. evi XXXIV. tv.): Atvezetesre kerult a biztositasi titok aloli mentesseg a Bit.-ben es a kotvenynyilvantartasi adatszolgaltatas a Gfbt.-ben.'
    )

    foreach ($item in $bList) {
        $selection.TypeText(' - ' + $item)
        $selection.TypeParagraph()
    }

    # Section 2
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 14
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText('2. TETELES JOGSZABALYVALTOZASI LAJSTROM (2026.04.12 - 2026.11.05)')
    $selection.TypeParagraph()

    $selection.Font.Size = 10
    $selection.Font.Bold = $false
    $selection.Font.Color = $colorDarkGray
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 8
    $selection.TypeText('Az alabbi tablazat tartalmazza a vizsgalt idoszakban a Magyar Kozlonyben kihirdetett relevans birosagi vegrehajtasi es biztositasi szektoralis jogszabalyokat:')
    $selection.TypeParagraph()

    # Table definition
    $tableData = @(
        @('Kozlony & Datum', 'Jogszabaly Szama & Cime', 'Modositott Szakasz', 'Hatalybalepes', 'Fo Targykor', 'Biztositasi & Vegrehajtasi Relevancia'),
        @("MK 2026/53.`n2026.05.18.", '1158/2026. (V. 18.) Korm. hat.', '1994. evi LIII. tv. (Vht.) atfogo felulvizsgalata', "2026.05.18.`n(hatarido: 07.14.)", 'Birosagi vegrehajtas nonprofit allami modellje', 'KIEMELTEN RELEVANS: Vegrehajtoi halozat birosagi integralasa. Az Uniqa hard collection SLA-it at kell hangolni a lassulo vegrehajtoi kapacitasok miatt.'),
        @("MK 2026/63.`n2026.06.03.", '16/2026. (VI. 3.) MNB rend.', '66/2021. MNB rend. helyebe; Bit. 439. par. a)', '2026.07.18.', 'Penzugyi szervezetek uj panaszkezelesi szabalyai', 'KIFEJEZETTEN RELEVANS: A dijhatralek miatti felszolitasokra erkezo panaszok gyors, digitalis kezelese (portal, app); a retention csapat legfobb korai eszkoze.'),
        @("MK 2026/66.`n2026.06.09.", '2026. evi XVI. torveny', 'Devizahiteles tv.-ek; Vht. eljarasi szabalyok', '2026.06.10.', 'Tisztessegtelen szerzodesek vegrehajtasi moratoriuma', 'KIEMELTEN RELEVANS: Vegrehajtasi stop es kotelezo perfuggesztes. A hitelfedezeti elet- es vagyonbiztositasok dijkoveteleseit retention egyezseggel kell vedeni.'),
        @("MK 2026/97.`n2026.07.23.", '2026. evi XXXIII. torveny', 'Vht. 92. par. uj e) pont', '2026.07.24.', 'Iskolakezdesi tamogatas (100e Ft) mentesitese', 'VEGREHAJTASI RELEVANCIA: A Vht. mentes vagyontargyak kore bovult. Hatosagi atutalasi megbizasnal a bankok nem vonhatjak le a biztositasi dijtartozasra sem.'),
        @("MK 2026/100.`n2026.07.28.", '2026. evi XXXIV. torveny', 'Vht. 13., 16., 35. par.; Bit. 135-142. par.; Gfbt. 47., 51. par.', "2026.08.01.`n(fokozatos)", 'Nemzeti Vagyonvisszaszerzesi Hivatal (NVVH) felallitasa', 'BIZTOSITASI & VHT. RELEVANCIA: A biztositasi titok aloli mentesseg bovult. Az NVVH kozvetlenul igenyelhet adatot a biztositoi- es karnyilvantartasbol.'),
        @("MK 2026/112.`n2026.08.13.", '2026. evi XL. torveny', 'Vht. 253/E. par. (5), 253/G. par. (3)', '2026.09.01.', 'Rendeszeti feladatellatas es adatszolgaltatas pontositas', 'VEGREHAJTASI TECHNIKAI: Gepjarmu-foglalasoknal es korozeseknel az eljaro rendori szervek pontositasra kerultek Casco es lizingfedezeteknel.'),
        @("MK 2026/131.`n2026.09.11.", '2026. evi XLIX. torveny', 'Vht. 16., 23/B., 33. par.; Gfbt. 1. par. d)', '2026.10.01.', 'Varmegye -> Megye terminologiai korrekcio', '[Nincs erdemi relevancia - pusztan kozigazgatasi elnevezesbeli atvezetes]'),
        @("MK 2026/148.`n2026.10.08.", '10/2026. (X. 8.) SZTFH rend.', '13/2021. SZTFH rend. es kapcsolodo dijrendeletek', '2026.10.15.', 'SZTFH hataskori es arveresi terminologia', 'KOZVETETT VHT. RELEVANCIA: Az elektronikus arveresi rendszer es igazgatasi dijak jogszabalyi hivatkozasainak korrekcioja.')
    )

    $rowCount = $tableData.Count
    $colCount = $tableData[0].Count
    $range = $selection.Range
    $table = $doc.Tables.Add($range, $rowCount, $colCount)
    
    $table.Borders.Enable = $true
    $table.Borders.InsideColor = 13421772
    $table.Borders.OutsideColor = 10066329

    $colWidths = @(70, 95, 75, 65, 95, 104)
    for ($c = 1; $c -le $colCount; $c++) {
        $table.Columns.Item($c).Width = $colWidths[$c - 1]
    }

    for ($r = 1; $r -le $rowCount; $r++) {
        $row = $table.Rows.Item($r)
        $row.HeadingFormat = ($r -eq 1)
        $row.AllowBreakAcrossPages = $false

        for ($c = 1; $c -le $colCount; $c++) {
            $cell = $table.Cell($r, $c)
            $cell.Range.Text = $tableData[$r - 1][$c - 1]
            $cell.Range.Font.Name = 'Calibri'

            if ($r -eq 1) {
                $cell.Range.Font.Size = 9
                $cell.Range.Font.Bold = $true
                $cell.Range.Font.Color = 16777215
                $cell.Shading.BackgroundPatternColor = $colorNavy
                $cell.VerticalAlignment = 1
                $cell.Range.ParagraphFormat.Alignment = 1
            } else {
                $cell.Range.Font.Size = 8.5
                $cell.Range.Font.Bold = $false
                $cell.Range.Font.Color = $colorDarkGray
                $cell.VerticalAlignment = 0

                if ($r % 2 -eq 0) {
                    $cell.Shading.BackgroundPatternColor = 16316664
                }

                if ($c -eq 1 -or $c -eq 4) {
                    $cell.Range.ParagraphFormat.Alignment = 1
                } else {
                    $cell.Range.ParagraphFormat.Alignment = 0
                }
            }
        }
    }

    $selection.SetRange($table.Range.End + 1, $table.Range.End + 1)
    $selection.TypeParagraph()

    # Section 3
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 14
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText('3. TARTALMI VALTOZASOK ES INDOKOLASOK OSSZEGZESE')
    $selection.TypeParagraph()

    $sections = @(
        @{
            Title = '3.1. A birosagi vegrehajtas atalakitasa (1158/2026. Korm. hatarozat)'
            Body = 'A Kormany dontese ertelmeben a birosagi vegrehajtast ki kell vonni a maganvallalkozasi profitorientalt keretek kozul, es a birosagok belso szervezeti egysegekent kell mukodtetni. Az igazsagugyi miniszter altal elokeszitett reform celja a korrupcios kockazatok felszamolasa es az igazsag szolgaltatasaba vetett kozbizalom helyreallitasa. Biztositoi koveteleskezelesi szempontbol ez azt jelenti, hogy az atallasi idoszakban megno a vegrehajtasi ugyhatralek, a birosagi apparatus leterheltsege no, ezert a peresites elotti soft collection egyezsegek felertekelodnek.'
        },
        @{
            Title = '3.2. Fogyasztoi moratorium es perfuggesztes (2026. evi XVI. torveny)'
            Body = 'A torveny a devizahiteles szerzodesek felulvizsgalata idejere hivatalbol elrendeli a polgari perek soron kivuli felfuggeszteset es a vegrehajtasi cselekmenyek tilalmat. Az Indokolasok Tara (IT 2026/39.) kiemeli: a jogalkoto celja az irreverzibilis adosi vagyonvesztes (pl. ingatlanarveres, kilakoltatas) megakadalyozasa az egyseges Kuria/EUB joggyakorlat megszuleteseig. A biztositonal ez erinti a hitelfedezeti eletbiztositasi allomanyokat: nem szabad hagyni, hogy a felfuggesztett hitelek miatt a biztositasi fedezet automatikusan dijtorlessel megszunjon.'
        },
        @{
            Title = '3.3. Iskolakezdesi tamogatas mentesitese (Vht. 92. par. e) pont - 2026. evi XXXIII. tv.)'
            Body = 'A jogalkoto a Vht. 92. par.-at egy uj e) ponttal egeszitette ki, 100%-ban mentesitve a vegrehajtas alol a gyermekenkent 100 000 Ft osszegu allami iskolakezdesi tamogatast. Az Indokolasok Tara (IT 2026/50.) szerint ezen osszeg nem valhat hitelezoi kielegitesi alappa. Folyoszamla-letiltasnal a penzintezetek felelossege a mentesitett osszeg levonas aloli feloldasa.'
        },
        @{
            Title = '3.4. Uj MNB panaszkezelesi keretrendszer (16/2026. (VI. 3.) MNB rendelet)'
            Body = 'Az MNB rendelete julius 18-tol valtja fel a 66/2021. MNB rendeletet. Kotelezove teszi a digitalis panaszkezelesi opciokat: szobeli panasz eseten az ugyfel keresere kotelezo digitalis csatornan (ugyfelkapu, e-mail) elkuldani az erdemi valaszt. Minden biztositonak kotelezo a honlapjan elerhetove tenni a panaszkezelesi tajekoztatot. A kintlevoseg-kezelesi es dijnemfizetesi vitakban a panasz kivizsgalasaig felfuggesztett eljaras megelozi a kesobbi felugyeleti fogyasztovedelmi birsagokat.'
        }
    )

    foreach ($s in $sections) {
        $selection.Font.Size = 11.5
        $selection.Font.Bold = $true
        $selection.Font.Color = $colorUniqaBlue
        $selection.ParagraphFormat.SpaceBefore = 8
        $selection.ParagraphFormat.SpaceAfter = 3
        $selection.TypeText($s.Title)
        $selection.TypeParagraph()

        $selection.Font.Size = 10
        $selection.Font.Bold = $false
        $selection.Font.Color = $colorDarkGray
        $selection.ParagraphFormat.SpaceBefore = 0
        $selection.ParagraphFormat.SpaceAfter = 6
        $selection.TypeText($s.Body)
        $selection.TypeParagraph()
    }

    # Section 4
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = $colorNavy
    $selection.ParagraphFormat.SpaceBefore = 14
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText('4. UNIQA COLLECTION & RETENTION LEAD INTERJU-STRATEGIA')
    $selection.TypeParagraph()

    $qaList = @(
        @{
            Q = 'Kerdes 1: Hogyan reagalna vezetokent a birosagi vegrehajtas allamositasara (1158/2026. Korm. hat.), ami atmenetileg novelheti a jogi behajtas atfutasi idejet?'
            A = 'Valaszstrategia: A hitelezoi hatekonysagot a soft collection szakaszban kell megnyerni. Mivel a birosagi vegrehajtoi apparatus leterheltsege atmenetileg no, harom pillerre tamaszkodom: 1) D+15 es D+45 napos prediktiv scoring-alapu korai megkereses, azonnali digitalis reszletfizetesi opcioval; 2) FMH es zaradekolasi szures: csak a likvid vagyonnal rendelkezo, nagy osszegu dij- es regresszkoveteleseket engedjuk vegrehajtasba, elkerulve a felesleges eljarasi illetekek beragadasat; 3) A kulso jogi kepviselokkel kotott keretszerzodesekben az atfutasi idokhöz kotott merfoldko-KPI-okat vezetunk be.'
        },
        @{
            Q = 'Kerdes 2: Milyen hatassal van a 16/2026. MNB rendelet az Uniqa Retention (ugyfelmegtartasi) folyamataira?'
            A = 'Valaszstrategia: Az uj MNB rendelet kotelezove teszi a panaszkezeles digitalizalasat es a 30 napos hatarido szigoru betartasat. A dijnemfizetes miatti megkeresesek gyakran torkollnak panaszba (nem kapott csekket, vitatott karkifizetes). Ha a Collection es a Panaszkezeles szetvalik, az ugyfel elveszik. Vezetokent integralt folyamatot vezetek be: beerkezo panasz eseten a jogi behajtasi orat megallitjuk, az ugyfelportalon vagy appon keresztul 5 munkanapon belul kozvetlen megoldast kinalunk (ujrakotesi ajanlattal parositva), megelozve a lapse rate emelkedeset es az MNB birsagot.'
        },
        @{
            Q = 'Kerdes 3: Hogyan kezeli a regressz-koveteleseket es a hitelfedezeti allomanyokat a 2026. evi XVI. torveny alatti moratoriumban?'
            A = 'Valaszstrategia: A fogyasztoi hitelek vegrehajtasi tilalma kockazatot jelent a kapcsolodo hitelfedezeti elet- es vagyonbiztositasok dijfolytonossagara. Azonnali portfolio-auditot vegzunk a bankszovetsegi es partnerbanki kapcsolatokon keresztul. Az erintett szerzodeseknel nem az agressziv jogi behajtast alkalmazzuk, hanem a Retention First logikat: felvilagositjuk az ugyfelet, hogy a hitel moratoriuma ellenere a fedezet fenntartasa (akar dijcsokkentes vagy konstrukcio-valtas mellett) vedi a vagyontargyat, igy az ugyfel lojalitasa megmarad.'
        }
    )

    foreach ($qa in $qaList) {
        $selection.Font.Size = 11
        $selection.Font.Bold = $true
        $selection.Font.Color = $colorUniqaBlue
        $selection.ParagraphFormat.SpaceBefore = 8
        $selection.ParagraphFormat.SpaceAfter = 3
        $selection.TypeText($qa.Q)
        $selection.TypeParagraph()

        $selection.Font.Size = 9.5
        $selection.Font.Bold = $false
        $selection.Font.Color = $colorDarkGray
        $selection.ParagraphFormat.SpaceBefore = 0
        $selection.ParagraphFormat.SpaceAfter = 6
        $selection.TypeText($qa.A)
        $selection.TypeParagraph()
    }

    # Save
    $doc.SaveAs([ref]$outputPath, [ref]16)
    Write-Output "DOCX_CREATED_SUCCESSFULLY: $outputPath"
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

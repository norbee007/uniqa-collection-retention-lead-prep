# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0] - 2026-11-05

### Added & Verified
- **Teljes Közlöny-audit (2026.04.12 – 2026.11.05: MK 36. – 165. számok)**:
  - Átfogó forrásvizsgálat: az RSS feed 100 elemes korlátjának feltárása és a 2026. április 12. és 2026. augusztus 14. közötti lapszámok (MK 36. – 109. számok) szisztematikus ellenőrzése.
  - **Új azonosított jogszabály**: **2026. évi XIV. törvény** (MK 2026/44. szám, kihirdetve: 2026.05.09., hatályos: 2026.05.14.).
    - Beiktatta a Vht. 92. § új *d)* pontját (sportszervezetek működési célú állami támogatásának végleges mentesítése).
    - Beiktatta a Vht. új **92/A. §**-át (az Otthontámogatás teljes körű végrehajtási mentessége, levonási tilalma és ex lege visszajárása).
  - A Markdown elemzés és a formázott `.docx` jelentés frissítve az új jogszabállyal és hatásvizsgálattal.

## [1.1.2] - 2026-11-05

### Fixed
- **Bullet karakter kódolási hiba (`â€˘` mojibake) megszüntetése**:
  - A PowerShell generátorban korábban hardkódolt UTF-8 bullet karakter (`•`, U+2022) a Word COM API hívásakor ANSI kontextusban `â€˘` karakterláncként jelent meg.
  - A felsorolási elemek átalakításra kerültek tiszta, natív Word bekezdés-behúzásra (`LeftIndent = 18`, `FirstLineIndent = -18`) és szabványos kötőjeles jelölésre.
  - A `docs/Uniqa_Collection_Retention_Lead_Joganyag_2026.docx` belső XML vizsgálata (0 darab `0x00E2` karakter) igazolta az anomália teljes megszűnését.

## [1.1.1] - 2026-11-05

### Fixed
- **Magyar ékezetek és karakterkódolás (UTF-8) javítása a Word (.docx) állományban**:
  - A korábbi PowerShell script beépített ANSI sztring-kezelése miatt elveszett ékezetek (á, é, í, ó, ö, ő, ú, ü, ű) teljes helyreállítása.
  - A dokumentum szövegének szétválasztása és külső, tiszta UTF-8 kódolású adatbázisba (`data/docx_content.json`) szervezése.
  - A generátor script (`scripts/generate_docx.ps1`) átírása natív `[System.IO.File]::ReadAllText` UTF-8 beolvasásra, így a Word COM felé már hibátlan Unicode karakterek kerülnek átadásra.
  - A `docs/Uniqa_Collection_Retention_Lead_Joganyag_2026.docx` újragenerálva teljes és helyes magyar ékezethasználattal.

## [1.1.0] - 2026-11-05

### Added
- **Word (.docx) Dokumentum Generálás**: Strukturált, professzionális Word dokumentum készítése a teljes jogszabály-gyűjtésből, tételes táblázattal, indokolásokkal és az interjú-stratégiával:
  - Fájl: `docs/Uniqa_Collection_Retention_Lead_Joganyag_2026.docx`
  - Generáló szkript: `scripts/generate_docx.ps1`
  - Fejléc, lábléc, oldalszámozás, Uniqa kék & Navy színpaletta, zebra-csíkozású lajstromtáblázat.
- **Launcher frissítés (`run.bat`)**: 4. menüpont hozzáadva a Word dokumentum közvetlen megnyitásához.

## [1.0.0] - 2026-11-05

### Added
- **5P Mester Prompt Architecture**: Bírósági végrehajtási és biztosítási szektorális joganyag letöltő és elemző pipeline.
- **Tételes Jogszabály-lajstrom (2026.04.12 – 2026.11.05)**:
  - 1158/2026. (V. 18.) Korm. határozat (bírósági végrehajtás átalakítása nonprofit bírósági rendszerré).
  - 16/2026. (VI. 3.) MNB rendelet (új panaszkezelési keretrendszer biztosítóknak).
  - 2026. évi XVI. törvény (fogyasztói devizaügyletek per- és végrehajtási moratóriuma).
  - 2026. évi XXXIII. törvény (Vht. 92. § *e)* mentesített iskolakezdési támogatás).
  - 2026. évi XXXIV. törvény (NVVH adatszolgáltatási kötelezettség és biztosítási titok módosítás).
  - 2026. évi XL. törvény (rendészeti szervek és körözési adatszolgáltatás pontosítása a Vht.-ban).
  - 10/2026. (X. 8.) SZTFH rendelet (árverési és felügyeleti harmonizáció).
- **Indokolások Tára Kivonatok**: Jogpolitikai és jogalkotói célok szisztematikus feldolgozása.
- **Uniqa Collection & Retention Lead Interjú-Stratégia**:
  - Soft & Hard collection Mermaid folyamatábra.
  - Portfólió-tisztítási és SLA-optimalizálási stratégia.
  - Kulcskérdések és modell-válaszok a vezetői interjúra.
- **PowerShell Automatizációs Eszköztár**:
  - `scripts/fetch_feed.ps1`: Valós idejű Magyar Közlöny RSS feed letöltő és szűrő.
  - `run.bat`: Gyökérszintű egykattintásos indító launcher Windows környezetre.

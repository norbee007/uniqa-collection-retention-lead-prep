# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

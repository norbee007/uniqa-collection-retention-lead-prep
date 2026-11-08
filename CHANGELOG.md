# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.1.0] - 2026-11-08

### 🚀 Új Funkciók (Features)
- **Személyre szabott Vezetői Interjú Felkészítő Rendszer (`docs/FELKESZITO_ANYAG_UNIQA_COLLECTION_RETENTION_LEAD.md`)**:
  - Teleki Norbert szakmai profiljára, 10+ éves pénzügyi és követeléskezelési tapasztalatára, 50 fős operatív vezetésére, 6 Mrd Ft-os számlaegyeztetési projektjére és Innovációs Díjas (18 RPA bot / 20% kapacitásbővítés) eredményeire szabott 30 másodperces Executive Elevator Pitch.
  - A pozíció pénzügyi-számviteli (Finance & Accounting) mélyelemzése: a kintlévőség-kezelés eltolódása a peres behajtástól a díjkönyvelési/technikai hibák kiszűrése (60%) és a Soft Reminder / Retention (40%) felé.
  - 10 kidolgozott, éles STAR modellválasz HR és Szakmai vezetői kérdésekre (számszerűsített eredményekkel, MNB 13/2022, 16/2026 rendeletekkel, 1158/2026 nonprofit végrehajtási reformmal és Vht. 92. § mentesítésekkel).
  - Intelligens visszakérdezési stratégia (Reverse Interviewing) Mészáros Petrának és a Szakmai Vezetőnek.
  - 1-oldalas Interjú Napi Puskázó Kártya (Quick-Glance Cheat Sheet).
- **Interjú Szereplők és Döntéshozók Szakmai Dossziéja (`docs/INTERJU_SZEREPLOK_DOSSSZIER.md`)**:
  - **Mészáros Petra (HR Generalist)**: Közel 6 éves Grafton Finance & Accounting toborzási múltjának, Corvinus MA (Vezetés és Szervezés) és NKE adóigazgatási végzettségének, HPTI kompetencia-értékelési modelljének feltárása.
  - **Szakmai Vezetés (Finance & Accounting)**: Kiss Mariann (Számviteli rendért felelős vezető / Head of Accounting, MABISZ bizottsági tag) és Pavel Kopylov (CFO / Igazgatósági tag) elvárásainak, IFRS / Solvency II kintlévőségi tőkekövetelményeinek és folyamat-tisztasági igényeinek részletes profilja.

### 🔒 Biztonság & Stabilitás (Security & Hardening)
- **Zero Data Leak v2.1 Audit**: `.gitignore` kiterjesztése a személyes rekrúteri profilkivonatok (`*profile*.pdf`, `*recruiter*.pdf`) szűrésére, megelőzve az érzékeny adatok Git repóba kerülését.

### ⚡ Rendszerintegráció & Launcher
- **run.bat Kibővítése**: 8-pontos gyökérszintű interaktív launcher menü a felkészítő anyagok és döntéshozói profilok azonnali megnyitásával.

## [2.0.0] - 2026-11-06

### Added & Architecture Expansion
- **4 Új Jogforrás-Pillér Integrációja (Összesen 271 feldolgozott tétel)**:
  - **Pillér 1 – MNB Releváns Ajánlások (2020-tól napjainkig)**: 143 elem feldolgozva (37 kiemelten releváns), köztük az etikus adóskezelés (13/2022. MNB ajánlás), a panaszkezelési határidők és minőségbiztosítás (16/2026. MNB rendelet és 7/2021. ajánlás), valamint a termékfelügyelet és irányítás (14/2025. POG ajánlás).
  - **Pillér 2 – BHGY Anonimizált Döntések (2025-ig visszamenőleg)**: Bírósági határozatok gyűjteménye a Vht. 41/56. § végrehajtás-megszüntetési perekről, a mentesített juttatások inkasszójáról és a fedezetlenségi díj jogérvényesítéséről.
  - **Pillér 3 – Kúria Döntések és Jogegységi Határozatok (2025-ig visszamenőleg)**: 42 Kúriai határozat és JPE a kötelmi elévülésről (Ptk. 6:468. §), a türelmi időről és póthatáridős felmondásokról (Ptk. 6:448. §), az egyoldalú szerződési feltételek érvénytelenségéről és a biztosítói regresszről (Gfbt. 34. §).
  - **Pillér 4 – SZTFH Rendeletek és Határozatok (2025-től napjainkig)**: 80 SZTFH norma (20 végrehajtási relevanciájú) a bírósági végrehajtói hálózat felügyeletéről (11/2025. rendelet), az Elektronikus Árverési Rendszerről (10/2026. és 13/2021. EÁR) és végrehajtói panaszügyekről.
- **Automatizált Web-Scraping Pipeline**:
  - `scripts/run_all_scrapers.ps1`: Fő orchestrator és aggregáló modul.
  - `scripts/scrape_mnb.ps1`, `scripts/scrape_sztfh.ps1`, `scripts/scrape_kuria.ps1`, `scripts/scrape_bhgy.ps1`.
  - Normalizált JSON adatbázisok a `data/` könyvtárban (`aggregated_legal_data.json` statisztikákkal).
- **Új Dokumentáció és Vezetői Eszközök**:
  - `docs/SPECIFIKACIO_ES_ADATFORRASOK_V2.md`: Részletes rendszerspecifikáció, szakterületi szűrőszótár, JSON adatséma (v2.0) és Zero Data Leak v2 protokoll.
  - `docs/MESTER_PROMPT_V2_MULTI_SOURCE.md`: Teljes multidiszciplináris 5P Mester Prompt sablonrendszer Uniqa Collection & Retention Lead szintre optimalizálva.
  - `docs/INTEGRALT_JOGFORRAS_PLAYBOOK_2026.md`: Átfogó stratégiai playbook, Soft & Hard Collection folyamatmodell, KPI mátrix és 5 nyerő interjú modell-válasz.
- **Újragenerált Formázott Word (.docx) Dokumentum**:
  - `docs/Uniqa_Collection_Retention_Lead_Joganyag_2026.docx`: 5 strukturált fejezet, 6 formázott táblázat (KPI mátrix, MK 2026, MNB, Kúria, BHGY, SZTFH), Uniqa arculati színpaletta, 100%-os UTF-8 magyar karakterkészlet.
  - `scripts/generate_docx.ps1`: Automatizált Word COM riportgenerátor frissítve.
- **Launcher Frissítés**:
  - `run.bat`: Kibővítve v2.0.0-ra 1-kattintásos scraper futtatással és dokumentum-megnyitással.

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

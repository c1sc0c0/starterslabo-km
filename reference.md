# Kilometervergoeding — cell reference

Official blank: download the FAQ bijlage (VL-Brabant or your regio). This skill does not ship that file.

Sheets: `Blanco` (empty template), month tabs named `MM_YY` (e.g. `10_26`).

## Header (grey `C1`–`C11`)

| Cell | Label | JSON key |
|------|-------|----------|
| `C1` | Naam KO | `header.name` |
| `C2` | Adres | `header.address` |
| `C3` | Locatie | `header.location` |
| `C4` | Tel. n° | `header.phone` |
| `C5` | Email | `header.email` |
| `C6` | Merk + type | `header.vehicle` |
| `C7` | N° plaat | `header.plate` |
| `C8` | Brandstof | `header.fuel` |
| `C9` | CO2 uitstoot | `header.co2` |
| `C10` | % VU (dropdown 0%/50%/60%) | `header.vu_pct` — from Verklaring op eer / coach |
| `C11` | Bankrek.n° | `header.iban` |

`E5` is often pre-filled with the regional Starterslabo address (do not overwrite unless asked).

## Period / rate

| Cell | Meaning |
|------|---------|
| `E12` | Indieningsdatum = laatste dag van de maand |
| `E13` | Periode (e.g. `oktober 2026`) |
| `E56` | Rate label |
| `F56` | €/km rate |
| `F53` | `=SUM(F16:F52)` month km |
| `F58` | `=F53*F56` amount |
| `D54` | Datum + km-stand begin facturatie |
| `D55` | Km-stand bij indiening laatste km-vergoeding |
| `D56` | **Totaal aantal km (traject)** — all months, max 24.000 |

### Rates (forfait ambtenaren)

| Vanaf | Tarief |
|-------|--------|
| 01/01/2025 | €0,4290 |
| 01/01/2026 | €0,4326 |
| 01/07/2026 | €0,4440 |

## Trip rows (grey)

Rows **17–52** (36 lines). Columns:

| Col | Field |
|-----|-------|
| `B` | Datum |
| `C` | Contact |
| `D` | Actie |
| `E` | Locatie + adres (straat + nr + postcode + gemeente) |
| `F` | Aant. KM |

Row 16 is in the SUM range but has no grey fill in the official blank — prefer starting at 17.

## Trips JSON

See [trips.example.json](trips.example.json). Fiction placeholders only — never commit a real filled workbook or profile JSON.

Combine heen+terug on **one line**; `km` = single leg × 2 (unless one-way only).

## Portal upload (onkostennota)

Use skill **starterslabo-costs** / `create_draft_expense.py` with `--dry-run` first.

| Field | Value |
|-------|-------|
| Leverancier | `KO <Achternaam> <Voornaam>` |
| BTW regime | Geen |
| Onderwerp | `KM <JJ> <MM> <Achternaam> <Voornaam>` — vb. `KM 22 01 Jansen Jan` |
| Factuurnummer | `Km <JJ> <MM>` — vb. `Km 22 01` |
| Datum | Laatste dag van de maand |
| Rekening | `614317` (lichte vracht / elektrisch, 0% verworpen) or `614324` (otherwise; % VU from Verklaring) |
| Bedrag | `F58` from Excel (comma decimal) |
| Bijlage | De ingevulde `.xlsx` |

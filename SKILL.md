---
name: starterslabo-km
description: >-
  Fill and maintain the Starters Labo kilometervergoeding Excel
  (rittenadministratie), apply ambtenaren rates, track traject km, and
  upload via Aankoop onkostennota on mijn.starterslabo.be. Use when the
  user mentions kilometervergoeding, KM-vergoeding, rittenadministratie,
  beroepsmatig gereden kilometers, or KM Excel tabs MM_YY.
---

# Starters Labo — Kilometervergoeding

Help maintain the official LABO **kilometervergoeding** Excel and submit it
via [mijn.starterslabo.be](https://mijn.starterslabo.be) **Aankoop**.

Not affiliated with Starterslabo. Work only on workbooks the user provides.
Never invent trip facts (contact, actie, adres). Never commit filled
workbooks or profile JSON with real data.

This skill does **not** ship the official Excel. The user downloads the blank
from the portal FAQ bijlage.

## FAQ rules (verbatim)

Dit document gebruik je voor opgave van de beroepsmatig gereden kilometers met jouw wagen in functie van jouw traject bij Starterslabo.

Je kan een kilometervergoeding indienen als je 'Verklaring op eer' in orde is - zie FAQ \ Verklaring op eer - en als je voldoende budget hebt.

Maar een nieuwe leverancier aan via de button 'Aankoop'

Naam van de leverancier: KO Jouw Achternaam + voornaam (deze leverancier gebruik je ook voor je onkostennota 's)

BTW regime: Geen

Verantwoording conform rittenadministratie: alle gevraagde gegevens invullen (in de grijze vakjes): datum, contact, actie, locatie + adres (straatnaam + huisnummer + postcode + gemeente)

Gehanteerd tarief: forfait voor ambtenaren :

      -tarief per kilometer vanaf 01/01/2025 = €0,4290
      -tarief per kilometer vanaf 01/01/2026 = €0,4326
      -tarief per kilometer vanaf 01/07/2026 = €0,4440

Dit tarief omvat ALLE kosten (parkeerkosten, carwash, keuring, verzekering, afschrijving, ...). Je kan m.a.w. GEEN facturen op naam van Starterslabo laten opstellen m.b.t. tot autokosten, noch zullen deze type kosten in onkostennota aanvaard worden.

Je kan maximum 24.000 km indienen gedurende jouw traject. Noteer de totaal ingediende kilometers in het voorziene vakje: 'Totaal aantal km (traject)'

Elke maand copieer je in de excell de tab van de vorige maand en pas je de naam van de tab aan naar de huidige maand en pas je ook de data in deze tab aan. Je zorgt dat het totaal aantal kilometers voor je gehele traject klopt. Indien nodig pas je het tarief aan.

De kilometervergoeding uploaden in het portaal via 'Aankoop'

Onderwerp: KM+ jaartal + maand + jouw naam - Voorbeeld: KM 22 01 Jansen Jan

## Files

| Path | Role |
|------|------|
| User workbook (`KM-vergoeding*.xlsx`) | Working copy (month tabs `MM_YY`) — private |
| Tab `Blanco` | Source for first month |
| Tab `MM_YY` | One month of trips |
| `STARTERSLABO_KM_DIR` or `km_dir` in `starterslabo.yaml` | Optional path hint |

**Data vs skill:** keep the filled Excel outside this repo. Do not store a second filled workbook under the skill tree.

Dependency: `openpyxl` (`scripts/requirements.txt`). Portal upload reuses **starterslabo-costs**.

## Workflow

```
Progress:
- [ ] 1. Locate workbook + previous month / Blanco
- [ ] 2. Gather header (once) + trips (all grey fields)
- [ ] 3. Confirm draft table with user
- [ ] 4. Copy tab → MM_YY, write trips, rate, traject total
- [ ] 5. Print verification (month km, €, traject)
- [ ] 6. Only if asked: dry-run portal upload as onkostennota
```

### 1. Locate & inspect

Ask for the workbook path, or resolve `STARTERSLABO_KM_DIR` / `km_dir` from `starterslabo.yaml`. Do not assume a machine path or a person’s filename.

```bash
python3 - <<'PY'
import os, openpyxl
from pathlib import Path
folder = Path(os.environ.get("STARTERSLABO_KM_DIR") or ".")
for p in sorted(folder.glob("KM-vergoeding*.xlsx")) + sorted(folder.glob("*km*.xlsx")):
    wb = openpyxl.load_workbook(p, read_only=True)
    print(p, "→", wb.sheetnames)
    wb.close()
PY
```

- First month: copy **`Blanco`** → `MM_YY`.
- Later months: copy **previous** `MM_YY`, clear trip rows, keep header unless changed.
- Do not overwrite an existing month tab without asking (`--force`).

### 2. Gather from user (ask what’s missing)

**Header** (once per workbook, grey `C1`–`C11`): naam, adres, locatie, tel, email, merk+type, plaat, brandstof, CO2, % VU, IBAN.

**Each trip** (required grey fields):

1. **Datum**
2. **Contact**
3. **Actie**
4. **Locatie + adres** (straat + nr + postcode + gemeente)
5. **Aant. KM**

Heen/terug: **one row per day** with total km heen+terug unless the user asks to split.

Also ask when relevant: km-stand begin (`D54`) / einde (`D55`); electric or lichte vracht → rekening `614317` + often `% VU` 0%; else `614324` and **`% VU` from Verklaring op eer** (confirm with the coach).

**Indienen:** einde van de maand via Aankoop als er voldoende budget is; anders km bijhouden tot er wel budget is.

### 3. Confirm

Show a compact table (date | contact | actie | locatie | km) + month total + estimated € + traject running total. Only write Excel after confirmation, unless the user already gave final data and asked to fill.

### 4. Write Excel

```bash
python3 -m venv .venv && .venv/bin/pip install -r scripts/requirements.txt

.venv/bin/python scripts/fill_km_tab.py \
  --workbook /path/to/KM-vergoeding.xlsx \
  --tab 10_26 \
  --from-tab Blanco \
  --trips /tmp/km_trips.json
```

Update existing month without re-copy: add `--in-place`. Replace: `--force`.

Script sets `E12` (last day of month), `E13` (periode), `F56`/`E56` (rate), `D56` (traject = sum of all `MM_YY` tabs unless `traject_total_km` overridden).

Cell map + JSON shape: [reference.md](reference.md). Placeholder: [trips.example.json](trips.example.json).

### 5. Verify

Re-open the tab; print trip count, `F53` month km, `D56` traject, `F58` amount. Tell the user the path + tab name.

### 6. Portal upload (optional — only when asked)

Same form as onkostennota (**Aankoop → Nieuwe onkostennota**), skill **starterslabo-costs**:

| Field | Value |
|-------|-------|
| Leverancier | `KO <Achternaam> <Voornaam>` (already on file for onkosten) |
| BTW | Geen |
| Onderwerp | `KM <JJ> <MM> <Achternaam> <Voornaam>` |
| Factuurnummer | `Km <JJ> <MM>` |
| Datum | Last day of month |
| Rekening | `614317` or `614324` (confirm) |
| Bedrag | Excel total (`F58`) |
| Bijlage | The filled `.xlsx` |

```bash
uv run python /path/to/starterslabo-costs/create_draft_expense.py \
  --title "KM 26 10 Jansen Jan" \
  --account 614324 \
  --amount "7,99" \
  --date 2026-10-31 \
  --invoice-nr "Km 26 10" \
  --description "Kilometervergoeding oktober 2026 conform rittenadministratie Excel." \
  --attachment /path/to/KM-vergoeding.xlsx \
  --dry-run
```

Default: concept only. Never `--send` unless the user explicitly asks.

## Safety / privacy

- Do **not** invent contact, actie, or addresses — ask or leave TODO and stop before save.
- Do **not** invent auto/onkosten outside the forfait (parking etc. are included in the rate).
- Keep previous month tabs intact (audit trail).
- Do **not** commit filled `*.xlsx`, real trip JSON, `.env`, or credentials.
- Cap: warn if traject km would exceed **24.000**.

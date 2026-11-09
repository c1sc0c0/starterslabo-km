# starterslabo-km

Unofficial [Cursor](https://cursor.com) / [Claude](https://claude.ai) skill that
helps fill the Starters Labo **kilometervergoeding** Excel (rittenadministratie):
copy last month’s tab, write trips, apply ambtenaren rates, track traject km.

Portal upload of the filled sheet is a separate **onkostennota** (skill
[starterslabo-costs](https://github.com/c1sc0c0/starterslabo-costs)).

Not affiliated with Starterslabo. You bring your own workbook from the portal FAQ.
This Excel-fill step needs **no** username/password.

## Install

Clone this repo **as the skill folder** (`SKILL.md` at the root):

```bash
git clone https://github.com/c1sc0c0/starterslabo-km.git .cursor/skills/starterslabo-km
git clone https://github.com/c1sc0c0/starterslabo-km.git ~/.cursor/skills/starterslabo-km
git clone https://github.com/c1sc0c0/starterslabo-km.git .claude/skills/starterslabo-km
git clone https://github.com/c1sc0c0/starterslabo-km.git ~/.claude/skills/starterslabo-km
```

Restart the agent chat so it picks up the skill.

**Claude.ai:** zip this folder and upload it as a custom skill.

## Use

Point the agent at your workbook (or set `STARTERSLABO_KM_DIR`), then:

> Vul de kilometervergoeding voor oktober 2026. Vorige maand staat in tab `09_26`.

The agent will interview for grey-field trip facts, show a draft table, then run:

```bash
python3 -m venv .venv
.venv/bin/pip install -r scripts/requirements.txt

.venv/bin/python scripts/fill_km_tab.py \
  --workbook /path/to/your-KM-vergoeding.xlsx \
  --tab 10_26 \
  --from-tab 09_26 \
  --trips ./trips.json
```

See `reference.md` and `trips.example.json`. Never commit a real filled workbook or real trip file.

## Privacy

- No credentials in this repo (none required for the Excel step).
- `.gitignore` excludes `*.xlsx`, `trips*.json`, `profile*.json`, `.env`.
- Do not paste real addresses, plates, or IBAN into public issues/PRs.
- The official blank Excel is Starterslabo’s; this skill does not republish it.

## Related

Index: [starterslabo-skills](https://github.com/c1sc0c0/starterslabo-skills).

- [starterslabo-costs](https://github.com/c1sc0c0/starterslabo-costs) — onkostennota (KM upload)
- [starterslabo-faq](https://github.com/c1sc0c0/starterslabo-faq) — portal FAQ (rates, Verklaring op eer)

## License

MIT for this skill’s code and instructions. Starterslabo’s website, portal,
and Excel templates remain theirs.

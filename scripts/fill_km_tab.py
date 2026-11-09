#!/usr/bin/env python3
"""Copy / create a month tab in a Starters Labo kilometervergoeding workbook and fill trips."""

from __future__ import annotations

import argparse
import calendar
import json
import re
import sys
from copy import copy
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Official template layout (KM-vergoeding 07-26 VL BR.xlsx)
HEADER_CELLS = {
    "name": "C1",
    "address": "C2",
    "location": "C3",
    "phone": "C4",
    "email": "C5",
    "vehicle": "C6",
    "plate": "C7",
    "fuel": "C8",
    "co2": "C9",
    "vu_pct": "C10",
    "iban": "C11",
}
DATE_CELL = "E12"  # indieningsdatum (laatste dag van de maand)
PERIOD_CELL = "E13"  # periode label, e.g. "oktober 2026"
TRIP_START_ROW = 17
TRIP_END_ROW = 52
TRIP_COLS = ("B", "C", "D", "E", "F")  # datum, contact, actie, locatie+adres, km
MONTH_TOTAL_FORMULA = "F53"  # =SUM(F16:F52)
ODO_START_CELL = "D54"
ODO_END_CELL = "D55"
TRAJECT_TOTAL_CELL = "D56"
RATE_LABEL_CELL = "E56"
RATE_CELL = "F56"
AMOUNT_FORMULA = "F58"  # =F53*F56

# Forfait ambtenaren (FAQ)
RATES = (
    (date(2026, 7, 1), Decimal("0.4440"), "KM vergoeding vanaf 01/07/2026"),
    (date(2026, 1, 1), Decimal("0.4326"), "KM vergoeding vanaf 01/01/2026"),
    (date(2025, 1, 1), Decimal("0.4290"), "KM vergoeding vanaf 01/01/2025"),
)

MONTH_TAB_RE = re.compile(r"^(\d{2})_(\d{2})$")
MAX_TRAJECT_KM = Decimal("24000")


def parse_date(value: str | date | datetime | None) -> date:
    if value is None or value == "":
        raise ValueError("date is required")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"unrecognized date: {value!r}")


def month_tab_name(d: date) -> str:
    return f"{d.month:02d}_{d.year % 100:02d}"


def last_day_of_month(d: date) -> date:
    return date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def dutch_month(d: date) -> str:
    names = (
        "",
        "januari",
        "februari",
        "maart",
        "april",
        "mei",
        "juni",
        "juli",
        "augustus",
        "september",
        "oktober",
        "november",
        "december",
    )
    return f"{names[d.month]} {d.year}"


def rate_for(d: date) -> tuple[Decimal, str]:
    for since, rate, label in RATES:
        if d >= since:
            return rate, label
    raise ValueError(f"no rate defined for {d.isoformat()}")


def copy_cell_style(src, dst) -> None:
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def clear_trips(ws) -> None:
    for row in range(TRIP_START_ROW, TRIP_END_ROW + 1):
        for col in TRIP_COLS:
            ws[f"{col}{row}"] = None


def sum_trip_km(ws) -> Decimal:
    total = Decimal("0")
    for row in range(TRIP_START_ROW, TRIP_END_ROW + 1):
        val = ws[f"F{row}"].value
        if val is None or val == "":
            continue
        total += Decimal(str(val))
    return total


def sum_traject_from_other_tabs(wb, exclude: str) -> Decimal:
    total = Decimal("0")
    for name in wb.sheetnames:
        if name == exclude or name in ("Blanco", "Blad1"):
            continue
        if not MONTH_TAB_RE.match(name):
            continue
        ws = wb[name]
        total += sum_trip_km(ws)
    return total


def write_trips(ws, trips: list[dict]) -> Decimal:
    if len(trips) > (TRIP_END_ROW - TRIP_START_ROW + 1):
        raise ValueError(
            f"too many trips ({len(trips)}); max {TRIP_END_ROW - TRIP_START_ROW + 1} rows"
        )
    clear_trips(ws)
    month_km = Decimal("0")
    for i, trip in enumerate(trips):
        row = TRIP_START_ROW + i
        d = parse_date(trip.get("date") or trip.get("datum"))
        contact = (trip.get("contact") or "").strip()
        actie = (trip.get("actie") or trip.get("action") or "").strip()
        locatie = (
            trip.get("locatie")
            or trip.get("location")
            or trip.get("adres")
            or ""
        ).strip()
        km = trip.get("km") if trip.get("km") is not None else trip.get("kilometers")
        if km is None:
            raise ValueError(f"trip {i + 1} missing km")
        km_dec = Decimal(str(km))
        if not contact or not actie or not locatie:
            raise ValueError(
                f"trip {i + 1} needs contact, actie, and locatie+adres (grey fields)"
            )
        ws[f"B{row}"] = d  # datetime.date — Excel serial
        ws[f"B{row}"].number_format = "dd/mm/yyyy"  # match official Blanco template
        ws[f"C{row}"] = contact
        ws[f"D{row}"] = actie
        ws[f"E{row}"] = locatie
        ws[f"E{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws[f"F{row}"] = float(km_dec)
        ws[f"F{row}"].number_format = "0.00"
        month_km += km_dec
    return month_km


def apply_header(ws, header: dict | None) -> None:
    if not header:
        return
    for key, coord in HEADER_CELLS.items():
        if key in header and header[key] is not None:
            ws[coord] = header[key]


def ensure_formulas(ws) -> None:
    if not ws[MONTH_TOTAL_FORMULA].value:
        ws[MONTH_TOTAL_FORMULA] = "=SUM(F16:F52)"
    if not ws[AMOUNT_FORMULA].value:
        ws[AMOUNT_FORMULA] = "=F53*F56"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--workbook", required=True, help="Path to KM-vergoeding .xlsx")
    ap.add_argument(
        "--tab",
        help="Target month tab MM_YY (default: from first trip date / --period)",
    )
    ap.add_argument(
        "--from-tab",
        help="Source tab to copy (default: latest MM_YY, else Blanco)",
    )
    ap.add_argument(
        "--trips",
        required=True,
        help='JSON file: {"header": {...}, "trips": [...], "odo_start": "...", ...}',
    )
    ap.add_argument(
        "--period",
        help="Period YYYY-MM (sets E13 + rate + indieningsdatum). Default: from trips.",
    )
    ap.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing --tab (re-copy from source)",
    )
    ap.add_argument(
        "--in-place",
        action="store_true",
        help="Update an existing --tab without re-copying (keep header as-is unless JSON header set)",
    )
    args = ap.parse_args()

    path = Path(args.workbook)
    if not path.exists():
        print(f"ERROR: workbook not found: {path}", file=sys.stderr)
        return 1

    payload = json.loads(Path(args.trips).read_text(encoding="utf-8"))
    if isinstance(payload, list):
        payload = {"trips": payload}
    trips = payload.get("trips") or []
    if not trips:
        print("ERROR: trips JSON must include a non-empty trips array", file=sys.stderr)
        return 1

    header = payload.get("header") or {}
    period_month: date | None = None
    if args.period:
        y, m = args.period.split("-", 1)
        period_month = date(int(y), int(m), 1)
    else:
        period_month = parse_date(trips[0].get("date") or trips[0].get("datum")).replace(
            day=1
        )

    tab = args.tab or month_tab_name(period_month)
    submit_date = last_day_of_month(period_month)
    rate, rate_label = rate_for(submit_date)

    wb = load_workbook(path)

    if tab in wb.sheetnames and args.in_place and not args.force:
        ws = wb[tab]
    else:
        if tab in wb.sheetnames:
            if not args.force:
                print(
                    f"ERROR: tab {tab!r} already exists (pass --force or --in-place)",
                    file=sys.stderr,
                )
                return 1
            del wb[tab]

        src_name = args.from_tab
        if not src_name:
            month_tabs = [n for n in wb.sheetnames if MONTH_TAB_RE.match(n)]
            src_name = month_tabs[-1] if month_tabs else (
                "Blanco" if "Blanco" in wb.sheetnames else wb.sheetnames[0]
            )
        if src_name not in wb.sheetnames:
            print(
                f"ERROR: source tab {src_name!r} missing; have {wb.sheetnames}",
                file=sys.stderr,
            )
            return 1

        ws = wb.copy_worksheet(wb[src_name])
        ws.title = tab
        # Place after source / before Blad1 helpers when possible
        try:
            target_idx = wb.sheetnames.index(src_name) + 1
            current_idx = wb.sheetnames.index(tab)
            wb.move_sheet(ws, offset=target_idx - current_idx)
        except Exception:
            pass

    apply_header(ws, header)
    ensure_formulas(ws)

    # Clear leftover sample date from Blanco if present
    if ws["B17"].value == datetime(1900, 1, 16):
        ws["B17"] = None

    month_km = write_trips(ws, trips)

    ws[DATE_CELL] = submit_date
    ws[DATE_CELL].number_format = "dd/mm/yyyy"
    ws[PERIOD_CELL] = payload.get("periode") or dutch_month(period_month)
    ws[RATE_LABEL_CELL] = rate_label
    ws[RATE_CELL] = float(rate)
    ws[RATE_CELL].number_format = "0.0000"

    if "odo_start" in payload:
        ws[ODO_START_CELL] = payload["odo_start"]
    if "odo_end" in payload:
        ws[ODO_END_CELL] = payload["odo_end"]

    other = sum_traject_from_other_tabs(wb, exclude=tab)
    traject = other + month_km
    if "traject_total_km" in payload:
        traject = Decimal(str(payload["traject_total_km"]))
    ws[TRAJECT_TOTAL_CELL] = float(traject)
    ws[TRAJECT_TOTAL_CELL].number_format = "0.00"

    if traject > MAX_TRAJECT_KM:
        print(
            f"WARNING: traject total {traject} km exceeds max {MAX_TRAJECT_KM} km",
            file=sys.stderr,
        )

    amount = (month_km * rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    wb.save(path)

    print(f"workbook: {path}")
    print(f"tab: {tab} (from period {period_month.isoformat()[:7]})")
    print(f"trips: {len(trips)}")
    print(f"month_km: {month_km}")
    print(f"traject_km: {traject} (other tabs {other})")
    print(f"rate: {rate} → amount ≈ €{amount}")
    print(f"indiening: {submit_date.isoformat()}")
    print(f"onderwerp hint: KM {period_month.year % 100:02d} {period_month.month:02d} <Achternaam> <Voornaam>")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

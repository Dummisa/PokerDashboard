"""One-time migration: data/Pokerbilanz.xlsx -> docs/data/games.json

The website reads docs/data/games.json directly; new games are added through
the entry form (docs/eintragen.html). This script is only needed if you want
to re-import the old spreadsheet.

Usage:  python scripts/convert_excel.py
"""
from datetime import datetime, date
from pathlib import Path
import json

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "Pokerbilanz.xlsx"
OUT = ROOT / "docs" / "data" / "games.json"


def parse_date(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError(f"Unbekanntes Datum: {v!r}")


def num(v):
    if v is None or str(v).strip() == "":
        return None
    if isinstance(v, (int, float)):
        f = float(v)
    else:
        f = float(str(v).replace(",", ".").replace("'", ""))
    f = round(f, 2)
    return int(f) if f.is_integer() else f


def dump_games(games):
    """One game per line -> readable git diffs."""
    lines = [json.dumps(g, ensure_ascii=False, separators=(", ", ": ")) for g in games]
    return '{"games": [\n' + ",\n".join(lines) + "\n]}\n"


def main():
    ws = openpyxl.load_workbook(SRC, data_only=False)["Tabelle1"]
    header = [c.value for c in ws[1]]
    date_idx = next(i for i, h in enumerate(header) if str(h).strip().lower() == "datum")
    players = [(i, str(h).strip()) for i, h in enumerate(header[:date_idx]) if h]

    games = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = row[date_idx]
        if d is None:
            continue
        results = {}
        for i, name in players:
            v = num(row[i])
            if v is not None:
                results[name] = v
        if results:
            games.append({"date": parse_date(d), "results": results})

    games.sort(key=lambda g: g["date"])
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(dump_games(games), encoding="utf-8")
    print(f"{len(games)} Spiele -> {OUT}")
    for g in games:
        s = round(sum(g["results"].values()), 2)
        if abs(s) > 0.009:
            print(f"  Achtung: {g['date']} Summe = {s}")


if __name__ == "__main__":
    main()

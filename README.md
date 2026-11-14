# Poker Statistik

Statische Website (GitHub Pages aus `docs/`) mit Bestenliste, Verlauf, Spieler-Profilen, Duell-Vergleich und Gruppen-Statistiken.

## Neues Spiel eintragen
Auf der Website oben rechts **＋ Spiel** (`eintragen.html`):
1. Einmalig pro Gerät einen GitHub *fine-grained token* einfügen (nur Repo `PokerDashboard`, Permission *Contents: Read and write*). Er bleibt nur in diesem Browser gespeichert.
2. Spieler antippen, Beträge eingeben (± Knopf für Verlust), die Summe sollte 0 sein.
3. **Speichern** → committet `docs/data/games.json`, GitHub Pages ist nach ca. 1–2 Minuten aktualisiert.

Spiele lassen sich unter „Letzte Spiele“ bearbeiten oder löschen.

## Daten
- `docs/data/games.json`: einzige Datenquelle, ein Spiel pro Zeile. Kann auch direkt auf GitHub bearbeitet werden.
- `data/Pokerbilanz.xlsx`: altes Archiv. `python scripts/convert_excel.py` importiert es erneut (überschreibt `games.json`!).

## Aufbau
- `docs/index.html` + `docs/assets/app.js`: Dashboard, alle Statistiken werden im Browser berechnet (kein Build-Schritt).
- `docs/assets/stats.js`: gemeinsame Berechnungen (Statistik, Twint-Zahlungen, Datenformat).
- `docs/eintragen.html` + `docs/assets/entry.js`: Eingabeformular (GitHub Contents API).
- `docs/vendor/chart.umd.min.js`: Chart.js 4.6.1.

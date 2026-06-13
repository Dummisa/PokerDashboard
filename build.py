from pathlib import Path
from datetime import datetime
import json

import pandas as pd
import plotly.graph_objects as go
from jinja2 import Environment, FileSystemLoader, select_autoescape

ROOT = Path(__file__).parent.resolve()
DATA_XLSX = ROOT / "data" / "Pokerbilanz.xlsx"
DOCS_DIR = ROOT / "docs"
TEMPLATES_DIR = ROOT / "templates"

def find_date_col(cols) -> str:
    for c in cols:
        if str(c).strip().lower() == "datum":
            return c
    raise ValueError("Konnte die Spalte 'Datum' nicht finden. Stelle sicher, dass die letzte Spalte 'Datum' heißt.")

def read_pokerbilanz(path: Path) -> tuple[pd.DataFrame, str, list[str]]:
    if not path.exists():
        raise FileNotFoundError(f"Excel nicht gefunden: {path}. Lege die Datei unter data/Pokerbilanz.xlsx ab.")

    df_raw = pd.read_excel(path, header=0, engine="openpyxl")
    if df_raw.empty:
        raise ValueError("Die gelesene Tabelle ist leer.")

    date_col = find_date_col(df_raw.columns)
    date_idx = df_raw.columns.get_loc(date_col)
    df = df_raw.iloc[:, : date_idx + 1].copy()

    # drop the second physical row (irrelevant numbers)
    if len(df) >= 1:
        df = df.iloc[1:].reset_index(drop=True)

    # remove blank/"Unnamed" columns except the date
    keep = []
    for c in df.columns:
        if c == date_col:
            keep.append(c); continue
        s = str(c)
        if s.strip() == "" or s.lower().startswith("unnamed"):
            continue
        keep.append(c)
    df = df[keep]

    player_cols = [c for c in df.columns if c != date_col]

    # numeric coercion
    for p in player_cols:
        df[p] = pd.to_numeric(df[p], errors="coerce")

    # parse date dd.mm.yyyy
    df[date_col] = pd.to_datetime(df[date_col], format="%d.%m.%Y", errors="coerce")

    # drop fully empty player columns
    player_cols = [p for p in player_cols if not df[p].isna().all()]
    df = df[player_cols + [date_col]]

    # drop fully empty rows (no date and all NaN)
    mask_all_nan = df[player_cols].isna().all(axis=1) if player_cols else pd.Series(False, index=df.index)
    df = df[~(df[date_col].isna() & mask_all_nan)].copy()

    # sort by date
    df = df.sort_values(date_col, kind="mergesort").reset_index(drop=True)

    return df, date_col, player_cols

def make_cumulative_and_markers(df: pd.DataFrame, date_col: str, player_cols: list[str]):
    """Return cumulated dataframe and per-player marker size arrays (0 when original value is NaN).
       Cumulative values before a player's first game are set to NaN (not plotted)."""
    cum = pd.DataFrame({date_col: df[date_col]})
    marker_sizes = {}
    for p in player_cols:
        vals = df[p]
        cumvals = vals.fillna(0).cumsum()
        first_idx = vals.first_valid_index()
        if first_idx is not None:
            cumvals.iloc[:first_idx] = None
        else:
            cumvals[:] = None
        cum[p] = cumvals
        marker_sizes[p] = [6 if pd.notna(v) else 0 for v in vals]
    return cum, marker_sizes

def sort_players_by_attendance(df: pd.DataFrame, player_cols: list[str], last_n: int = 10) -> list[str]:
    """Sort players by attendance in the last `last_n` games (desc), then alphabetically."""
    tail = df.tail(last_n)
    attendance = {}
    for p in player_cols:
        attendance[p] = tail[p].notna().sum()
    return sorted(player_cols, key=lambda p: (-attendance[p], str(p).lower()))

def compute_payments(df: pd.DataFrame, date_col: str, player_cols: list[str]) -> list[dict]:
    """Compute minimal Twint payments for the last game.
    Returns list of {'from': str, 'to': str, 'amount': float}."""
    if df.empty:
        return []

    last_row = df.iloc[-1]
    # Collect non-null deltas from last game
    balances = {}
    for p in player_cols:
        v = last_row[p]
        if pd.notna(v) and v != 0:
            balances[str(p)] = float(v)

    if not balances:
        return []

    # Split into debtors (negative = must pay) and creditors (positive = receive)
    debtors = []   # (amount_owed, player)  — positive numbers
    creditors = [] # (amount_due, player)   — positive numbers
    for p, v in balances.items():
        if v < 0:
            debtors.append([-v, p])
        elif v > 0:
            creditors.append([v, p])

    # Greedy: always settle the largest debtor against the largest creditor
    debtors.sort(reverse=True)
    creditors.sort(reverse=True)

    payments = []
    di, ci = 0, 0
    while di < len(debtors) and ci < len(creditors):
        amount = min(debtors[di][0], creditors[ci][0])
        amount = round(amount, 2)
        if amount > 0:
            payments.append({
                "from": debtors[di][1],
                "to": creditors[ci][1],
                "amount": amount,
            })
        debtors[di][0] -= amount
        creditors[ci][0] -= amount
        if debtors[di][0] < 0.005:
            di += 1
        if creditors[ci][0] < 0.005:
            ci += 1

    # Sort: group by payer, then by amount descending
    payments.sort(key=lambda t: (t["from"].lower(), -t["amount"]))
    return payments

def make_hist_chart(df: pd.DataFrame, cum: pd.DataFrame, markers: dict, date_col: str, player_cols: list[str]) -> str:
    fig = go.Figure()
    x = df[date_col]
    for p in player_cols:
        fig.add_trace(go.Scatter(
            x=x,
            y=cum[p],
            mode="lines+markers",
            name=str(p),
            marker=dict(size=markers[p]),
            connectgaps=False,
            hovertemplate="%{x|%d.%m.%Y}<br>%{y}<extra>" + str(p) + "</extra>",
        ))
    fig.update_layout(
        title=None,
        margin=dict(l=40, r=10, t=10, b=40),
        legend_title_text="Spieler",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.15,
            xanchor="left",
            x=0,
            font=dict(size=11),
        ),
        hovermode="closest",
        height=500,
        dragmode="pan",
    )
    fig.update_xaxes(
        title_text="Datum",
        fixedrange=False,
    )
    fig.update_yaxes(
        title_text="Kumulierte Gewinne/Verluste",
    )
    html = fig.to_html(
        full_html=False,
        include_plotlyjs="cdn",
        div_id="historical_chart",
        config={
            "responsive": True,
            "displaylogo": False,
            "scrollZoom": True,
            "modeBarButtonsToRemove": [
                "lasso2d", "select2d",
                "zoom2d",
                "zoomIn2d", "zoomOut2d",
                "autoScale2d",
            ],
        }
    )
    return html

def render_page(context: dict) -> str:
    env = Environment(
        loader=FileSystemLoader(TEMPLATES_DIR),
        autoescape=select_autoescape(["html", "xml"]),
    )
    template = env.get_template("index.html.j2")
    return template.render(**context)

def write_site(html: str) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    out = DOCS_DIR / "index.html"
    out.write_text(html, encoding="utf-8")
    return out

def main():
    df, date_col, player_cols = read_pokerbilanz(DATA_XLSX)
    player_cols = sort_players_by_attendance(df, player_cols)
    cum, marker_sizes = make_cumulative_and_markers(df, date_col, player_cols)
    chart_hist = make_hist_chart(df, cum, marker_sizes, date_col, player_cols)
    payments = compute_payments(df, date_col, player_cols)

    # JSON payloads for client-side leaderboard logic
    dates_iso = df[date_col].dt.strftime("%Y-%m-%d").tolist()
    dates_labels = df[date_col].dt.strftime("%d.%m.%Y").tolist()
    last_game_date = dates_labels[-1] if dates_labels else ""

    deltas = {str(p): [ (None if pd.isna(v) else float(v)) for v in df[p].tolist() ] for p in player_cols}
    cums   = {str(p): [ (None if pd.isna(v) else float(v)) for v in cum[p].tolist() ] for p in player_cols}

    select_size = max(6, min(12, len(player_cols))) if player_cols else 6
    context = {
        "first_date": dates_iso[0] if dates_iso else "",
        "last_date": dates_iso[-1] if dates_iso else "",
        "default_visible": 5,
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "chart_hist": chart_hist,
        "payments": payments,
        "last_game_date": last_game_date,
        "players": [str(p) for p in player_cols],
        "players_json": json.dumps([str(p) for p in player_cols], ensure_ascii=False),
        "select_size": select_size,
        "dates_iso_json": json.dumps(dates_iso, ensure_ascii=False),
        "dates_labels_json": json.dumps(dates_labels, ensure_ascii=False),
        "deltas_json": json.dumps(deltas, ensure_ascii=False),
        "cum_json": json.dumps(cums, ensure_ascii=False),
        "last_index": len(dates_iso) - 1 if len(dates_iso) else 0,
    }

    page_html = render_page(context)
    out = write_site(page_html)
    print(f"Built site → {out}")

if __name__ == "__main__":
    main()

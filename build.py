from pathlib import Path
from datetime import datetime

import pandas as pd
import plotly.express as px
from jinja2 import Environment, FileSystemLoader, select_autoescape

ROOT = Path(__file__).parent.resolve()
DATA_XLSX = ROOT / "data" / "data.xlsx"
DOCS_DIR = ROOT / "docs"
TEMPLATES_DIR = ROOT / "templates"

def ensure_sample_excel(path: Path):
    """Create a tiny example Excel if none exists yet."""
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame({
        "date": pd.date_range("2024-12-30", periods=10, freq="D"),
        "value": [10, 13, 11, 15, 12, 18, 17, 22, 21, 25],
    })
    df.to_excel(path, index=False)

def read_data(path: Path) -> pd.DataFrame:
    return pd.read_excel(path)

def make_chart(df: pd.DataFrame) -> str:
    # Basic interactive line chart
    fig = px.line(df, x="date", y="value", markers=True, title="Hello Dashboard (from Excel)")
    # Output Plotly as an embeddable HTML fragment (JS via CDN)
    html = fig.to_html(full_html=False, include_plotlyjs="cdn", config={"responsive": True})
    return html

def render_page(chart_html: str) -> str:
    env = Environment(
        loader=FileSystemLoader(TEMPLATES_DIR),
        autoescape=select_autoescape(["html", "xml"]),
    )
    template = env.get_template("index.html.j2")
    return template.render(
        title="My Excel Dashboard",
        last_updated=datetime.now().strftime("%Y-%m-%d %H:%M"),
        chart_html=chart_html,
    )

def write_site(html: str):
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    out = DOCS_DIR / "index.html"
    out.write_text(html, encoding="utf-8")
    return out

def main():
    ensure_sample_excel(DATA_XLSX)   # creates data/data.xlsx if you don't have one yet
    df = read_data(DATA_XLSX)
    chart_html = make_chart(df)
    page_html = render_page(chart_html)
    out = write_site(page_html)
    print(f"Built site → {out}")

if __name__ == "__main__":
    main()
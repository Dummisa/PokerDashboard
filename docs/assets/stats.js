// Shared data + statistics helpers (used by dashboard and entry form)
const PS = (() => {
  const DATA_URL = 'data/games.json';

  async function loadGames() {
    const res = await fetch(`${DATA_URL}?t=${Date.now()}`, { cache: 'no-store' });
    if (!res.ok) throw new Error(`Daten konnten nicht geladen werden (${res.status})`);
    const json = await res.json();
    return normalize(json.games || []);
  }

  function normalize(games) {
    return games
      .filter(g => g && g.date && g.results)
      .map(g => ({ date: g.date, results: { ...g.results } }))
      .sort((a, b) => a.date.localeCompare(b.date));
  }

  // ── Formatting ──
  const nf = new Intl.NumberFormat('de-CH', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  const nf2 = new Intl.NumberFormat('de-CH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const round2 = v => Math.round(v * 100) / 100;
  function fmt(v, { sign = false, fixed = false } = {}) {
    if (v == null || Number.isNaN(v)) return '–';
    const r = round2(v);
    const s = (fixed ? nf2 : nf).format(Math.abs(r));
    if (r < 0) return '−' + s;
    return (sign && r > 0 ? '+' : '') + s;
  }
  const cls = v => (v == null || Math.abs(v) < 0.005) ? '' : (v > 0 ? 'pos' : 'neg');
  const pct = v => v == null ? '–' : `${Math.round(v)}%`;
  function dateLabel(iso, short = false) {
    const [y, m, d] = iso.split('-');
    return short ? `${d}.${m}.${y.slice(2)}` : `${d}.${m}.${y}`;
  }
  const MONTHS = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'];
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // ── Stats ──
  function summarize(values) {
    const v = values.filter(x => x != null);
    const n = v.length;
    if (!n) return { n: 0, total: 0, mean: null, median: null, sd: null, max: null, min: null, wins: 0, losses: 0, winRate: null };
    const total = v.reduce((a, b) => a + b, 0);
    const mean = total / n;
    const s = [...v].sort((a, b) => a - b);
    const median = n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
    const sd = n > 1 ? Math.sqrt(v.reduce((a, x) => a + (x - mean) ** 2, 0) / (n - 1)) : null;
    const wins = v.filter(x => x > 0).length;
    const losses = v.filter(x => x < 0).length;
    return { n, total, mean, median, sd, max: s[n - 1], min: s[0], wins, losses, winRate: wins / n * 100 };
  }

  function longestStreak(values, cond) {
    let best = 0, cur = 0;
    for (const v of values) {
      if (cond(v)) { cur++; best = Math.max(best, cur); } else cur = 0;
    }
    return best;
  }

  /** Model over a list of games. */
  function model(games) {
    const players = new Map(); // name -> {name, games: [{idx, date, v}]}
    games.forEach((g, idx) => {
      for (const [name, v] of Object.entries(g.results)) {
        if (!players.has(name)) players.set(name, { name, games: [] });
        players.get(name).games.push({ idx, date: g.date, v });
      }
    });
    return { games, players };
  }

  /** Players sorted by attendance in the last `n` games, then total games, then name. */
  function playersByRecent(games, n = 20) {
    const recent = new Map(), all = new Map();
    games.forEach((g, i) => {
      for (const p of Object.keys(g.results)) {
        all.set(p, (all.get(p) || 0) + 1);
        if (i >= games.length - n) recent.set(p, (recent.get(p) || 0) + 1);
      }
    });
    return [...all.keys()].sort((a, b) =>
      (recent.get(b) || 0) - (recent.get(a) || 0) || all.get(b) - all.get(a) || a.localeCompare(b, 'de'));
  }

  /** Minimal set of payments to settle one game. */
  function payments(results) {
    const debt = [], cred = [];
    for (const [p, v] of Object.entries(results)) {
      if (v < 0) debt.push([-v, p]); else if (v > 0) cred.push([v, p]);
    }
    debt.sort((a, b) => b[0] - a[0]); cred.sort((a, b) => b[0] - a[0]);
    const out = [];
    let i = 0, j = 0;
    while (i < debt.length && j < cred.length) {
      const amt = round2(Math.min(debt[i][0], cred[j][0]));
      if (amt > 0) out.push({ from: debt[i][1], to: cred[j][1], amount: amt });
      debt[i][0] -= amt; cred[j][0] -= amt;
      if (debt[i][0] < 0.005) i++;
      if (cred[j][0] < 0.005) j++;
    }
    return out.sort((a, b) => a.from.localeCompare(b.from, 'de') || b.amount - a.amount);
  }

  // Stable colours per player
  const PALETTE = ['#6366f1', '#ef4444', '#10b981', '#f59e0b', '#06b6d4', '#a855f7', '#ec4899', '#84cc16',
    '#f97316', '#3b82f6', '#14b8a6', '#eab308', '#8b5cf6', '#f43f5e', '#22c55e', '#0ea5e9'];
  function colorMap(games) {
    const counts = new Map();
    for (const g of games) for (const p of Object.keys(g.results)) counts.set(p, (counts.get(p) || 0) + 1);
    const order = [...counts.keys()].sort((a, b) => counts.get(b) - counts.get(a) || a.localeCompare(b, 'de'));
    const m = new Map();
    order.forEach((p, i) => m.set(p, PALETTE[i % PALETTE.length]));
    return m;
  }

  // Same on-disk format as scripts/convert_excel.py (one game per line → readable diffs)
  function serialize(games) {
    const line = g => `{"date": ${JSON.stringify(g.date)}, "results": {` +
      Object.entries(g.results).map(([k, v]) => `${JSON.stringify(k)}: ${round2(v)}`).join(', ') + '}}';
    return '{"games": [\n' + games.map(line).join(',\n') + '\n]}\n';
  }

  return { loadGames, normalize, fmt, cls, pct, dateLabel, MONTHS, esc, round2, summarize, longestStreak,
    model, playersByRecent, payments, colorMap, serialize };
})();

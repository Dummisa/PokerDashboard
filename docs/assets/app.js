(() => {
  const { fmt, cls, pct, dateLabel, MONTHS, esc, summarize, longestStreak, round2 } = PS;
  const $ = s => document.querySelector(s);
  const store = {
    get(k, d) { try { const v = localStorage.getItem('ps.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { localStorage.setItem('ps.' + k, JSON.stringify(v)); } catch { /* ignore */ } },
  };

  let GAMES = [], M, COLORS, ORDER, YEARS = [];
  const charts = {};
  const rendered = {};
  const state = {
    lbPeriod: 'all', lbSort: 'total', lbAsc: false, lbAll: false,
    histPeriod: 'all', histSel: null,
    player: null, monthYear: null,
    duelA: null, duelB: null,
    game: null,
  };

  // ── Theme-aware chart defaults ──
  function css(v) { return getComputedStyle(document.documentElement).getPropertyValue(v).trim(); }
  function applyChartTheme() {
    Chart.defaults.color = css('--muted');
    Chart.defaults.borderColor = css('--border');
    Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
    Chart.defaults.font.size = 11;
    Chart.defaults.plugins.legend.display = false;
    Chart.defaults.plugins.tooltip.backgroundColor = css('--card-2');
    Chart.defaults.plugins.tooltip.titleColor = css('--text');
    Chart.defaults.plugins.tooltip.bodyColor = css('--text');
    Chart.defaults.plugins.tooltip.borderColor = css('--border');
    Chart.defaults.plugins.tooltip.borderWidth = 1;
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.maintainAspectRatio = false;
  }
  function makeChart(id, config) {
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart(document.getElementById(id), config);
    return charts[id];
  }

  // ── Periods ──
  function periods(withLast = true) {
    const p = [{ key: 'all', label: 'Gesamt' }];
    YEARS.slice().reverse().forEach(y => p.push({ key: 'y' + y, label: String(y) }));
    if (withLast) p.push({ key: 'last10', label: 'Letzte 10' }, { key: 'last25', label: 'Letzte 25' });
    return p;
  }
  function gamesIn(period) {
    if (period === 'all') return GAMES;
    if (period.startsWith('last')) return GAMES.slice(-Number(period.slice(4)));
    const y = period.slice(1);
    return GAMES.filter(g => g.date.startsWith(y));
  }
  function renderChips(el, items, active, onPick) {
    el.innerHTML = items.map(i => `<button class="chip${i.key === active ? ' active' : ''}" data-k="${esc(i.key)}">${esc(i.label)}</button>`).join('');
    el.onclick = e => {
      const b = e.target.closest('.chip'); if (!b) return;
      el.querySelectorAll('.chip').forEach(c => c.classList.toggle('active', c === b));
      onPick(b.dataset.k);
    };
  }
  const dot = p => `<span class="dot" style="background:${COLORS.get(p) || '#888'}"></span>`;
  const playerLink = p => `<a href="#spieler/${encodeURIComponent(p)}" class="player-cell">${dot(p)}${esc(p)}</a>`;

  // ═════════ Übersicht ═════════
  function renderOverview() {
    const sel = $('#gameSelect');
    sel.innerHTML = GAMES.map((g, i) => `<option value="${i}">${dateLabel(g.date)}</option>`).reverse().join('');
    sel.value = state.game ?? GAMES.length - 1;
    sel.onchange = () => { state.game = +sel.value; renderGame(); };
    renderGame();

    renderChips($('#lbPeriods'), periods(), state.lbPeriod, k => { state.lbPeriod = k; renderLeaderboard(); });
    const mg = $('#minGames');
    mg.value = store.get('minGames', '1');
    mg.onchange = () => { store.set('minGames', mg.value); renderLeaderboard(); };
    document.querySelectorAll('#leaderboard th').forEach(th => th.onclick = () => {
      const k = th.dataset.k; if (k === 'rank') return;
      if (state.lbSort === k) state.lbAsc = !state.lbAsc;
      else { state.lbSort = k; state.lbAsc = k === 'name'; }
      renderLeaderboard();
    });
    $('#lbMore').onclick = () => { state.lbAll = !state.lbAll; renderLeaderboard(); };
    renderLeaderboard();
  }

  function renderGame() {
    const g = GAMES[+$('#gameSelect').value];
    const res = Object.entries(g.results).sort((a, b) => b[1] - a[1]);
    const sum = round2(res.reduce((a, [, v]) => a + v, 0));
    $('#gameResults').innerHTML = res.map(([p, v], i) =>
      `<div class="list-row"><span class="muted small" style="width:1.2em">${i + 1}</span><span class="grow">${playerLink(p)}</span><span class="num ${cls(v)}" style="font-weight:600">${fmt(v, { sign: true })}</span></div>`).join('') +
      (Math.abs(sum) >= 0.01 ? `<div class="hint">Differenz in diesem Spiel: ${fmt(sum, { sign: true })}</div>` : '');
    const pays = PS.payments(g.results);
    $('#gamePayments').innerHTML = pays.length ? pays.map(t =>
      `<div class="list-row"><span class="from">${esc(t.from)}</span><span class="arrow">→</span><span class="to">${esc(t.to)}</span><span class="amt">CHF ${fmt(t.amount, { fixed: true })}</span></div>`).join('')
      : '<div class="muted small">Keine Zahlungen nötig.</div>';
  }

  function leaderboardRows(period) {
    const games = gamesIn(period);
    const per = new Map();
    for (const g of games) for (const [p, v] of Object.entries(g.results)) {
      if (!per.has(p)) per.set(p, []);
      per.get(p).push(v);
    }
    return [...per.entries()].map(([name, vals]) => ({ name, ...summarize(vals) }));
  }

  function renderLeaderboard() {
    const min = +$('#minGames').value;
    const rows = leaderboardRows(state.lbPeriod).filter(r => r.n >= min);
    const k = state.lbSort, dir = state.lbAsc ? 1 : -1;
    rows.sort((a, b) => {
      if (k === 'name') return dir * a.name.localeCompare(b.name, 'de');
      const av = a[k] ?? Infinity * dir, bv = b[k] ?? Infinity * dir;
      return dir * (av - bv) || b.total - a.total;
    });
    const medals = k === 'total' && !state.lbAsc;
    document.querySelectorAll('#leaderboard th').forEach(th => {
      th.classList.toggle('sorted', th.dataset.k === k);
      th.classList.toggle('asc', th.dataset.k === k && state.lbAsc);
    });
    const LIMIT = 15;
    $('#lbMore').hidden = rows.length <= LIMIT;
    $('#lbMore').textContent = state.lbAll ? 'Weniger anzeigen' : `Alle ${rows.length} anzeigen`;
    $('#leaderboard tbody').innerHTML = (state.lbAll ? rows : rows.slice(0, LIMIT)).map((r, i) => `<tr>
      <td class="rank">${medals && i < 3 ? `<span class="medal">${['🥇', '🥈', '🥉'][i]}</span>` : i + 1}</td>
      <td class="name">${playerLink(r.name)}</td>
      <td class="num ${cls(r.total)}" style="font-weight:700">${fmt(r.total, { sign: true })}</td>
      <td class="num">${r.n}</td>
      <td class="num ${cls(r.mean)}">${fmt(r.mean, { sign: true })}</td>
      <td class="num">${fmt(r.sd)}</td>
      <td class="num">${pct(r.winRate)}</td>
      <td class="num pos">${fmt(r.max, { sign: true })}</td>
      <td class="num neg">${fmt(r.min, { sign: true })}</td>
    </tr>`).join('') || '<tr><td colspan="9" class="muted">Keine Spieler mit so vielen Spielen in diesem Zeitraum.</td></tr>';
  }

  // ═════════ Verlauf ═════════
  function renderHistory() {
    renderChips($('#histPeriods'), periods(), state.histPeriod, k => { state.histPeriod = k; drawHistory(); });
    if (!state.histSel) state.histSel = new Set(store.get('histSel', null) || ORDER.slice(0, 5));
    const el = $('#histPlayers');
    el.onclick = e => {
      const b = e.target.closest('.chip'); if (!b) return;
      const p = b.dataset.p;
      state.histSel.has(p) ? state.histSel.delete(p) : state.histSel.add(p);
      saveSel(); drawHistory();
    };
    $('#histTop').onclick = () => { state.histSel = new Set(ORDER.slice(0, 5)); saveSel(); drawHistory(); };
    $('#histAll').onclick = () => { state.histSel = new Set(playersIn(gamesIn(state.histPeriod))); saveSel(); drawHistory(); };
    $('#histNone').onclick = () => { state.histSel = new Set(); saveSel(); drawHistory(); };
    drawHistory();
  }
  const saveSel = () => store.set('histSel', [...state.histSel]);
  function playersIn(games) {
    const s = new Set(); games.forEach(g => Object.keys(g.results).forEach(p => s.add(p)));
    return ORDER.filter(p => s.has(p));
  }

  function cumulativeSeries(games, p) {
    let acc = 0, started = false;
    return games.map(g => {
      const v = g.results[p];
      if (v != null) { acc += v; started = true; }
      return started ? round2(acc) : null;
    });
  }

  function drawHistory() {
    const games = gamesIn(state.histPeriod);
    const avail = playersIn(games);
    $('#histPlayers').innerHTML = avail.map(p =>
      `<button class="chip player${state.histSel.has(p) ? ' active' : ''}" data-p="${esc(p)}"><span class="dot" style="background:${COLORS.get(p)}"></span>${esc(p)}</button>`).join('');
    const shown = avail.filter(p => state.histSel.has(p));
    const many = games.length > 60;
    makeChart('histChart', {
      type: 'line',
      data: {
        labels: games.map(g => dateLabel(g.date, true)),
        datasets: shown.map(p => ({
          label: p,
          data: cumulativeSeries(games, p),
          borderColor: COLORS.get(p), backgroundColor: COLORS.get(p),
          borderWidth: 2, tension: .15, spanGaps: false,
          pointRadius: games.map(g => g.results[p] != null ? (many ? 0 : 2.5) : 0),
          pointHoverRadius: games.map(g => g.results[p] != null ? 5 : 0),
        })),
      },
      options: {
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 6 }, grid: { display: false } },
          y: { ticks: { callback: v => fmt(v) }, grid: { color: ctx => ctx.tick?.value === 0 ? css('--muted') : css('--border') } },
        },
        plugins: {
          tooltip: {
            itemSort: (a, b) => b.parsed.y - a.parsed.y,
            filter: i => i.parsed.y != null,
            callbacks: {
              label: c => {
                const d = games[c.dataIndex].results[c.dataset.label];
                return ` ${c.dataset.label}: ${fmt(c.parsed.y, { sign: true })}` + (d != null ? `  (${fmt(d, { sign: true })})` : '');
              },
            },
          },
        },
      },
    });
  }

  // ═════════ Spieler ═════════
  function fillPlayerSelect(sel, value) {
    sel.innerHTML = ORDER.map(p => `<option value="${esc(p)}">${esc(p)}</option>`).join('');
    sel.value = value;
  }

  function renderPlayer() {
    const sel = $('#playerSelect');
    if (!state.player || !M.players.has(state.player)) state.player = ORDER[0];
    fillPlayerSelect(sel, state.player);
    sel.onchange = () => { location.hash = 'spieler/' + encodeURIComponent(sel.value); };
    const p = state.player;
    const pg = M.players.get(p).games;
    const vals = pg.map(x => x.v);
    const s = summarize(vals);

    // Rank by all-time total
    const totals = leaderboardRows('all').sort((a, b) => b.total - a.total);
    const rank = totals.findIndex(r => r.name === p) + 1;
    const firstIdx = pg[0].idx;
    const attendance = pg.length / (GAMES.length - firstIdx) * 100;
    const maxW = pg.reduce((a, x) => x.v > (a?.v ?? -Infinity) ? x : a, null);
    const maxL = pg.reduce((a, x) => x.v < (a?.v ?? Infinity) ? x : a, null);
    const last = pg[pg.length - 1];
    const last10 = summarize(vals.slice(-10));
    const streakW = longestStreak(vals, v => v > 0), streakL = longestStreak(vals, v => v < 0);
    // current streak
    let cur = 0; const lastSign = Math.sign(last.v);
    for (let i = vals.length - 1; i >= 0 && Math.sign(vals[i]) === lastSign && lastSign !== 0; i--) cur++;

    const k = (label, value, sub = '', c = '') => `<div class="kpi"><div class="label">${label}</div><div class="value ${c}">${value}</div><div class="sub">${sub}</div></div>`;
    $('#playerKpis').innerHTML = [
      k('Total', fmt(s.total, { sign: true }), `Rang ${rank} von ${totals.length}`, cls(s.total)),
      k('Spiele', s.n, `${pct(attendance)} Teilnahme seit ${dateLabel(pg[0].date, true)}`),
      k('Ø pro Spiel', fmt(s.mean, { sign: true }), `Median ${fmt(s.median, { sign: true })}`, cls(s.mean)),
      k('Standardabw. <span class="lc">σ</span>', fmt(s.sd), s.sd != null ? `typisch ${fmt(s.mean - s.sd, { sign: true })} … ${fmt(s.mean + s.sd, { sign: true })}` : ''),
      k('Win-Rate', pct(s.winRate), `${s.wins} S · ${s.losses} N · ${s.n - s.wins - s.losses} ±0`),
      k('Form (10)', fmt(last10.total, { sign: true }), `Ø ${fmt(last10.mean, { sign: true })} in den letzten ${last10.n}`, cls(last10.total)),
      k('Max. Gewinn', fmt(maxW.v, { sign: true }), `am ${dateLabel(maxW.date)}`, cls(maxW.v)),
      k('Max. Verlust', fmt(maxL.v, { sign: true }), `am ${dateLabel(maxL.date)}`, cls(maxL.v)),
      k('Serien', `${streakW} / ${streakL}`, `längste Sieg- / Verlustserie`),
      k('Aktuell', cur ? `${cur}× ${lastSign > 0 ? 'Plus' : 'Minus'}` : '–', 'laufende Serie'),
      k('Letztes Spiel', fmt(last.v, { sign: true }), `am ${dateLabel(last.date)}`, cls(last.v)),
    ].join('');

    // Cumulative chart (own games only)
    let acc = 0;
    const cum = pg.map(x => round2(acc += x.v));
    const col = COLORS.get(p);
    makeChart('playerCumChart', {
      type: 'line',
      data: { labels: pg.map(x => dateLabel(x.date, true)), datasets: [{ data: cum, borderColor: col, backgroundColor: col + '22', fill: true, borderWidth: 2, pointRadius: 0, pointHoverRadius: 4, tension: .15 }] },
      options: {
        interaction: { mode: 'index', intersect: false },
        scales: { x: { ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 5 }, grid: { display: false } }, y: { ticks: { callback: v => fmt(v) } } },
        plugins: { tooltip: { callbacks: { label: c => ` Total ${fmt(c.parsed.y, { sign: true })}  (Spiel ${fmt(pg[c.dataIndex].v, { sign: true })})` } } },
      },
    });

    // Form chart
    const form = pg.slice(-20);
    barChart('playerFormChart', form.map(x => dateLabel(x.date, true)), form.map(x => x.v));

    // Per month
    const years = [...new Set(pg.map(x => x.date.slice(0, 4)))].sort();
    if (!state.monthYear || !years.includes(state.monthYear)) state.monthYear = years[years.length - 1];
    renderChips($('#playerMonthYears'), years.slice().reverse().map(y => ({ key: y, label: y })), state.monthYear, y => { state.monthYear = y; drawMonths(pg); });
    drawMonths(pg);

    // Per year table (with rank within that year)
    $('#playerYears tbody').innerHTML = years.slice().reverse().map(y => {
      const ys = summarize(pg.filter(x => x.date.startsWith(y)).map(x => x.v));
      const yr = leaderboardRows('y' + y).sort((a, b) => b.total - a.total);
      const r = yr.findIndex(x => x.name === p) + 1;
      return `<tr><td><strong>${y}</strong></td><td class="num">${ys.n}</td><td class="num ${cls(ys.total)}" style="font-weight:700">${fmt(ys.total, { sign: true })}</td>
        <td class="num ${cls(ys.mean)}">${fmt(ys.mean, { sign: true })}</td><td class="num">${fmt(ys.sd)}</td><td class="num">${pct(ys.winRate)}</td><td class="num">${r}/${yr.length}</td></tr>`;
    }).join('');
  }

  function drawMonths(pg) {
    const sums = Array(12).fill(null);
    pg.filter(x => x.date.startsWith(state.monthYear)).forEach(x => {
      const m = +x.date.slice(5, 7) - 1;
      sums[m] = round2((sums[m] || 0) + x.v);
    });
    barChart('playerMonthChart', MONTHS, sums);
  }

  function barChart(id, labels, data) {
    const pos = css('--pos'), neg = css('--neg');
    makeChart(id, {
      type: 'bar',
      data: { labels, datasets: [{ data, backgroundColor: data.map(v => v >= 0 ? pos : neg), borderRadius: 4, maxBarThickness: 28 }] },
      options: {
        scales: { x: { ticks: { maxRotation: 0, autoSkip: true }, grid: { display: false } }, y: { ticks: { callback: v => fmt(v) } } },
        plugins: { tooltip: { callbacks: { label: c => ' ' + fmt(c.parsed.y, { sign: true }) } } },
      },
    });
  }

  // ═════════ Duell ═════════
  function renderDuel() {
    if (!state.duelA || !M.players.has(state.duelA)) state.duelA = ORDER[0];
    if (!state.duelB || !M.players.has(state.duelB) || state.duelB === state.duelA) state.duelB = ORDER.find(p => p !== state.duelA);
    const a = state.duelA, b = state.duelB;
    fillPlayerSelect($('#duelA'), a); fillPlayerSelect($('#duelB'), b);
    const go = () => { location.hash = `duell/${encodeURIComponent($('#duelA').value)}/${encodeURIComponent($('#duelB').value)}`; };
    $('#duelA').onchange = go; $('#duelB').onchange = go;

    const shared = GAMES.filter(g => g.results[a] != null && g.results[b] != null);
    let aAhead = 0, bAhead = 0, sa = 0, sb = 0;
    const ca = [], cb = [];
    for (const g of shared) {
      const va = g.results[a], vb = g.results[b];
      if (va > vb) aAhead++; else if (vb > va) bAhead++;
      ca.push(round2(sa += va)); cb.push(round2(sb += vb));
    }
    const k = (label, value, sub = '', c = '') => `<div class="kpi"><div class="label">${label}</div><div class="value ${c}">${value}</div><div class="sub">${sub}</div></div>`;
    $('#duelKpis').innerHTML = [
      k('Gemeinsame Spiele', shared.length, shared.length ? `seit ${dateLabel(shared[0].date, true)}` : ''),
      k('Öfter vorne', aAhead === bAhead ? 'Gleichstand' : esc(aAhead > bAhead ? a : b), `${esc(a)} ${aAhead} : ${bAhead} ${esc(b)}`),
      k(`${esc(a)} gemeinsam`, fmt(sa, { sign: true }), shared.length ? `Ø ${fmt(sa / shared.length, { sign: true })}` : '', cls(sa)),
      k(`${esc(b)} gemeinsam`, fmt(sb, { sign: true }), shared.length ? `Ø ${fmt(sb / shared.length, { sign: true })}` : '', cls(sb)),
    ].join('');

    const colA = COLORS.get(a), colB = COLORS.get(b) === colA ? '#888888' : COLORS.get(b);
    makeChart('duelChart', {
      type: 'line',
      data: {
        labels: shared.map(g => dateLabel(g.date, true)),
        datasets: [
          { label: a, data: ca, borderColor: colA, backgroundColor: colA, borderWidth: 2, pointRadius: shared.length > 40 ? 0 : 2.5, tension: .15 },
          { label: b, data: cb, borderColor: colB, backgroundColor: colB, borderWidth: 2, pointRadius: shared.length > 40 ? 0 : 2.5, tension: .15 },
        ],
      },
      options: {
        interaction: { mode: 'index', intersect: false },
        scales: { x: { ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 5 }, grid: { display: false } }, y: { ticks: { callback: v => fmt(v) } } },
        plugins: { legend: { display: true, labels: { boxWidth: 10, boxHeight: 10, usePointStyle: true } } },
      },
    });

    const sA = summarize(M.players.get(a).games.map(x => x.v)), sB = summarize(M.players.get(b).games.map(x => x.v));
    const rows = [
      ['Total', 'total', true, v => fmt(v, { sign: true })], ['Spiele', 'n', true, v => v],
      ['Ø pro Spiel', 'mean', true, v => fmt(v, { sign: true })], ['Median', 'median', true, v => fmt(v, { sign: true })],
      ['Standardabw. σ', 'sd', null, v => fmt(v)], ['Win-Rate', 'winRate', true, pct],
      ['Max. Gewinn', 'max', true, v => fmt(v, { sign: true })], ['Max. Verlust', 'min', true, v => fmt(v, { sign: true })],
    ];
    $('#duelTable').innerHTML = `<tr><td style="font-weight:700">${playerLink(a)}</td><td></td><td style="font-weight:700">${playerLink(b)}</td></tr>` +
      rows.map(([label, key, higherBetter, f]) => {
        const va = sA[key], vb = sB[key];
        const ba = higherBetter && va != null && vb != null && va > vb, bb = higherBetter && va != null && vb != null && vb > va;
        return `<tr><td class="num ${ba ? 'better' : ''}" style="text-align:center">${f(va)}</td><td class="lbl">${label}</td><td class="num ${bb ? 'better' : ''}" style="text-align:center">${f(vb)}</td></tr>`;
      }).join('');
  }

  // ═════════ Gruppe ═════════
  function renderGroup() {
    const nPlayers = M.players.size;
    const sizes = GAMES.map(g => Object.keys(g.results).length);
    const volumes = GAMES.map(g => Object.values(g.results).reduce((a, v) => a + (v > 0 ? v : 0), 0));
    const avg = a => a.reduce((x, y) => x + y, 0) / a.length;
    const thisYear = YEARS[YEARS.length - 1];
    const bigNight = volumes.indexOf(Math.max(...volumes));
    const mostGames = ORDER.slice().sort((a, b) => M.players.get(b).games.length - M.players.get(a).games.length)[0];
    const k = (label, value, sub = '') => `<div class="kpi"><div class="label">${label}</div><div class="value">${value}</div><div class="sub">${sub}</div></div>`;
    $('#groupKpis').innerHTML = [
      k('Spielabende', GAMES.length, `seit ${dateLabel(GAMES[0].date)}`),
      k(`Abende ${thisYear}`, GAMES.filter(g => g.date.startsWith(thisYear)).length, ''),
      k('Spieler', nPlayers, 'insgesamt dabei gewesen'),
      k('Ø Spieler / Abend', fmt(avg(sizes)), `max. ${Math.max(...sizes)}`),
      k('Ø umverteilt / Abend', 'CHF ' + fmt(avg(volumes), { fixed: false }), 'Summe aller Gewinne'),
      k('Grösster Abend', 'CHF ' + fmt(volumes[bigNight]), `am ${dateLabel(GAMES[bigNight].date)}`),
      k('Meiste Spiele', esc(mostGames), `${M.players.get(mostGames).games.length} Abende`),
      k('Ø Streuung <span class="lc">σ</span>', fmt(avg(leaderboardRows('all').filter(r => r.n >= 10).map(r => r.sd))), 'Spieler mit ≥ 10 Spielen'),
    ].join('');

    // Records
    let bestWin = null, worstLoss = null;
    for (const g of GAMES) for (const [p, v] of Object.entries(g.results)) {
      if (!bestWin || v > bestWin.v) bestWin = { p, v, d: g.date };
      if (!worstLoss || v < worstLoss.v) worstLoss = { p, v, d: g.date };
    }
    let bestYear = null, worstYear = null;
    for (const y of YEARS) for (const r of leaderboardRows('y' + y)) {
      if (!bestYear || r.total > bestYear.total) bestYear = { ...r, y };
      if (!worstYear || r.total < worstYear.total) worstYear = { ...r, y };
    }
    let streakW = { n: 0 }, attend = { n: 0 };
    for (const [p, info] of M.players) {
      const n = longestStreak(info.games.map(x => x.v), v => v > 0);
      if (n > streakW.n) streakW = { p, n };
      let best = 0, cur = 0, prev = -2;
      for (const x of info.games) { cur = x.idx === prev + 1 ? cur + 1 : 1; prev = x.idx; best = Math.max(best, cur); }
      if (best > attend.n) attend = { p, n: best };
    }
    const bestMean = leaderboardRows('all').filter(r => r.n >= 20).sort((a, b) => b.mean - a.mean)[0];
    const steady = leaderboardRows('all').filter(r => r.n >= 20).sort((a, b) => a.sd - b.sd)[0];
    const wild = leaderboardRows('all').filter(r => r.n >= 20).sort((a, b) => b.sd - a.sd)[0];
    const rec = (icon, label, p, val, sub) => `<div class="list-row"><span>${icon}</span><span class="grow"><span class="muted small">${label}</span><br>${playerLink(p)} <span class="muted small">${sub}</span></span><span class="num" style="font-weight:700">${val}</span></div>`;
    $('#groupRecords').innerHTML = [
      rec('💰', 'Höchster Gewinn an einem Abend', bestWin.p, `<span class="pos">${fmt(bestWin.v, { sign: true })}</span>`, dateLabel(bestWin.d)),
      rec('💸', 'Höchster Verlust an einem Abend', worstLoss.p, `<span class="neg">${fmt(worstLoss.v, { sign: true })}</span>`, dateLabel(worstLoss.d)),
      rec('📈', 'Bestes Jahr', bestYear.name, `<span class="pos">${fmt(bestYear.total, { sign: true })}</span>`, bestYear.y),
      rec('📉', 'Schlechtestes Jahr', worstYear.name, `<span class="neg">${fmt(worstYear.total, { sign: true })}</span>`, worstYear.y),
      bestMean ? rec('🎯', 'Bester Schnitt (≥ 20 Spiele)', bestMean.name, `<span class="${cls(bestMean.mean)}">Ø ${fmt(bestMean.mean, { sign: true })}</span>`, '') : '',
      rec('🔥', 'Längste Siegesserie', streakW.p, `${streakW.n}×`, ''),
      rec('📅', 'Meiste Abende am Stück', attend.p, `${attend.n}×`, ''),
      steady ? rec('🧊', 'Konstantester Spieler (kleinste σ, ≥ 20 Spiele)', steady.name, `σ ${fmt(steady.sd)}`, '') : '',
      wild ? rec('🎢', 'Wildester Spieler (grösste σ, ≥ 20 Spiele)', wild.name, `σ ${fmt(wild.sd)}`, '') : '',
    ].join('');

    // Year charts
    const perYear = YEARS.map(y => GAMES.map((g, i) => [g, i]).filter(([g]) => g.date.startsWith(y)));
    const accent = css('--accent');
    makeChart('groupYearChart', {
      type: 'bar',
      data: { labels: YEARS, datasets: [{ data: perYear.map(a => a.length), backgroundColor: accent, borderRadius: 5, maxBarThickness: 40 }] },
      options: { scales: { x: { grid: { display: false } }, y: { beginAtZero: true, ticks: { precision: 0 } } } },
    });
    makeChart('groupNightChart', {
      data: {
        labels: YEARS,
        datasets: [
          { type: 'bar', label: 'Ø umverteilt (CHF)', data: perYear.map(a => round2(avg(a.map(([, i]) => volumes[i])))), backgroundColor: css('--pos') + 'aa', borderRadius: 5, maxBarThickness: 40, yAxisID: 'y' },
          { type: 'line', label: 'Ø Spieler', data: perYear.map(a => round2(avg(a.map(([, i]) => sizes[i])))), borderColor: accent, backgroundColor: accent, borderWidth: 2, yAxisID: 'y2', tension: .2 },
        ],
      },
      options: {
        interaction: { mode: 'index', intersect: false },
        scales: {
          x: { grid: { display: false } },
          y: { beginAtZero: true, position: 'left', ticks: { callback: v => fmt(v) } },
          y2: { beginAtZero: true, position: 'right', grid: { display: false } },
        },
        plugins: { legend: { display: true, labels: { boxWidth: 10, boxHeight: 10, usePointStyle: true } } },
      },
    });

    // Month heatmap
    const counts = {};
    let maxC = 1;
    GAMES.forEach(g => { const key = g.date.slice(0, 7); counts[key] = (counts[key] || 0) + 1; maxC = Math.max(maxC, counts[key]); });
    const shade = (n, max) => n ? `background:color-mix(in srgb, var(--accent) ${Math.round(15 + 70 * n / max)}%, transparent);${n / max > .55 ? 'color:var(--accent-text);font-weight:600' : ''}` : 'background:var(--card-2);color:transparent';
    $('#groupHeat').innerHTML = `<thead><tr><th></th>${MONTHS.map(m => `<th>${m}</th>`).join('')}<th>Σ</th></tr></thead><tbody>` +
      YEARS.slice().reverse().map(y => `<tr><td class="label">${y}</td>${MONTHS.map((_, m) => {
        const n = counts[`${y}-${String(m + 1).padStart(2, '0')}`] || 0;
        return `<td style="${shade(n, maxC)}">${n || '·'}</td>`;
      }).join('')}<td class="label" style="text-align:center">${GAMES.filter(g => g.date.startsWith(y)).length}</td></tr>`).join('') + '</tbody>';

    // Attendance per year
    const top = [...M.players.values()].sort((a, b) => b.games.length - a.games.length).slice(0, 15);
    let maxA = 1;
    const att = top.map(info => YEARS.map(y => { const n = info.games.filter(x => x.date.startsWith(y)).length; maxA = Math.max(maxA, n); return n; }));
    $('#groupAttendance').innerHTML = `<thead><tr><th></th>${YEARS.map(y => `<th>${String(y).slice(2)}</th>`).join('')}<th>Σ</th></tr></thead><tbody>` +
      top.map((info, i) => `<tr><td class="label name">${playerLink(info.name)}</td>${att[i].map(n => `<td style="${shade(n, maxA)}">${n || '·'}</td>`).join('')}<td class="label" style="text-align:center">${info.games.length}</td></tr>`).join('') + '</tbody>';
  }

  // ═════════ Routing ═════════
  const VIEWS = { uebersicht: renderOverview, verlauf: renderHistory, spieler: renderPlayer, duell: renderDuel, gruppe: renderGroup };
  function route() {
    const [view, ...args] = (location.hash.slice(1) || 'uebersicht').split('/').map(decodeURIComponent);
    const v = VIEWS[view] ? view : 'uebersicht';
    if (v === 'spieler' && args[0]) state.player = args[0];
    if (v === 'duell' && args[0]) { state.duelA = args[0]; state.duelB = args[1] || null; }
    document.querySelectorAll('.view').forEach(s => s.classList.toggle('active', s.id === 'view-' + v));
    document.querySelectorAll('#nav a').forEach(a => a.classList.toggle('active', a.dataset.view === v));
    // Spieler/Duell depend on args; others render once
    if (!rendered[v] || v === 'spieler' || v === 'duell') { VIEWS[v](); rendered[v] = true; }
    if (args.length || !rendered.scrolled) window.scrollTo({ top: 0 });
    rendered.scrolled = true;
  }

  async function init() {
    try {
      GAMES = await PS.loadGames();
    } catch (e) {
      $('#loading').textContent = e.message; return;
    }
    $('#loading').remove();
    M = PS.model(GAMES);
    COLORS = PS.colorMap(GAMES);
    ORDER = PS.playersByRecent(GAMES);
    YEARS = [...new Set(GAMES.map(g => g.date.slice(0, 4)))].sort();
    applyChartTheme();
    $('#footer').innerHTML = `${GAMES.length} Spielabende · letztes Spiel ${dateLabel(GAMES[GAMES.length - 1].date)} · <a href="eintragen.html">Spiel eintragen</a>`;
    window.addEventListener('hashchange', route);
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
      applyChartTheme();
      Object.keys(rendered).forEach(k => delete rendered[k]);
      route();
    });
    route();
  }
  init();
})();

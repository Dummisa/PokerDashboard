(() => {
  const REPO = { owner: 'Dummisa', repo: 'PokerDashboard', branch: 'main', path: 'docs/data/games.json' };
  const API = 'https://api.github.com';
  const { fmt, cls, dateLabel, esc, round2 } = PS;
  const $ = s => document.querySelector(s);
  const store = {
    get(k, d) { try { const v = localStorage.getItem('ps.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { v == null ? localStorage.removeItem('ps.' + k) : localStorage.setItem('ps.' + k, JSON.stringify(v)); } catch { /* ignore */ } },
  };

  let token = store.get('token', null);
  let GAMES = [];
  let PLAYERS = [];
  // form state
  let rows = [];          // [{name, sign: 1|-1, text}]
  let editDate = null;    // date of game being edited (null = new game)
  let showAllPlayers = false;

  // ── GitHub API ──
  async function gh(path, opts = {}) {
    const res = await fetch(API + path, {
      ...opts,
      headers: {
        Accept: 'application/vnd.github+json',
        Authorization: `Bearer ${token}`,
        'X-GitHub-Api-Version': '2022-11-28',
        ...(opts.body ? { 'Content-Type': 'application/json' } : {}),
      },
      cache: 'no-store',
    });
    if (!res.ok) {
      let msg = '';
      try { msg = (await res.json()).message; } catch { /* ignore */ }
      const err = new Error(msg || res.statusText);
      err.status = res.status;
      throw err;
    }
    return res.json();
  }
  const b64encode = str => {
    const bytes = new TextEncoder().encode(str);
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  };
  const b64decode = b64 => new TextDecoder().decode(Uint8Array.from(atob(b64.replace(/\s/g, '')), c => c.charCodeAt(0)));
  const filePath = () => `/repos/${REPO.owner}/${REPO.repo}/contents/${REPO.path}`;

  async function readRemote() {
    const f = await gh(`${filePath()}?ref=${REPO.branch}`);
    return { sha: f.sha, games: PS.normalize(JSON.parse(b64decode(f.content)).games || []) };
  }
  async function writeRemote(games, sha, message) {
    return gh(filePath(), {
      method: 'PUT',
      body: JSON.stringify({ message, content: b64encode(PS.serialize(games)), sha, branch: REPO.branch }),
    });
  }
  function explain(e) {
    if (e.status === 401) return 'Token ungültig oder abgelaufen. Bitte neu verbinden.';
    if (e.status === 403) return 'Keine Schreibrechte. Hat der Token „Contents: Read and write“ für PokerDashboard?';
    if (e.status === 404) return 'Repository oder Datei nicht gefunden. Hat der Token Zugriff auf PokerDashboard?';
    return e.message || 'Unbekannter Fehler';
  }
  const notice = (el, type, html) => { $(el).innerHTML = html ? `<div class="notice ${type}">${html}</div>` : ''; };

  // ── Connection ──
  async function connect(t) {
    token = t;
    notice('#connMsg', '', '');
    try {
      const [user] = await Promise.all([gh('/user').catch(() => null), readRemote()]);
      store.set('token', t);
      showConnected(user && user.login);
      await loadGames();
    } catch (e) {
      token = null;
      notice('#connMsg', 'err', esc(explain(e)));
    }
  }
  function showConnected(login) {
    $('#connSetup').hidden = true;
    $('#connDone').hidden = false;
    $('#connStatus').innerHTML = '<span class="pos">● verbunden</span>';
    $('#connWho').textContent = login ? `Angemeldet als ${login}` : 'Token gespeichert';
  }
  function showDisconnected() {
    $('#connSetup').hidden = false;
    $('#connDone').hidden = true;
    $('#connStatus').textContent = 'nicht verbunden';
  }
  $('#tokenSave').onclick = () => { const t = $('#tokenInput').value.trim(); if (t) connect(t); };
  $('#tokenInput').onkeydown = e => { if (e.key === 'Enter') $('#tokenSave').click(); };
  $('#tokenRemove').onclick = () => {
    if (!confirm('Token von diesem Gerät entfernen?')) return;
    store.set('token', null); token = null; showDisconnected();
  };

  // ── Data ──
  async function loadGames() {
    try {
      GAMES = token ? (await readRemote()).games : await PS.loadGames();
    } catch {
      GAMES = await PS.loadGames().catch(() => []);
    }
    PLAYERS = PS.playersByRecent(GAMES, 15);
    renderPicker(); renderRecent();
  }

  // ── Form ──
  const today = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
  function parseNum(text) {
    const t = String(text).trim().replace(/\s|'/g, '').replace(',', '.').replace('−', '-');
    if (t === '') return null;
    if (!/^-?\d*\.?\d+$|^-?\d+\.$/.test(t)) return NaN;
    return parseFloat(t);
  }
  const rowValue = r => { const n = parseNum(r.text); return n == null || Number.isNaN(n) ? n : round2(n < 0 ? n : r.sign * n); };
  function saveDraft() { store.set('draft', { date: $('#date').value, rows, editDate }); }

  function renderPicker() {
    const inGame = new Set(rows.map(r => r.name));
    const extra = rows.map(r => r.name).filter(n => !PLAYERS.includes(n));
    const shown = showAllPlayers ? PLAYERS : PLAYERS.filter((p, i) => i < 12 || inGame.has(p));
    $('#morePlayers').hidden = PLAYERS.length <= 12;
    $('#morePlayers').textContent = showAllPlayers ? 'Weniger zeigen' : `Alle ${PLAYERS.length} Spieler zeigen`;
    $('#pickChips').innerHTML = [...shown, ...extra].map(p =>
      `<button class="chip${inGame.has(p) ? ' picked' : ''}" data-p="${esc(p)}">${inGame.has(p) ? '✓ ' : ''}${esc(p)}</button>`).join('');
  }
  $('#morePlayers').onclick = () => { showAllPlayers = !showAllPlayers; renderPicker(); };
  $('#pickChips').onclick = e => {
    const b = e.target.closest('.chip'); if (!b) return;
    togglePlayer(b.dataset.p);
  };
  function togglePlayer(name) {
    const i = rows.findIndex(r => r.name === name);
    if (i >= 0) rows.splice(i, 1);
    else rows.push({ name, sign: 1, text: '' });
    renderAll();
    if (i < 0) { const inputs = document.querySelectorAll('#rows input'); inputs[inputs.length - 1]?.focus(); }
  }
  $('#addPlayer').onclick = () => {
    const name = $('#newPlayer').value.trim().replace(/\s+/g, ' ');
    if (!name) return;
    const existing = [...PLAYERS, ...rows.map(r => r.name)].find(p => p.toLowerCase() === name.toLowerCase());
    $('#newPlayer').value = '';
    if (existing && rows.some(r => r.name === existing)) return;
    togglePlayer(existing || name);
  };
  $('#newPlayer').onkeydown = e => { if (e.key === 'Enter') $('#addPlayer').click(); };

  function renderRows() {
    $('#noRows').hidden = rows.length > 0;
    $('#rows').innerHTML = rows.map((r, i) => `
      <div class="list-row entry-row" data-i="${i}">
        <span class="pname">${esc(r.name)}</span>
        <button class="btn sign ${r.sign > 0 ? 'plus' : 'minus'}" data-act="sign" aria-label="Vorzeichen wechseln">${r.sign > 0 ? '+' : '−'}</button>
        <input type="text" inputmode="decimal" enterkeyhint="next" placeholder="0.00" value="${esc(r.text)}" data-act="val" aria-label="Betrag ${esc(r.name)}" />
        <button class="btn x" data-act="rm" aria-label="Entfernen">✕</button>
      </div>`).join('');
  }
  $('#rows').addEventListener('click', e => {
    const b = e.target.closest('button'); if (!b) return;
    const i = +b.closest('.entry-row').dataset.i;
    if (b.dataset.act === 'sign') {
      rows[i].sign *= -1;
      // if the user typed a minus, flip it with the button instead
      const n = parseNum(rows[i].text); if (n < 0) { rows[i].text = String(Math.abs(n)); rows[i].sign = 1; }
      renderAll();
    } else if (b.dataset.act === 'rm') { rows.splice(i, 1); renderAll(); }
  });
  $('#rows').addEventListener('input', e => {
    if (e.target.dataset.act !== 'val') return;
    const i = +e.target.closest('.entry-row').dataset.i;
    rows[i].text = e.target.value;
    let t = e.target.value.trim();
    if (t.startsWith('-') || t.startsWith('−')) { rows[i].sign = -1; rows[i].text = t.slice(1); e.target.value = rows[i].text; renderRows(); focusRow(i); }
    renderSummary(); saveDraft();
  });
  $('#rows').addEventListener('keydown', e => {
    if (e.key !== 'Enter' || e.target.dataset.act !== 'val') return;
    e.preventDefault();
    const inputs = [...document.querySelectorAll('#rows input')];
    const next = inputs[inputs.indexOf(e.target) + 1];
    next ? next.focus() : e.target.blur();
  });
  function focusRow(i) { const inp = document.querySelectorAll('#rows input')[i]; if (inp) { inp.focus(); inp.setSelectionRange(inp.value.length, inp.value.length); } }

  function currentResults() {
    const res = {};
    for (const r of rows) { const v = rowValue(r); if (v != null && !Number.isNaN(v)) res[r.name] = v; }
    return res;
  }
  function renderSummary() {
    const res = currentResults();
    const bad = rows.filter(r => Number.isNaN(rowValue(r)));
    const empty = rows.filter(r => rowValue(r) == null);
    const sum = round2(Object.values(res).reduce((a, b) => a + b, 0));
    $('#sum').textContent = (sum > 0 ? '+' : '') + fmt(sum);
    $('#sum').className = 'num ' + cls(sum);
    let msg = '', type = '';
    if (bad.length) { type = 'err'; msg = `Ungültige Zahl bei ${bad.map(r => esc(r.name)).join(', ')}`; }
    else if (rows.length && empty.length) { type = 'warn'; msg = `Noch kein Betrag: ${empty.map(r => esc(r.name)).join(', ')} (leer = nicht gespeichert)`; }
    else if (rows.length >= 2 && Math.abs(sum) >= 0.01) { type = 'warn'; msg = `Die Summe sollte 0 sein. Differenz: ${fmt(sum, { sign: true })}. Bitte Tippfehler prüfen.`; }
    else if (rows.length >= 2) { type = 'ok'; msg = '✓ Geht auf'; }
    notice('#sumMsg', type, msg);
    $('#overrideWrap').hidden = !(rows.length >= 2 && Math.abs(sum) >= 0.01 && !bad.length);
    const pays = PS.payments(res);
    $('#payWrap').hidden = !pays.length;
    $('#payList').innerHTML = pays.map(t => `<div class="list-row"><span class="from">${esc(t.from)}</span><span class="arrow">→</span><span class="to">${esc(t.to)}</span><span class="amt">CHF ${fmt(t.amount, { fixed: true })}</span></div>`).join('');
    const { share } = PS.imbalance(res);
    $('#payNote').innerHTML = rows.length >= 2 && Math.abs(sum) >= 0.01
      ? `Differenz gleichmässig verteilt: ${fmt(share, { sign: true, fixed: true })} pro Spieler. Gespeichert werden die eingegebenen Werte.` : '';
  }
  function renderMode() {
    $('#formTitle').textContent = editDate ? `Spiel vom ${dateLabel(editDate)} bearbeiten` : 'Neues Spiel';
    $('#cancelEdit').hidden = !editDate;
    $('#del').hidden = !editDate;
    $('#save').textContent = editDate ? 'Änderungen speichern' : 'Speichern';
  }
  function renderAll() { renderRows(); renderPicker(); renderSummary(); renderMode(); saveDraft(); }
  $('#date').onchange = saveDraft;

  function resetForm() {
    rows = []; editDate = null; $('#date').value = today(); $('#override').checked = false;
    store.set('draft', null); renderAll();
  }
  $('#cancelEdit').onclick = () => { resetForm(); notice('#saveMsg', '', ''); };

  function startEdit(date) {
    const g = GAMES.find(x => x.date === date); if (!g) return;
    editDate = date;
    $('#date').value = date;
    rows = Object.entries(g.results).map(([name, v]) => ({ name, sign: v < 0 ? -1 : 1, text: String(Math.abs(v)) }));
    notice('#saveMsg', '', '');
    renderAll();
    $('#formCard').scrollIntoView({ behavior: 'smooth' });
  }

  function renderRecent() {
    $('#recent').innerHTML = GAMES.slice(-8).reverse().map(g => {
      const res = Object.entries(g.results).sort((a, b) => b[1] - a[1]);
      return `<button class="list-row" data-d="${g.date}" style="border:0;font:inherit;color:inherit;text-align:left;cursor:pointer;width:100%">
        <strong style="white-space:nowrap">${dateLabel(g.date)}</strong>
        <span class="muted small" style="flex:1;min-width:0;text-align:left">${res.map(([p, v]) => `${esc(p)} <span class="${cls(v)}">${fmt(v, { sign: true })}</span>`).join(' · ')}</span>
        <span class="muted">✎</span></button>`;
    }).join('') || '<div class="muted small">Noch keine Spiele.</div>';
  }
  $('#recent').onclick = e => { const b = e.target.closest('[data-d]'); if (b) startEdit(b.dataset.d); };

  // ── Save / delete ──
  async function commit(mutate, message) {
    for (let attempt = 0; attempt < 2; attempt++) {
      const { sha, games } = await readRemote();
      const next = mutate(games.map(g => ({ ...g, results: { ...g.results } })));
      try {
        const res = await writeRemote(next, sha, message);
        GAMES = PS.normalize(next);
        return res;
      } catch (e) {
        if ((e.status === 409 || e.status === 422) && attempt === 0) continue; // file changed meanwhile → retry once
        throw e;
      }
    }
  }
  function busy(on) { $('#save').disabled = on; $('#del').disabled = on; }
  function success(text, res) {
    const url = res && res.commit && res.commit.html_url;
    notice('#saveMsg', 'ok', `${text} Die Statistik ist in ca. 1–2 Minuten aktualisiert.${url ? ` <a href="${url}" target="_blank" rel="noopener">Commit ansehen</a>` : ''}`);
    PLAYERS = PS.playersByRecent(GAMES, 15);
    resetForm(); renderRecent();
  }

  $('#save').onclick = async () => {
    notice('#saveMsg', '', '');
    if (!token) { notice('#saveMsg', 'err', 'Bitte zuerst oben mit GitHub verbinden.'); return; }
    const date = $('#date').value;
    const results = currentResults();
    const sum = round2(Object.values(results).reduce((a, b) => a + b, 0));
    if (!date) return notice('#saveMsg', 'err', 'Bitte ein Datum wählen.');
    if (rows.some(r => Number.isNaN(rowValue(r)))) return notice('#saveMsg', 'err', 'Bitte ungültige Beträge korrigieren.');
    if (Object.keys(results).length < 2) return notice('#saveMsg', 'err', 'Mindestens zwei Spieler mit Betrag eintragen.');
    if (Math.abs(sum) >= 0.01 && !$('#override').checked) return notice('#saveMsg', 'warn', `Summe ist ${fmt(sum, { sign: true })} statt 0. Korrigieren oder „Trotzdem speichern“ anhaken.`);

    const wasEdit = editDate;
    busy(true);
    notice('#saveMsg', '', '<span class="muted">Speichere …</span>');
    try {
      const res = await commit(games => {
        const clash = games.find(g => g.date === date && g.date !== wasEdit);
        if (clash) { const e = new Error(`Am ${dateLabel(date)} gibt es schon ein Spiel. Bearbeite es unter „Letzte Spiele“.`); e.user = true; throw e; }
        const out = wasEdit ? games.filter(g => g.date !== wasEdit) : games;
        out.push({ date, results });
        return out.sort((a, b) => a.date.localeCompare(b.date));
      }, `Spiel ${dateLabel(date)} ${wasEdit ? 'bearbeitet' : 'eingetragen'}`);
      success(wasEdit ? 'Änderungen gespeichert.' : 'Spiel gespeichert.', res);
    } catch (e) {
      notice('#saveMsg', 'err', esc(e.user ? e.message : explain(e)));
    } finally { busy(false); }
  };

  $('#del').onclick = async () => {
    if (!editDate || !confirm(`Spiel vom ${dateLabel(editDate)} wirklich löschen?`)) return;
    const d = editDate;
    busy(true);
    try {
      const res = await commit(games => games.filter(g => g.date !== d), `Spiel ${dateLabel(d)} gelöscht`);
      success('Spiel gelöscht.', res);
    } catch (e) {
      notice('#saveMsg', 'err', esc(explain(e)));
    } finally { busy(false); }
  };

  // ── Init ──
  (async () => {
    const draft = store.get('draft', null);
    if (draft && (draft.rows?.length || draft.editDate)) {
      rows = draft.rows || []; editDate = draft.editDate || null; $('#date').value = draft.date || today();
    } else $('#date').value = today();
    renderAll();
    if (token) {
      showConnected(null);
      try { const u = await gh('/user'); $('#connWho').textContent = `Angemeldet als ${u.login}`; }
      catch (e) { if (e.status === 401) { notice('#connMsg', 'err', esc(explain(e))); showDisconnected(); } }
    } else showDisconnected();
    await loadGames();
  })();
})();

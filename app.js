/* Mes habitudes — application autonome : iPhone ⇄ OneDrive, sans serveur intermédiaire. */
(function () {
  'use strict';
  const CFG = window.APP_CONFIG || {};
  const VERSION = '1.2';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const SCOPE = 'openid profile offline_access User.Read Files.ReadWrite';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const DAYS_SHORT = ['L', 'M', 'M', 'J', 'V', 'S', 'D'];
  const DAYS_LONG = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'];
  const MONTHS_LOWER = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  const DOMAIN_COLOR = { sport: 'var(--sport)', alimentation: 'var(--alim)', mental: 'var(--mental)' };

  /* ---------- Stockage local (toujours protégé) ---------- */
  const store = {
    get(k, d) { try { const v = localStorage.getItem('hab.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { if (v === undefined || v === null) localStorage.removeItem('hab.' + k); else localStorage.setItem('hab.' + k, JSON.stringify(v)); } catch {} },
  };

  /* ---------- État ---------- */
  const S = {
    clientId: store.get('clientId', '') || CFG.clientId || '',
    tenant: CFG.tenant || 'consumers',
    tokens: store.get('tokens', null),
    user: store.get('user', null),
    file: store.get('file', null), // {id, name}
    sheet: store.get('sheet', null),
    model: (store.get('model', null) || {}).version === 2 ? store.get('model', null) : null,
    loadedAt: store.get('loadedAt', 0),
    pending: (store.get('model', null) || {}).version === 2 ? store.get('pending', {}) : {},
    view: 'today', day: null, planDay: null,
    showUnplanned: false, modal: null, browse: null,
    sync: 'idle', syncMsg: '', needLogin: false,
    files: null, setupError: '', toast: '',
  };
  const now = () => new Date();
  const redirectUri = () => new URL('./', location.href.split('#')[0].split('?')[0]).href;
  const authBase = () => `https://login.microsoftonline.com/${S.tenant}/oauth2/v2.0`;

  /* ---------- Authentification Microsoft (OAuth 2 + PKCE, sans bibliothèque) ---------- */
  const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const randomStr = (n = 48) => b64url(crypto.getRandomValues(new Uint8Array(n)));

  async function startLogin(silent) {
    const verifier = randomStr(48);
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const state = randomStr(16);
    store.set('pkce', { verifier, state, silent: !!silent, at: Date.now() });
    const p = new URLSearchParams({
      client_id: S.clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
      scope: SCOPE, state, code_challenge: challenge, code_challenge_method: 'S256',
    });
    if (silent) p.set('prompt', 'none'); else if (!S.user) p.set('prompt', 'select_account');
    if (S.user && S.user.mail) p.set('login_hint', S.user.mail);
    location.assign(`${authBase()}/authorize?${p}`);
  }

  async function handleRedirect() {
    const q = new URLSearchParams(location.search);
    if (!q.has('code') && !q.has('error')) return;
    const pk = store.get('pkce', null); store.set('pkce', null);
    history.replaceState(null, '', redirectUri());
    if (q.has('error')) {
      if (!(pk && pk.silent)) S.setupError = decodeURIComponent(q.get('error_description') || q.get('error')).split('\r')[0];
      S.needLogin = true; return;
    }
    if (!pk || pk.state !== q.get('state')) { S.setupError = 'La connexion a expiré, recommencez.'; return; }
    const body = new URLSearchParams({ client_id: S.clientId, grant_type: 'authorization_code', code: q.get('code'), redirect_uri: redirectUri(), code_verifier: pk.verifier, scope: SCOPE });
    const r = await fetch(`${authBase()}/token`, { method: 'POST', body });
    const j = await r.json();
    if (!r.ok) { S.setupError = (j.error_description || j.error || 'Connexion refusée').split('\r')[0]; return; }
    saveTokens(j); S.needLogin = false;
  }

  function saveTokens(j) {
    S.tokens = { at: j.access_token, exp: Date.now() + (j.expires_in - 120) * 1000, rt: j.refresh_token || (S.tokens && S.tokens.rt) };
    store.set('tokens', S.tokens);
  }

  class NeedLogin extends Error {}
  async function getToken() {
    if (S.tokens && S.tokens.at && Date.now() < S.tokens.exp) return S.tokens.at;
    if (S.tokens && S.tokens.rt) {
      try {
        const body = new URLSearchParams({ client_id: S.clientId, grant_type: 'refresh_token', refresh_token: S.tokens.rt, scope: SCOPE });
        const r = await fetch(`${authBase()}/token`, { method: 'POST', body });
        if (r.ok) { saveTokens(await r.json()); return S.tokens.at; }
        if (r.status >= 500) throw new Error('Microsoft indisponible');
      } catch (e) { if (!navigator.onLine || e instanceof TypeError) throw e; }
    }
    S.tokens = null; store.set('tokens', null);
    throw new NeedLogin();
  }

  /* Reconnexion : d'abord discrète (sans écran), sinon bouton « Se reconnecter ». */
  function onNeedLogin() {
    const last = store.get('silentTry', 0);
    if (navigator.onLine && Date.now() - last > 5 * 60 * 1000) { store.set('silentTry', Date.now()); startLogin(true); return; }
    S.needLogin = true; setSync('error', 'Reconnexion nécessaire'); render();
  }

  async function graph(path, opts = {}, retry = true) {
    const token = await getToken();
    const r = await fetch(path.startsWith('http') ? path : GRAPH + path, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + token } });
    if (r.status === 401 && retry) { S.tokens.exp = 0; return graph(path, opts, false); }
    return r;
  }
  async function graphJson(path) { const r = await graph(path); if (!r.ok) throw httpError(r); return r.json(); }
  function httpError(r) { const e = new Error('HTTP ' + r.status); e.status = r.status; return e; }

  /* ---------- Fichier OneDrive ---------- */
  async function downloadFile() {
    const item = await graphJson(`/me/drive/items/${S.file.id}`);
    let buf;
    const dl = item['@microsoft.graph.downloadUrl'];
    if (dl) { const r = await fetch(dl, { cache: 'no-store' }); if (r.ok) buf = await r.arrayBuffer(); }
    if (!buf) { const r = await graph(`/me/drive/items/${S.file.id}/content`); if (!r.ok) throw httpError(r); buf = await r.arrayBuffer(); }
    if (item.name !== S.file.name) { S.file.name = item.name; store.set('file', S.file); }
    return { buf, etag: item.eTag };
  }

  async function findFiles() {
    const seen = new Map(); const errors = []; let ok = 0;
    const add = (arr) => (arr || []).forEach((f) => { if (f.file && /\.xlsx$/i.test(f.name) && !seen.has(f.id)) seen.set(f.id, f); });
    const tryGet = async (path, pick) => {
      try { const j = await graphJson(path); ok++; add(pick(j)); }
      catch (e) { if (e instanceof NeedLogin) throw e; if (e.status !== 404) errors.push(e.status || e.message); else ok++; }
    };
    for (const n of [...new Set([CFG.defaultFileName, 'Habitudes.xlsx', 'Suivi-habitudes.xlsx'].filter(Boolean))]) await tryGet(`/me/drive/root:/${encodeURIComponent(n)}`, (f) => [f]);
    await tryGet(`/me/drive/root/children?$top=200`, (j) => j.value);
    for (const q of ['habitudes', 'suivi', 'xlsx']) await tryGet(`/me/drive/root/search(q='${q}')?$top=100`, (j) => j.value);
    if (!ok) { const e = new Error('graph'); e.userMsg = `OneDrive ne répond pas à l'application (erreur ${errors[0]}). Vérifiez que vous êtes connecté avec le bon compte Microsoft.`; throw e; }
    return [...seen.values()].sort((a, b) => (b.lastModifiedDateTime || '').localeCompare(a.lastModifiedDateTime || ''));
  }

  /* Explorateur de dossiers OneDrive (ne dépend pas de l'index de recherche) */
  async function browse(folder) {
    S.browse = { stack: S.browse ? S.browse.stack : [], items: null };
    if (folder === 'up') S.browse.stack.pop(); else if (folder) S.browse.stack.push(folder); else S.browse.stack = [];
    render();
    const cur = S.browse.stack[S.browse.stack.length - 1];
    try {
      const j = await graphJson(cur ? `/me/drive/items/${cur.id}/children?$top=200` : '/me/drive/root/children?$top=200');
      S.browse.items = (j.value || []).filter((x) => x.folder || (x.file && /\.xlsx$/i.test(x.name)))
        .sort((a, b) => (!!b.folder - !!a.folder) || a.name.localeCompare(b.name, 'fr'));
      S.setupError = '';
    } catch (e) {
      if (e instanceof NeedLogin) return onNeedLogin();
      S.browse.items = []; S.setupError = `Impossible d'ouvrir ce dossier OneDrive (erreur ${e.status || e.message}).`;
    }
    render();
  }

  async function load(silent) {
    if (!S.file) return;
    if (!navigator.onLine) { setSync('offline', 'Hors ligne'); render(); return; }
    if (!silent || !S.model) setSync('loading', 'Actualisation…');
    render();
    try {
      const { buf } = await downloadFile();
      const wb = await XL.openWorkbook(buf);
      const model = await XL.buildModel(wb);
      if (!model.trackers.length) throw Object.assign(new Error('format'), { userMsg: "Ce fichier ne contient pas d'onglet de suivi reconnu (mois en D2, année en D3)." });
      setModel(model);
      setSync('ok', 'À jour');
      if (Object.keys(S.pending).length) flush();
    } catch (e) { handleError(e); }
    render();
  }

  function setModel(model) {
    applyPendingTo(model);
    S.model = model; S.loadedAt = Date.now();
    store.set('model', model); store.set('loadedAt', S.loadedAt);
    if (!S.sheet || !model.trackers.some((t) => t.sheetName === S.sheet)) {
      const n = now(); const cur = model.trackers.find((t) => t.year === n.getFullYear() && t.monthIdx === n.getMonth());
      S.sheet = (cur || model.trackers[model.trackers.length - 1]).sheetName; store.set('sheet', S.sheet);
    }
  }

  function handleError(e) {
    if (e instanceof NeedLogin) return onNeedLogin();
    if (e.userMsg) { setSync('error', 'Erreur'); toast(e.userMsg); return; }
    if (e.status === 404) { setSync('error', 'Fichier introuvable'); toast("Le fichier n'est plus à cet emplacement sur OneDrive. Choisissez-le à nouveau dans les réglages."); return; }
    if (!navigator.onLine || e instanceof TypeError) { setSync('offline', 'Hors ligne'); return; }
    setSync('error', 'Erreur de synchro'); console.error(e);
  }

  /* ---------- Écritures : file d'attente, envoi groupé, sans écraser les modifs faites ailleurs ---------- */
  let flushTimer = null, flushing = false, again = false, retryTimer = null;
  function queueEdit(sheet, addr, val) {
    (S.pending[sheet] = S.pending[sheet] || {})[addr] = val;
    store.set('pending', S.pending);
    applyPendingTo(S.model); store.set('model', S.model);
    render();
    clearTimeout(flushTimer); flushTimer = setTimeout(flush, 1200);
  }

  async function flush() {
    if (flushing) { again = true; return; }
    if (!Object.keys(S.pending).length || !S.file) return;
    if (!navigator.onLine) { setSync('offline', 'Hors ligne'); render(); return; }
    flushing = true; clearTimeout(retryTimer);
    setSync('saving', 'Enregistrement…'); render();
    const snap = JSON.parse(JSON.stringify(S.pending));
    try {
      let done = false;
      for (let i = 0; i < 4 && !done; i++) {
        const { buf, etag } = await downloadFile();
        const out = await XL.applyEdits(buf, snap);
        const r = await graph(`/me/drive/items/${S.file.id}/content`, { method: 'PUT', headers: { 'Content-Type': XLSX_MIME, 'If-Match': etag }, body: out });
        if (r.status === 412 || r.status === 409) continue; // modifié entre-temps : on recommence sur la version la plus récente
        if (r.status === 423) { const e = new Error('locked'); e.locked = true; throw e; }
        if (!r.ok) throw httpError(r);
        for (const [sh, cells] of Object.entries(snap)) for (const [a, v] of Object.entries(cells)) {
          if (S.pending[sh] && S.pending[sh][a] === v) delete S.pending[sh][a];
          if (S.pending[sh] && !Object.keys(S.pending[sh]).length) delete S.pending[sh];
        }
        store.set('pending', S.pending);
        setModel(await XL.buildModel(await XL.openWorkbook(out)));
        done = true;
      }
      if (!done) throw new Error('conflict');
      setSync('ok', 'Enregistré');
    } catch (e) {
      if (e.locked) { setSync('saving', 'Nouvel essai…'); }
      else handleError(e);
      retryTimer = setTimeout(flush, 20000);
    }
    flushing = false; render();
    if (again) { again = false; flush(); }
  }

  /* Reporte les modifications en attente dans le modèle affiché */
  function applyPendingTo(model) {
    if (!model) return;
    for (const [sh, cells] of Object.entries(S.pending)) {
      const t = model.trackers.find((x) => x.sheetName === sh); if (!t) continue;
      for (const [addr, val] of Object.entries(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(addr); const col = m[1], row = +m[2], on = !!val;
        if (!t.cols) continue;
        if (col === t.cols.monthDone) { const it = t.monthly.find((x) => x.row === row); if (it) it.done = on; continue; }
        const h = t.habits.find((x) => x.row === row);
        if (h) { const d = XL.colToNum(col) - t.cols.day1 + 1; if (d >= 1 && d <= t.daysInMonth) h.checks[d - 1] = on; continue; }
        const w = t.weekly.find((x) => x.row === row); const wi = t.cols.week.indexOf(col);
        if (w && wi >= 0) w.weeks[wi] = on;
      }
    }
  }

  /* ---------- Calculs (identiques aux formules du fichier) ---------- */
  const T = () => S.model && S.model.trackers.find((t) => t.sheetName === S.sheet);
  const pct = (a, b) => (b ? Math.min(1, a / b) : 0);
  function stats(t) {
    const objective = (h) => t.weekdays.reduce((s, wd) => s + (h.planned[wd - 1] ? 1 : 0), 0);
    const habits = t.habits.map((h) => ({ ...h, objective: objective(h), done: h.checks.filter(Boolean).length }));
    const dailyDone = habits.reduce((s, h) => s + h.done, 0);
    const dailyObj = habits.reduce((s, h) => s + h.objective, 0);
    const weeklyDone = t.weekly.reduce((s, w) => s + w.weeks.filter(Boolean).length, 0);
    const weeklyObj = t.weekly.reduce((s, w) => s + w.objective, 0);
    const monthlyDone = t.monthly.filter((m) => m.done).length;
    const global = (dailyDone + weeklyDone + monthlyDone) / ((dailyObj + weeklyObj + t.monthly.length) || 1);
    const domains = ['Sport', 'Alimentation', 'Mental'].map((d) => {
      const hs = habits.filter((h) => h.domain === d);
      const obj = hs.reduce((s, h) => s + h.objective, 0);
      return { name: d, value: obj ? hs.reduce((s, h) => s + h.done, 0) / obj : 0 };
    });
    const dayScore = (day) => {
      const wd = t.weekdays[day - 1]; const planned = habits.filter((h) => h.planned[wd - 1]).length;
      return planned ? habits.filter((h) => h.checks[day - 1]).length / planned : 0;
    };
    return { habits, dailyDone, dailyObj, weeklyDone, weeklyObj, monthlyDone, global, domains, dayScore };
  }
  const globalMsg = (v) => v >= 0.8 ? 'Mois exceptionnel, bravo ! 🎉' : v >= 0.5 ? 'Tu es sur la bonne voie, continue 💪' : v > 0 ? 'Chaque case compte, avance pas à pas ✨' : 'Nouveau mois, nouveau départ 🌱';
  const habitMsg = (v) => v >= 0.9 ? "Bravo, c'est ancré ! 🎉" : v >= 0.6 ? 'Belle régularité, continue 💪' : v >= 0.3 ? 'Ça progresse, tiens bon ✨' : 'Une case à la fois, tu es lancé·e 🌱';
  const dayMsg = (v, n) => n === 0 ? "Rien de prévu aujourd'hui, profite 🌿" : v >= 1 ? 'Journée complète, bravo ! 🎉' : v >= 0.6 ? 'Belle journée, continue 💪' : v > 0 ? 'Une case à la fois ✨' : 'On commence quand tu veux 🌱';
  const isCurrentMonth = (t) => { const n = now(); return t.year === n.getFullYear() && t.monthIdx === n.getMonth(); };
  const todayInMonth = (t) => (isCurrentMonth(t) ? now().getDate() : null);
  const isFuture = (t, d) => new Date(t.year, t.monthIdx, d) > now();
  const prettySection = (s) => (s ? s.charAt(0) + s.slice(1).toLowerCase() : 'Autre');
  const color = (d) => DOMAIN_COLOR[XL.norm(d)] || 'var(--other)';
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const P = (v) => Math.round(v * 100) + ' %';

  /* ---------- Icônes ---------- */
  const ICON = {
    check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
    gear: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
    back: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 5l-7 7 7 7"/></svg>',
    today: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="4.2"/><path d="M12 2.5v2.2M12 19.3v2.2M4.6 4.6l1.6 1.6M17.8 17.8l1.6 1.6M2.5 12h2.2M19.3 12h2.2M4.6 19.4l1.6-1.6M17.8 6.2l1.6-1.6"/></svg>',
    month: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M4 20V11M10 20V5M16 20v-6M22 20H2"/></svg>',
    goals: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.4" fill="currentColor"/></svg>',
    plan: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
  };
  const ring = (v, size, stroke, inner) => {
    const r = (size - stroke) / 2, c = 2 * Math.PI * r;
    return `<div class="ring" style="width:${size}px;height:${size}px"><svg width="${size}" height="${size}"><circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="var(--faint)" stroke-width="${stroke}"/><circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="var(--green)" stroke-width="${stroke}" stroke-linecap="round" stroke-dasharray="${c}" stroke-dashoffset="${c * (1 - Math.min(1, v))}" style="transition:stroke-dashoffset .5s"/></svg><div class="lbl">${inner}</div></div>`;
  };

  /* ---------- Vues ---------- */
  function vToday(t, st) {
    if (S.day == null || S.day > t.daysInMonth) S.day = todayInMonth(t) || 1;
    const d = S.day, wd = t.weekdays[d - 1], tday = todayInMonth(t);
    const planned = st.habits.filter((h) => h.planned[wd - 1]);
    const unplanned = st.habits.filter((h) => !h.planned[wd - 1]);
    const doneN = planned.filter((h) => h.checks[d - 1]).length;
    const v = planned.length ? doneN / planned.length : 0;
    const strip = t.weekdays.map((w, i) => {
      const n = i + 1, sc = st.dayScore(n);
      return `<button class="day ${n === d ? 'sel' : ''} ${n === tday ? 'today' : ''}" data-act="pickday" data-day="${n}"><small>${DAYS_SHORT[w - 1]}</small><b>${n}</b><div class="bar"><span style="width:${isFuture(t, n) ? 0 : Math.round(sc * 100)}%"></span></div></button>`;
    }).join('');
    const sections = [];
    planned.forEach((h) => { let s = sections.find((x) => x.name === h.section); if (!s) sections.push((s = { name: h.section, items: [] })); s.items.push(h); });
    sections.forEach((s) => s.items.sort((a, b) => XL.timeMinutes(a.times[wd - 1]) - XL.timeMinutes(b.times[wd - 1])));
    const row = (h, extra = '') => `<button class="habit ${h.checks[d - 1] ? 'done' : ''} ${extra}" data-act="toggle" data-row="${h.row}" data-day="${d}">
        <span class="dot" style="background:${color(h.domain)}"></span>
        <span class="nm"><b>${esc(h.name)}</b><small>${esc([h.times[wd - 1] || h.time, h.domain].filter(Boolean).join(' · '))}</small></span>
        <span class="check">${ICON.check}</span></button>`;
    return `
      <div class="days" id="days">${strip}</div>
      <div class="card hero">${ring(v, 92, 9, `<span>${doneN}<small>sur ${planned.length}</small></span>`)}
        <div><div class="msg">${dayMsg(v, planned.length)}</div><div class="sub">${d === tday ? "Aujourd'hui" : DAYS_LONG[wd - 1] + ' ' + d + ' ' + MONTHS_LOWER[t.monthIdx]}${isFuture(t, d) ? ' · à venir' : ''}</div></div></div>
      ${sections.map((s) => `<h2>${esc(prettySection(s.name))}</h2><div class="list">${s.items.map((h) => row(h)).join('')}</div>`).join('')}
      ${!planned.length ? '<div class="empty">Aucune habitude prévue ce jour-là.</div>' : ''}
      ${unplanned.length ? `<button class="disclose" data-act="unplanned">${S.showUnplanned ? 'Masquer' : 'Afficher'} les habitudes non prévues ce jour (${unplanned.length})</button>
        ${S.showUnplanned ? `<div class="list" style="margin-top:8px">${unplanned.map((h) => row(h, 'unplanned')).join('')}</div>` : ''}` : ''}`;
  }

  function vMonth(t, st) {
    const tday = todayInMonth(t);
    const bars = t.weekdays.map((_, i) => {
      const n = i + 1, fut = isFuture(t, n), sc = st.dayScore(n);
      return `<div class="${fut ? 'future' : ''} ${n === tday ? 'today' : ''}" style="height:${fut ? 4 : Math.max(3, sc * 100)}%" title="${n} : ${P(sc)}"></div>`;
    }).join('');
    return `
      <div class="card hero" style="margin-top:6px">${ring(st.global, 104, 10, `<span style="font-size:24px">${Math.round(st.global * 100)}<span style="font-size:14px;color:var(--muted)"> %</span></span>`)}
        <div><div class="sub" style="margin:0 0 4px">Progression globale</div><div class="msg">${globalMsg(st.global)}</div></div></div>
      <div class="tiles">
        <div class="tile"><small>Quotidiennes</small><b>${st.dailyDone}<span style="color:var(--muted);font-size:14px"> / ${st.dailyObj}</span></b></div>
        <div class="tile"><small>Hebdo</small><b>${st.weeklyDone}<span style="color:var(--muted);font-size:14px"> / ${st.weeklyObj}</span></b></div>
        <div class="tile"><small>Du mois</small><b>${st.monthlyDone}<span style="color:var(--muted);font-size:14px"> / ${t.monthly.length}</span></b></div>
      </div>
      <h2>Par domaine</h2>
      <div class="card">${st.domains.map((d) => `<div class="dom"><span>${d.name}</span><div class="meter"><span style="width:${Math.round(d.value * 100)}%;background:${color(d.name)}"></span></div><em>${P(d.value)}</em></div>`).join('')}</div>
      <h2>Progression quotidienne</h2>
      <div class="card"><div class="chart">${bars}</div><div class="axis"><span>1</span><span>${Math.ceil(t.daysInMonth / 2)}</span><span>${t.daysInMonth}</span></div></div>
      <h2>Chaque habitude</h2>
      <div class="card" style="padding:4px 14px">${st.habits.map((h) => {
        const v = pct(h.done, h.objective);
        return `<button class="hrow" data-act="habit" data-row="${h.row}"><div class="l1">${esc(h.name)}<span>${h.done} / ${h.objective}</span></div>
          <div class="meter"><span style="width:${Math.round(v * 100)}%;background:${color(h.domain)}"></span></div><small>${habitMsg(v)}</small></button>`;
      }).join('')}</div>`;
  }

  function vGoals(t) {
    const tday = todayInMonth(t); const curWeek = tday ? Math.min(5, Math.ceil(tday / 7)) : 0;
    return `
      <h2 style="margin-top:12px">Habitudes hebdomadaires</h2>
      ${t.weekly.length ? t.weekly.map((w) => `<div class="card goal"><div class="l1"><b>${esc(w.name)}</b><span>${w.weeks.filter(Boolean).length} / ${w.objective}</span></div>
        <div style="font-size:12.5px;color:var(--muted);margin-top:2px"><span class="dot" style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${color(w.domain)};margin-right:6px"></span>${esc([w.domain, w.weekday ? DAYS_LONG[w.weekday - 1] + (w.time ? ' ' + w.time : '') : w.time].filter(Boolean).join(' · '))}</div>
        <div class="weeks">${w.weeks.map((on, i) => `<button class="wk ${on ? 'on' : ''} ${i + 1 === curWeek ? 'cur' : ''}" data-act="week" data-row="${w.row}" data-w="${i}">S${i + 1}</button>`).join('')}</div></div>`).join('')
        : '<div class="empty card">Aucune habitude hebdomadaire dans le fichier.</div>'}
      <h2>À faire ce mois-ci</h2>
      <div class="card" style="padding:2px 0">${t.monthly.length ? t.monthly.map((m) => `<button class="todo ${m.done ? 'on' : ''}" data-act="monthly" data-row="${m.row}"><span class="box">${ICON.check}</span><span>${esc(m.name)}</span></button>`).join('') : '<div class="empty">Rien de prévu ce mois-ci.</div>'}</div>`;
  }

  function vPlan(t) {
    const todayWd = XL.excelWeekday(now());
    if (!S.planDay) S.planDay = todayWd;
    const seg = `<div class="seg" style="margin-top:8px">${DAYS_SHORT.map((l, i) => `<button class="${S.planDay === i + 1 ? 'on' : ''} ${todayWd === i + 1 ? 'today' : ''}" data-act="planday" data-d="${i + 1}">${l}</button>`).join('')}</div>`;
    const nowM = now().getHours() * 60 + now().getMinutes();
    if (!t.hasTimes) return seg + vPlanSheet(todayWd, nowM);
    // Planning calculé depuis le Tracker (heure habituelle ou heure du jour), comme dans le fichier
    const wd = S.planDay, n = now();
    const date = new Date(n.getFullYear(), n.getMonth(), n.getDate() + (wd - todayWd));
    const d = date.getFullYear() === t.year && date.getMonth() === t.monthIdx ? date.getDate() : null;
    const items = [];
    t.habits.forEach((h) => { if (h.planned[wd - 1]) items.push({ time: h.times[wd - 1], label: h.name, domain: h.domain, row: h.row, done: d ? h.checks[d - 1] : null }); });
    t.weekly.forEach((w) => { if (w.weekday === wd) items.push({ time: w.time, label: w.name, domain: w.domain, weekly: true }); });
    items.sort((a, b) => XL.timeMinutes(a.time) - XL.timeMinutes(b.time));
    let nowIdx = -1;
    if (wd === todayWd) items.forEach((it, i) => { if (XL.timeMinutes(it.time) <= nowM) nowIdx = i; });
    return `${seg}
      <h2>${DAYS_LONG[wd - 1]} ${date.getDate()} ${MONTHS_LOWER[date.getMonth()]}</h2>
      <div class="card"><div class="timeline">${items.length ? items.map((it, i) => `<div class="ti ${i === nowIdx ? 'now' : ''} ${nowIdx >= 0 && i < nowIdx ? 'past' : ''}">
        <time>${esc(it.time || '—')}</time><span class="pt" style="border-color:${color(it.domain)}"></span>
        <span>${i === nowIdx ? '<b>En ce moment</b>' : ''}${esc(it.label)}${it.weekly ? ' <small style="color:var(--muted)">· hebdo</small>' : ''}</span>
        ${it.done != null ? `<button class="tick ${it.done ? 'on' : ''}" data-act="toggle" data-row="${it.row}" data-day="${d}" aria-label="Cocher">${ICON.check}</button>` : '<span></span>'}</div>`).join('')
        : '<div class="empty">Rien de planifié ce jour-là.</div>'}</div></div>
      <p style="color:var(--muted);font-size:12.5px;text-align:center;margin-top:12px">Les heures viennent du Tracker : heure habituelle, ou heure du jour si vous en avez mis une.</p>`;
  }

  function vPlanSheet(todayWd, nowM) {
    const p = S.model.planning;
    if (!p || !p.blocks.length) return '<div class="empty">Aucun onglet « Planning » trouvé dans le fichier.</div>';
    const block = p.blocks.find((b) => b.weekday === XL.norm(DAYS_LONG[S.planDay - 1]));
    let nowIdx = -1;
    if (block && S.planDay === todayWd) block.items.forEach((it, i) => { if (XL.timeMinutes(it.time) <= nowM) nowIdx = i; });
    return `<h2>${DAYS_LONG[S.planDay - 1]}</h2>
      <div class="card"><div class="timeline">${block && block.items.length ? block.items.map((it, i) => `<div class="ti ${i === nowIdx ? 'now' : ''} ${nowIdx >= 0 && i < nowIdx ? 'past' : ''}"><time>${esc(it.time)}</time><span class="pt"></span><span>${i === nowIdx ? '<b>En ce moment</b>' : ''}${esc(it.label)}</span><span></span></div>`).join('') : '<div class="empty">Rien de planifié ce jour-là.</div>'}</div></div>`;
  }

  function vSettings() {
    const ago = S.loadedAt ? new Date(S.loadedAt).toLocaleString('fr-FR', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }) : '—';
    const trackers = (S.model && S.model.trackers) || [];
    return `
      <h2 style="margin-top:12px">Compte Microsoft</h2>
      <div class="card"><div class="row"><span>Connecté</span><span>${esc((S.user && (S.user.mail || S.user.name)) || '—')}</span></div>
      <div class="row"><span>Dernière synchro</span><span>${ago}</span></div>
      <div class="row"><span>Modifs en attente</span><span>${Object.values(S.pending).reduce((s, c) => s + Object.keys(c).length, 0)}</span></div></div>
      <button class="btn" data-act="refresh">Actualiser maintenant</button>
      <h2>Fichier OneDrive</h2>
      <div class="card"><div class="row"><span>Fichier</span><span>${esc((S.file && S.file.name) || '—')}</span></div></div>
      <button class="btn ghost" data-act="choosefile">Choisir un autre fichier</button>
      ${trackers.length > 1 ? `<h2>Mois affiché</h2><div class="card" style="padding:2px 0">${trackers.map((t) => `<button class="pick ${t.sheetName === S.sheet ? 'on' : ''}" data-act="sheet" data-s="${esc(t.sheetName)}"><span>${t.monthName} ${t.year}<small>Onglet « ${esc(t.sheetName)} »</small></span>${t.sheetName === S.sheet ? ICON.check.replace('<svg', '<svg width="18" height="18"') : ''}</button>`).join('')}</div>` : ''}
      <h2>Avancé</h2>
      <div class="card form" style="padding-top:2px"><label>ID d'application (client)</label><input id="cid" value="${esc(S.clientId)}" autocapitalize="off" autocorrect="off" spellcheck="false">
      <label>Adresse de redirection à déclarer chez Microsoft</label><div class="code">${esc(redirectUri())}</div>
      <button class="btn ghost" data-act="savecid">Enregistrer l'ID</button></div>
      <button class="btn danger" data-act="logout">Se déconnecter</button>
      <p style="text-align:center;color:var(--muted);font-size:12px;margin-top:20px">Mes habitudes · v${VERSION} · fonctionne seule, entre votre iPhone et OneDrive</p>`;
  }

  function vSetup() {
    const step = !S.clientId ? 1 : !S.tokens ? 2 : 3;
    let body = '';
    if (step === 1) body = `
      <div class="card"><ol class="steps"><li>Inscrivez l'application chez Microsoft (une seule fois, voir le guide).</li><li>Déclarez cette adresse de redirection :<div class="code">${esc(redirectUri())}</div></li><li>Collez ici l'« ID d'application (client) ».</li></ol>
      <div class="form"><input id="cid" placeholder="ex. 1a2b3c4d-…" autocapitalize="off" autocorrect="off" spellcheck="false"></div>
      <button class="btn" data-act="savecid">Continuer</button></div>`;
    else if (step === 2) body = `
      <div class="card"><p style="margin:0 0 4px;line-height:1.45">Connectez-vous avec le compte Microsoft qui contient votre fichier sur OneDrive.</p>
      <button class="btn" data-act="login">Se connecter avec Microsoft</button>
      <button class="btn ghost" data-act="resetcid">Modifier l'ID d'application</button></div>`;
    else if (S.browse) {
      const st = S.browse.stack; const here = st.length ? st[st.length - 1].name : 'Mon OneDrive';
      body = `
      <div class="card" style="padding:2px 0">
        <div class="pick" style="font-weight:700">${st.length ? '<button data-act="browseup" style="color:var(--green);font-weight:700">‹ Retour</button>' : ''}<span style="flex:1;text-align:${st.length ? 'right' : 'left'}">📂 ${esc(here)}</span></div>
        ${S.browse.items == null ? '<div class="spinner" style="margin:24px auto"></div>'
          : S.browse.items.length ? S.browse.items.map((f) => f.folder
            ? `<button class="pick" data-act="browsein" data-id="${esc(f.id)}" data-name="${esc(f.name)}"><span>📁 ${esc(f.name)}<small>${f.folder.childCount != null ? f.folder.childCount + ' élément(s)' : 'Dossier'}</small></span>›</button>`
            : `<button class="pick" data-act="pickfile" data-id="${esc(f.id)}" data-name="${esc(f.name)}"><span>📗 ${esc(f.name)}<small>Fichier Excel</small></span><b style="color:var(--green)">Choisir</b></button>`).join('')
          : '<div class="empty">Aucun dossier ni fichier Excel ici.</div>'}</div>
      <button class="btn ghost" data-act="findfiles">Revenir à la recherche</button>`;
    } else body = `
      <div class="card" style="padding:2px 0">${S.files == null ? '<div class="spinner"></div><div class="empty" style="padding-top:0">Recherche de vos fichiers Excel…</div>'
        : S.files.length ? S.files.map((f) => `<button class="pick" data-act="pickfile" data-id="${esc(f.id)}" data-name="${esc(f.name)}"><span>${esc(f.name)}<small>${esc(decodeURIComponent((f.parentReference && f.parentReference.path || '').replace(/^\/drive\/root:?/, '')) || 'Mon OneDrive')}</small></span>›</button>`).join('')
        : `<div class="empty">La recherche n'a trouvé aucun fichier Excel. Parcourez vos dossiers ci-dessous.</div>`}</div>
      <button class="btn" data-act="browse">Parcourir mon OneDrive</button>
      <button class="btn ghost" data-act="findfiles">Rechercher à nouveau</button>`;
    if (step === 3) body += `<p style="text-align:center;color:var(--muted);font-size:13px;margin-top:18px">Connecté : ${esc((S.user && (S.user.mail || S.user.name)) || 'compte Microsoft')}<br><button data-act="logout" style="color:var(--today);font-weight:600;margin-top:6px">Changer de compte</button></p>`;
    const titles = ['', 'Bienvenue', 'Connexion', 'Votre fichier'];
    return `<div class="welcome"><img src="icons/icon-192.png" alt=""><h1>${titles[step]}</h1><p>${step === 3 ? 'Choisissez votre fichier de suivi d’habitudes.' : 'Votre suivi d’habitudes, relié à votre fichier Excel sur OneDrive.'}</p></div>
      ${S.setupError ? `<div class="banner warn">${esc(S.setupError)}</div>` : ''}${body}`;
  }

  function vModal(t, st) {
    const h = st.habits.find((x) => x.row === S.modal); if (!h) return '';
    const first = t.weekdays[0], tday = todayInMonth(t);
    const cells = [];
    for (let i = 1; i < first; i++) cells.push('<span></span>');
    t.weekdays.forEach((wd, i) => {
      const n = i + 1, on = h.checks[i], planned = h.planned[wd - 1];
      cells.push(`<button class="${on ? 'done' : ''} ${!planned && !on ? 'off' : ''} ${n === tday ? 'today' : ''}" data-act="toggle" data-row="${h.row}" data-day="${n}">${n}</button>`);
    });
    return `<div class="sheet-bg" data-act="close"><div class="sheet" data-stop="1"><div class="grab"></div>
      <h3>${esc(h.name)}</h3><p class="sub">${esc(prettySection(h.section))} · ${esc(h.domain)} · ${h.done} / ${h.objective} · ${P(pct(h.done, h.objective))}</p>
      <div class="cal">${DAYS_SHORT.map((l) => `<span class="hd">${l}</span>`).join('')}${cells.join('')}</div>
      <div class="legend"><span><i style="background:var(--green)"></i>Fait</span><span><i style="border:1.5px solid var(--line);background:var(--card)"></i>Prévu</span><span><i style="background:var(--faint)"></i>Pas prévu</span></div>
      <button class="btn ghost" data-act="close">Fermer</button></div></div>`;
  }

  /* ---------- Rendu ---------- */
  const app = document.getElementById('app');
  function setSync(s, msg) { S.sync = s; S.syncMsg = msg || ''; }
  function toast(msg) { S.toast = msg; render(); clearTimeout(toast.t); toast.t = setTimeout(() => { S.toast = ''; render(); }, 4200); }

  function render() {
    const setup = !S.clientId || (!S.tokens && !S.model) || !S.file;
    const t = !setup && T();
    if (setup) {
      app.innerHTML = `<main>${vSetup()}</main>${S.toast ? `<div class="toast">${esc(S.toast)}</div>` : ''}`;
      return;
    }
    if (!t) { app.innerHTML = `<main><div class="spinner"></div><div class="empty" style="padding-top:0">Chargement de votre fichier…</div></main>`; return; }
    const st = stats(t);
    const dd = S.day || todayInMonth(t) || 1;
    const head = {
      today: [`${t.monthName} ${t.year}`, `${DAYS_LONG[t.weekdays[Math.min(dd, t.daysInMonth) - 1] - 1]} ${Math.min(dd, t.daysInMonth)}`],
      month: [`${t.monthName} ${t.year}`, 'Mon mois'],
      goals: [`${t.monthName} ${t.year}`, 'Objectifs'],
      plan: ['Planning', 'Ma semaine'],
      settings: ['Mes habitudes', 'Réglages'],
    }[S.view];
    const pendingN = Object.values(S.pending).reduce((s, c) => s + Object.keys(c).length, 0);
    const banner = S.needLogin ? `<div class="banner warn">Votre session Microsoft a expiré.<button data-act="login">Se reconnecter</button></div>`
      : S.sync === 'offline' && pendingN ? `<div class="banner">Hors ligne : ${pendingN} coche${pendingN > 1 ? 's' : ''} ser${pendingN > 1 ? 'ont' : 'a'} envoyée${pendingN > 1 ? 's' : ''} dès le retour du réseau.</div>` : '';
    const body = { today: vToday, month: vMonth, goals: vGoals, plan: vPlan, settings: vSettings }[S.view](t, st);
    const scrollDays = document.getElementById('days'); const keepScroll = scrollDays ? scrollDays.scrollLeft : null;
    app.innerHTML = `
      <header class="top">${S.view === 'settings' ? `<button class="iconbtn" data-act="tab" data-v="today" aria-label="Retour">${ICON.back}</button>` : ''}
        <div class="t"><div class="eyebrow">${esc(head[0])}</div><h1>${esc(head[1])}</h1></div>
        <button class="sync ${S.sync}" data-act="refresh"><i></i>${esc(S.syncMsg || 'À jour')}</button>
        ${S.view !== 'settings' ? `<button class="iconbtn" data-act="tab" data-v="settings" aria-label="Réglages">${ICON.gear}</button>` : ''}</header>
      <main>${banner}${body}</main>
      <nav class="tabs">${[['today', "Aujourd'hui"], ['month', 'Mois'], ['goals', 'Objectifs'], ['plan', 'Planning']].map(([v, l]) => `<button class="${S.view === v ? 'on' : ''}" data-act="tab" data-v="${v}">${ICON[v]}<span>${l}</span></button>`).join('')}</nav>
      ${S.modal ? vModal(t, st) : ''}${S.toast ? `<div class="toast">${esc(S.toast)}</div>` : ''}`;
    const days = document.getElementById('days');
    if (days) {
      if (keepScroll != null && !render.center) days.scrollLeft = keepScroll;
      else { const sel = days.querySelector('.sel'); if (sel) days.scrollLeft = sel.offsetLeft - days.clientWidth / 2 + sel.clientWidth / 2; render.center = false; }
    }
  }

  /* ---------- Actions ---------- */
  app.addEventListener('click', async (ev) => {
    const el = ev.target.closest('[data-act]'); if (!el) return;
    if (el.dataset.act === 'close' && ev.target.closest('[data-stop]') && !ev.target.closest('button[data-act="close"]')) return;
    const a = el.dataset.act, t = T();
    switch (a) {
      case 'tab': S.view = el.dataset.v; S.modal = null; if (S.view === 'today') render.center = true; window.scrollTo(0, 0); render(); break;
      case 'pickday': S.day = +el.dataset.day; render(); break;
      case 'unplanned': S.showUnplanned = !S.showUnplanned; render(); break;
      case 'toggle': {
        const h = t.habits.find((x) => x.row === +el.dataset.row); const d = +el.dataset.day;
        queueEdit(t.sheetName, XL.numToCol(t.cols.day1 + d - 1) + h.row, h.checks[d - 1] ? '' : 'x'); break;
      }
      case 'week': { const w = t.weekly.find((x) => x.row === +el.dataset.row); const i = +el.dataset.w; queueEdit(t.sheetName, t.cols.week[i] + w.row, w.weeks[i] ? '' : 'x'); break; }
      case 'monthly': { const m = t.monthly.find((x) => x.row === +el.dataset.row); queueEdit(t.sheetName, t.cols.monthDone + m.row, m.done ? '' : 'x'); break; }
      case 'habit': S.modal = +el.dataset.row; render(); break;
      case 'close': S.modal = null; render(); break;
      case 'planday': S.planDay = +el.dataset.d; render(); break;
      case 'sheet': S.sheet = el.dataset.s; store.set('sheet', S.sheet); S.day = null; render(); break;
      case 'refresh': if (Object.keys(S.pending).length) flush(); else load(); break;
      case 'savecid': {
        const v = (document.getElementById('cid').value || '').trim();
        if (!/^[0-9a-f-]{36}$/i.test(v)) { toast("L'ID doit ressembler à 1a2b3c4d-1234-…(36 caractères)."); break; }
        if (v !== S.clientId) { S.clientId = v; store.set('clientId', v); S.tokens = null; store.set('tokens', null); }
        S.setupError = ''; render(); break;
      }
      case 'resetcid': S.clientId = ''; store.set('clientId', null); render(); break;
      case 'login': store.set('silentTry', Date.now()); startLogin(false); break;
      case 'logout':
        ['tokens', 'user', 'file', 'model', 'pending', 'sheet', 'loadedAt'].forEach((k) => store.set(k, null));
        Object.assign(S, { tokens: null, user: null, file: null, model: null, pending: {}, sheet: null, view: 'today', files: null });
        location.assign(`${authBase()}/logout?post_logout_redirect_uri=${encodeURIComponent(redirectUri())}`); break;
      case 'choosefile': S.file = null; S.files = null; S.browse = null; store.set('file', null); render(); runFind(); break;
      case 'findfiles': S.files = null; S.browse = null; render(); runFind(); break;
      case 'browse': browse(null); break;
      case 'browsein': browse({ id: el.dataset.id, name: el.dataset.name }); break;
      case 'browseup': browse('up'); break;
      case 'pickfile':
        S.file = { id: el.dataset.id, name: el.dataset.name }; store.set('file', S.file); S.browse = null;
        S.model = null; S.sheet = null; S.pending = {}; store.set('pending', {}); render(); load(); break;
    }
  });

  async function runFind() {
    try { S.files = await findFiles(); S.setupError = ''; }
    catch (e) { if (e instanceof NeedLogin) return onNeedLogin(); S.files = []; S.setupError = e.userMsg || 'Impossible de lister vos fichiers OneDrive.'; }
    const hab = S.files.filter((f) => /habitudes/i.test(f.name));
    if (hab.length === 1 && !S.file) { S.file = { id: hab[0].id, name: hab[0].name }; store.set('file', S.file); load(); }
    else if (!S.files.length && !S.setupError) { browse(null); return; }
    render();
  }

  async function fetchUser() {
    try { const me = await graphJson('/me'); S.user = { name: me.displayName, mail: me.mail || me.userPrincipalName }; store.set('user', S.user); } catch (e) { if (e instanceof NeedLogin) throw e; }
  }

  /* ---------- Démarrage ---------- */
  window.addEventListener('online', () => { if (Object.keys(S.pending).length) flush(); else load(true); });
  window.addEventListener('offline', () => { setSync('offline', 'Hors ligne'); render(); });
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && S.file && S.tokens) {
      if (Object.keys(S.pending).length) flush(); else if (Date.now() - S.loadedAt > 30000) load(true);
      render.center = true; render();
    }
  });
  setInterval(() => { if (S.view === 'plan') render(); }, 60000);
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});

  (async function init() {
    if (S.model) { setSync(navigator.onLine ? 'loading' : 'offline', navigator.onLine ? 'Actualisation…' : 'Hors ligne'); render.center = true; }
    render();
    try { await handleRedirect(); } catch (e) { S.setupError = 'Connexion impossible : ' + e.message; }
    if (!S.clientId || !S.tokens) { render(); return; }
    try {
      if (!S.user) await fetchUser();
      if (!S.file) { render(); await runFind(); return; }
      render.center = true; await load(!!S.model);
    } catch (e) { handleError(e); render(); }
  })();
})();

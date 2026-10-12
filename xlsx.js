/* Lecture et modification minimale d'un fichier .xlsx (sans toucher à la mise en forme). */
(function (global) {
  'use strict';

  const MONTHS = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'];
  const MONTHS_FR = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
  const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();

  function colToNum(col) { let n = 0; for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function numToCol(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function splitAddr(a) { const m = /^([A-Z]+)(\d+)$/.exec(a); return { col: colToNum(m[1]), row: +m[2] }; }

  const parseXml = (s) => new DOMParser().parseFromString(s, 'application/xml');
  const byTag = (el, tag) => Array.from(el.getElementsByTagNameNS('*', tag));
  const textOf = (el) => byTag(el, 't').filter((t) => !t.closest || !hasAncestor(t, 'rPh')).map((t) => t.textContent).join('');
  function hasAncestor(el, name) { let p = el.parentNode; while (p && p.nodeType === 1) { if (p.localName === name) return true; p = p.parentNode; } return false; }

  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
    return parts.join('/');
  }

  async function openWorkbook(buf) {
    const zip = await JSZip.loadAsync(buf);
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const rels = {};
    byTag(parseXml(relsXml), 'Relationship').forEach((r) => { rels[r.getAttribute('Id')] = resolvePath('xl/workbook.xml', r.getAttribute('Target')); });
    const sheets = byTag(parseXml(wbXml), 'sheet').map((s) => ({
      name: s.getAttribute('name'),
      path: rels[s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id')],
    }));
    let sst = [];
    const sstFile = zip.file('xl/sharedStrings.xml');
    if (sstFile) sst = byTag(parseXml(await sstFile.async('string')), 'si').map(textOf);
    return { zip, sheets, sst, cache: {} };
  }

  async function readSheet(wb, name) {
    if (wb.cache[name]) return wb.cache[name];
    const sh = wb.sheets.find((s) => s.name === name);
    if (!sh) return null;
    const xml = await wb.zip.file(sh.path).async('string');
    const cells = {};
    byTag(parseXml(xml), 'c').forEach((c) => {
      const r = c.getAttribute('r'); if (!r) return;
      const t = c.getAttribute('t');
      const vEl = byTag(c, 'v')[0]; const fEl = byTag(c, 'f')[0];
      let v = null;
      if (t === 's') v = vEl ? wb.sst[+vEl.textContent] ?? '' : '';
      else if (t === 'inlineStr') v = textOf(c);
      else if (t === 'str' || t === 'e') v = vEl ? vEl.textContent : '';
      else if (t === 'b') v = vEl ? vEl.textContent === '1' : null;
      else if (vEl) v = Number(vEl.textContent);
      const cell = { v };
      if (fEl) cell.f = fEl.textContent;
      if (v !== null || fEl) cells[r] = cell;
    });
    wb.cache[name] = cells;
    return cells;
  }

  /* ---------- Modèle métier ---------- */
  const isX = (cell) => !!cell && typeof cell.v === 'string' && cell.v.trim().toLowerCase() === 'x';
  const str = (cell) => (cell && cell.v != null ? String(cell.v).trim() : '');
  const excelWeekday = (date) => ((date.getDay() + 6) % 7) + 1; // lundi = 1

  /* Heures : « 7h00 », « 7:00 » ou une heure Excel (fraction de jour) → « 7h00 » */
  function fmtTime(cell) {
    if (!cell || cell.v == null || cell.v === '') return '';
    if (typeof cell.v === 'number') {
      if (cell.v > 0 && cell.v < 1) { const m = Math.round(cell.v * 1440); return Math.floor(m / 60) + 'h' + String(m % 60).padStart(2, '0'); }
      return '';
    }
    const m = /^(\d{1,2})\s*[h:]\s*(\d{0,2})/i.exec(String(cell.v).trim());
    return m ? +m[1] + 'h' + (m[2] || '00').padStart(2, '0') : '';
  }
  const timeMinutes = (t) => { const m = /^(\d{1,2})h(\d{2})$/.exec(t || ''); return m ? +m[1] * 60 + +m[2] : 9999; };

  /* Repère les colonnes d'après leurs titres : l'app suit le fichier même si on déplace des colonnes */
  function layout(cells) {
    const where = {}; // texte normalisé → [{row, col}]
    for (const [a, c] of Object.entries(cells)) {
      if (typeof c.v !== 'string' || !c.v.trim()) continue;
      const k = norm(c.v); const { col, row } = splitAddr(a);
      (where[k] = where[k] || []).push({ row, col });
    }
    const first = (k, filter = () => true) => (where[norm(k)] || []).filter(filter).sort((x, y) => x.row - y.row || x.col - y.col)[0];
    const at = (row, col) => cells[numToCol(col) + row];

    // Mois et année : la première valeur à droite des libellés « Mois » et « Année »
    const right = (lbl, test) => { const p = first(lbl, (x) => x.row <= 8); if (!p) return null; for (let c = p.col + 1; c <= p.col + 6; c++) { const v = at(p.row, c); if (v && test(v)) return v; } return null; };
    const mCell = right('Mois', (v) => MONTHS.includes(norm(v.v)));
    const yCell = right('Année', (v) => typeof v.v === 'number');
    if (!mCell || !yCell) return null;

    const nameH = first('Suivi des habitudes'); if (!nameH) return null;
    const R = nameH.row; const inRow = (k) => first(k, (x) => x.row === R);
    const timeH = inRow('Heure'), domH = inRow('Domaine'), objH = inRow('Objectif'), planH = inRow('Jours prévus');
    if (!domH || !objH || !planH) return null;
    const scoreH = first('Score du jour', (x) => x.row > R);
    const weekH = first('Habitudes hebdomadaires', (x) => x.row > R);
    const L = {
      monthIdx: MONTHS.indexOf(norm(mCell.v)), year: yCell.v,
      sectionCol: Math.min(timeH ? timeH.col : nameH.col, nameH.col) - 1,
      timeCol: timeH ? timeH.col : null, nameCol: nameH.col, domainCol: domH.col, objCol: objH.col,
      day1: objH.col + 1, plan1: planH.col,
      firstRow: R + 2, lastRow: scoreH ? scoreH.row - 1 : R + 28,
    };
    if (weekH) {
      const W = weekH.row; const wRow = (k) => first(k, (x) => x.row === W);
      const sem = [1, 2, 3, 4, 5].map((i) => wRow('Sem. ' + i)).filter(Boolean).map((x) => x.col);
      const mo = wRow('Habitudes à faire sur le mois');
      const fait = first('Fait ?', (x) => x.row >= W - 2 && x.row <= W + 1);
      const wTime = wRow('Heure'), wDom = wRow('Domaine'), wObj = wRow('Objectif');
      L.weekly = { firstRow: W + 2, lastRow: W + 6, nameCol: weekH.col, timeCol: wTime ? wTime.col : null, domainCol: wDom ? wDom.col : weekH.col + 2, objCol: wObj ? wObj.col : null, sem };
      if (mo && fait) L.monthly = { labelCol: mo.col, doneCol: fait.col };
    }
    return L;
  }

  function trackerModel(cells, sheetName, L) {
    const { monthIdx, year } = L;
    const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();
    const weekdays = []; // jour de semaine (1..7) de chaque date
    for (let d = 1; d <= daysInMonth; d++) weekdays.push(excelWeekday(new Date(year, monthIdx, d)));
    const countWeekday = (wd) => weekdays.filter((w) => w === wd).length;
    const C = (col, row) => cells[numToCol(col) + row];

    const habits = []; let section = '';
    for (let r = L.firstRow; r <= L.lastRow; r++) {
      if (str(C(L.sectionCol, r))) section = str(C(L.sectionCol, r));
      const name = str(C(L.nameCol, r));
      if (!name) continue;
      const usual = L.timeCol ? fmtTime(C(L.timeCol, r)) : '';
      const planned = [], times = [];
      for (let i = 0; i < 7; i++) {
        const c = C(L.plan1 + i, r); const v = c && c.v;
        const on = v != null && v !== '' && v !== 0;
        planned.push(on);
        times.push(on ? (typeof v === 'string' && fmtTime(c)) || usual : '');
      }
      const checks = [];
      for (let d = 1; d <= daysInMonth; d++) checks.push(isX(C(L.day1 + d - 1, r)));
      habits.push({ row: r, section, name, time: usual, domain: str(C(L.domainCol, r)), planned, times, checks });
    }

    const weekly = []; const monthly = [];
    if (L.weekly) {
      const W = L.weekly;
      for (let r = W.firstRow; r <= W.lastRow; r++) {
        const name = str(C(W.nameCol, r));
        if (name) {
          const e = W.objCol ? C(W.objCol, r) : null; let objective = 0, weekday = null;
          const m = e && e.f ? /\$?([A-Z]{1,3})\$?\d+/.exec(e.f) : null;
          if (m) { const idx = colToNum(m[1]) - L.plan1; if (idx >= 0 && idx < 7) { weekday = idx + 1; objective = countWeekday(weekday); } }
          if (!weekday && e && typeof e.v === 'number') objective = e.v;
          weekly.push({ row: r, name, time: W.timeCol ? fmtTime(C(W.timeCol, r)) : '', weekday, domain: str(C(W.domainCol, r)), objective, weeks: W.sem.map((col) => isX(C(col, r))) });
        }
        if (L.monthly) { const task = str(C(L.monthly.labelCol, r)); if (task) monthly.push({ row: r, name: task, done: isX(C(L.monthly.doneCol, r)) }); }
      }
    }
    const cols = { day1: L.day1, week: L.weekly ? L.weekly.sem.map(numToCol) : [], monthDone: L.monthly ? numToCol(L.monthly.doneCol) : null };
    return { sheetName, monthIdx, monthName: MONTHS_FR[monthIdx], year, daysInMonth, weekdays, habits, weekly, monthly, cols, hasTimes: !!L.timeCol };
  }

  function planningModel(cells) {
    if (!cells) return null;
    const blocks = [];
    const cols = {};
    Object.keys(cells).forEach((a) => { const { col, row } = splitAddr(a); (cols[col] = cols[col] || []).push(row); });
    Object.entries(cols).forEach(([col, rows]) => {
      rows.sort((a, b) => a - b); let cur = null;
      rows.forEach((row) => {
        const text = str(cells[numToCol(+col) + row]); if (!text) return;
        const head = /^jour\s*\d+\s*[·.\-:]\s*(\S+)/i.exec(text);
        if (head || /^notes/i.test(text)) {
          cur = { title: text, weekday: head ? norm(head[1]) : 'notes', items: [], row, col: +col };
          blocks.push(cur);
        } else if (cur) {
          const m = /^(\d{1,2}h\d{0,2})\s*[·\-–:]\s*(.+)$/.exec(text);
          cur.items.push(m ? { time: m[1], label: m[2] } : { time: '', label: text });
        }
      });
    });
    const title = str(cells.B2);
    return { title, blocks };
  }

  async function buildModel(wb) {
    const trackers = [];
    for (const s of wb.sheets) {
      const cells = await readSheet(wb, s.name);
      const L = cells && layout(cells);
      if (L && L.monthIdx >= 0) trackers.push(trackerModel(cells, s.name, L));
    }
    const planSheet = wb.sheets.find((s) => /planning/i.test(s.name));
    const planning = planSheet ? planningModel(await readSheet(wb, planSheet.name)) : null;
    return { version: 2, trackers, planning };
  }

  /* ---------- Écriture d'une cellule dans le XML, sans rien toucher d'autre ---------- */
  function escapeXml(s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  function setCell(xml, addr, value) {
    const { col, row } = splitAddr(addr);
    const newCell = (attrs) => {
      const s = /\ss="(\d+)"/.exec(attrs || '');
      const sAttr = s ? ` s="${s[1]}"` : '';
      return value ? `<c r="${addr}"${sAttr} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>` : `<c r="${addr}"${sAttr}/>`;
    };
    // 1) la cellule existe
    const cellRe = new RegExp(`<c r="${addr}"((?:\\s[^>]*?)?)(\\/>|>[\\s\\S]*?<\\/c>)`);
    if (cellRe.test(xml)) return xml.replace(cellRe, (_, attrs) => newCell(attrs));
    if (!value) return xml; // rien à effacer
    // 2) la ligne existe
    const rowRe = new RegExp(`<row r="${row}"((?:\\s[^>]*?)?)(\\/>|>([\\s\\S]*?)<\\/row>)`);
    const rm = rowRe.exec(xml);
    if (rm) {
      const attrs = rm[1]; const inner = rm[3] || '';
      const cells = []; const re = /<c r="([A-Z]+)\d+"[\s\S]*?(?:\/>|<\/c>)/g; let m;
      let insertAt = inner.length;
      while ((m = re.exec(inner))) { if (colToNum(m[1]) > col) { insertAt = m.index; break; } }
      const newInner = inner.slice(0, insertAt) + newCell('') + inner.slice(insertAt);
      return xml.slice(0, rm.index) + `<row r="${row}"${attrs}>${newInner}</row>` + xml.slice(rm.index + rm[0].length);
    }
    // 3) créer la ligne à la bonne place
    const sd = /<sheetData\s*\/>|<sheetData>([\s\S]*?)<\/sheetData>/.exec(xml);
    if (!sd) throw new Error('Structure de feuille non reconnue');
    const inner = sd[1] || ''; const re = /<row r="(\d+)"/g; let m; let insertAt = inner.length;
    while ((m = re.exec(inner))) { if (+m[1] > row) { insertAt = m.index; break; } }
    const newInner = inner.slice(0, insertAt) + `<row r="${row}">${newCell('')}</row>` + inner.slice(insertAt);
    return xml.slice(0, sd.index) + `<sheetData>${newInner}</sheetData>` + xml.slice(sd.index + sd[0].length);
  }

  /* edits : { "NomFeuille": { "F16": "x", "G16": "" } } */
  async function applyEdits(buf, edits) {
    const wb = await openWorkbook(buf);
    for (const [sheetName, cellsMap] of Object.entries(edits)) {
      const sh = wb.sheets.find((s) => s.name === sheetName);
      if (!sh) continue;
      let xml = await wb.zip.file(sh.path).async('string');
      for (const [addr, val] of Object.entries(cellsMap)) xml = setCell(xml, addr, val);
      wb.zip.file(sh.path, xml, { createFolders: false });
    }
    // Excel recalcule toutes les formules à l'ouverture
    let wbXml = await wb.zip.file('xl/workbook.xml').async('string');
    if (/<calcPr\b/.test(wbXml) && !/fullCalcOnLoad=/.test(wbXml)) wbXml = wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
    else if (!/<calcPr\b/.test(wbXml)) {
      const after = /<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/.exec(wbXml);
      if (after) wbXml = wbXml.slice(0, after.index) + '<calcPr fullCalcOnLoad="1"/>' + wbXml.slice(after.index);
    }
    wb.zip.file('xl/workbook.xml', wbXml, { createFolders: false });
    return wb.zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  global.XL = { MODEL_VERSION: 2, timeMinutes, openWorkbook, buildModel, applyEdits, setCell, numToCol, colToNum, MONTHS_FR, norm, excelWeekday };
})(typeof window !== 'undefined' ? window : globalThis);

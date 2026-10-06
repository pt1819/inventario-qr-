// Logica pura dell'inventario: lettura codici, date, stato scadenze.
// Nessun accesso al DOM, così si può verificare anche fuori dal browser.
(function (root) {
  'use strict';

  /* ---------- Date (sempre in formato AAAA-MM-GG, senza fuso orario) ---------- */
  function todayISO() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function toDayNum(iso) { // giorni dal 1970, per confronti e differenze
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 : null;
  }
  function fromDayNum(n) {
    const d = new Date(n * 86400000);
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }
  function addDays(iso, days) { const n = toDayNum(iso); return n == null ? '' : fromDayNum(n + days); }
  function fmtIT(iso) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? m[3] + '/' + m[2] + '/' + m[1] : ''; }
  function lastDayOfMonth(y, mo) { return new Date(Date.UTC(y, mo, 0)).getUTCDate(); } // mo 1-12

  // Interpreta date scritte a mano o lette da Excel: 31/12/2027, 31-12-27, 2027-12-31, 12/2027, numero seriale Excel
  function parseDate(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number' && isFinite(v)) return fromDayNum(Math.round(v) - 25569); // seriale Excel
    const s = String(v).trim();
    let m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return valid(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/.exec(s))) return valid(year(+m[3]), +m[2], +m[1]);
    if ((m = /^(\d{1,2})[\/.\-](\d{4})$/.exec(s))) { const y = +m[2], mo = +m[1]; return valid(y, mo, lastDayOfMonth(y, mo)); }
    if ((m = /^(\d{4})[\/.\-](\d{1,2})$/.exec(s))) { const y = +m[1], mo = +m[2]; return valid(y, mo, lastDayOfMonth(y, mo)); }
    return '';
    function year(y) { return y < 100 ? 2000 + y : y; }
    function valid(y, mo, d) {
      if (mo < 1 || mo > 12 || d < 1 || d > lastDayOfMonth(y, mo)) return '';
      return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    }
  }

  /* ---------- Codici GS1 (DataMatrix / GS1-128 sulle confezioni dei produttori) ---------- */
  // Lunghezze fisse degli Application Identifier più comuni; gli altri sono variabili fino al separatore.
  const GS = '\u001d';
  const FIXED = { '00': 18, '01': 14, '02': 14, '11': 6, '12': 6, '13': 6, '15': 6, '16': 6, '17': 6, '20': 2 };
  const VARIABLE2 = { '10': 20, '21': 20, '22': 20, '30': 8, '37': 8, '90': 30, '91': 90, '92': 90, '93': 90, '94': 90, '95': 90, '96': 90, '97': 90, '98': 90, '99': 90 };
  const VARIABLE3 = { '240': 30, '241': 30, '242': 6, '250': 30, '251': 30, '400': 30, '401': 30, '403': 30 };

  function gs1Date(yymmdd) {
    const y = 2000 + +yymmdd.slice(0, 2), mo = +yymmdd.slice(2, 4); let d = +yymmdd.slice(4, 6);
    if (mo < 1 || mo > 12) return '';
    if (d === 0) d = lastDayOfMonth(y, mo); // giorno 00 = fine mese
    return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }

  function parseGS1(raw) {
    let s = String(raw);
    s = s.replace(/^\][A-Za-z]\d/, '');              // prefisso di simbologia (]d2, ]C1, ]Q3)
    const out = {};
    // Forma leggibile: (01)08012345678901(17)271231(10)ABC123
    if (/^\(\d{2,4}\)/.test(s)) {
      const re = /\((\d{2,4})\)([^(]*)/g; let m;
      while ((m = re.exec(s))) out[m[1]] = m[2].trim();
      return finish(out);
    }
    if (!/^(01|02|00)\d/.test(s)) return null;
    let i = 0, guard = 0;
    while (i < s.length && guard++ < 30) {
      if (s[i] === GS) { i++; continue; }
      const a2 = s.substr(i, 2), a3 = s.substr(i, 3), a4 = s.substr(i, 4);
      let ai, len, fixed;
      if (FIXED[a2]) { ai = a2; len = FIXED[a2]; fixed = true; }
      else if (/^3[1-6]\d\d$/.test(a4)) { ai = a4; len = 6; fixed = true; }   // misure (peso, volume…)
      else if (a4 === '7003') { ai = a4; len = 10; fixed = true; }
      else if (VARIABLE3[a3]) { ai = a3; len = VARIABLE3[a3]; fixed = false; }
      else if (VARIABLE2[a2]) { ai = a2; len = VARIABLE2[a2]; fixed = false; }
      else break;
      i += ai.length;
      let val;
      if (fixed) { val = s.substr(i, len); i += len; if (s[i] === GS) i++; }
      else { const end = s.indexOf(GS, i); const stop = end === -1 ? s.length : Math.min(end, i + len); val = s.slice(i, stop); i = stop; }
      out[ai] = val;
    }
    return finish(out);

    function finish(o) {
      const gtin = o['01'] || o['02'];
      if (!gtin) return null;
      return {
        code: gtin,
        lot: (o['10'] || '').trim(),
        exp: o['17'] && /^\d{6}$/.test(o['17']) ? gs1Date(o['17']) : '',
        gs1: true
      };
    }
  }

  /* ---------- Lettura di qualsiasi codice scansionato ---------- */
  // 1) GS1 del produttore  2) QR interno "CODICE;Prodotto;Ditta;Lotto;Scadenza"  3) codice semplice
  function parseScan(raw) {
    const s = String(raw || '').replace(/[\r\n]+$/, '').trim();
    if (!s) return null;
    const g = parseGS1(s);
    if (g) return { code: g.code, lot: g.lot, exp: g.exp, name: '', maker: '', source: 'gs1' };
    const parts = s.split(/[;|\t]/).map(x => x.trim());
    if (parts.length > 1 && parts[0]) {
      return { code: parts[0], name: parts[1] || '', maker: parts[2] || '', lot: parts[3] || '', exp: parseDate(parts[4] || ''), source: 'qr' };
    }
    return { code: s, name: '', maker: '', lot: '', exp: '', source: 'plain' };
  }

  /* ---------- Stato di un lotto ---------- */
  // Data limite = la più vicina tra scadenza del produttore e "aperto + giorni di stabilità".
  function lotStatus(lot, product, warnDays, today) {
    today = today || todayISO();
    const stab = product && +product.stab > 0 ? +product.stab : 0;
    const openLimit = lot.opened && stab ? addDays(lot.opened, stab) : '';
    let limit = lot.exp || '';
    let reason = 'scadenza';
    if (openLimit && (!limit || toDayNum(openLimit) < toDayNum(limit))) { limit = openLimit; reason = 'apertura'; }
    const days = limit ? toDayNum(limit) - toDayNum(today) : null;
    let level = 'ok';
    if (days == null) level = 'nd';
    else if (days < 0) level = 'scaduto';
    else if (days <= (warnDays == null ? 30 : warnDays)) level = 'inscadenza';
    return { limit, openLimit, days, level, reason };
  }

  /* ---------- Scadenza e lotto dal testo dell'etichetta (OCR) ---------- */
  const MONTHS = { GEN: 1, JAN: 1, FEB: 2, MAR: 3, APR: 4, MAG: 5, MAY: 5, GIU: 6, JUN: 6, LUG: 7, JUL: 7, AGO: 8, AUG: 8,
    SET: 9, SEP: 9, SEPT: 9, OTT: 10, OCT: 10, NOV: 11, DIC: 12, DEC: 12 };
  // Parole che di solito precedono la scadenza o la data di produzione
  const EXP_WORDS = /(EXP(IRY|IRATION|\.)?|USE\s*BY|USE\s*BEFORE|BEST\s*BEFORE|SCAD(ENZA|\.)?|UTILIZZARE\s*ENTRO|DA\s*USARE\s*ENTRO|VALID[AO]?\s*FINO|VERWENDBAR|VERFALL|CADUCIDAD|CAD\.?|PEREMP|DLU|⌛|⧖)/g;
  const MFG_WORDS = /(MFG|MFD|MANUF|PROD(\.|UCTION|OTTO|UZIONE)?|FABBR|DATE\s*OF\s*MANUF|HERST|FAB\.|⚒)/g;

  // Corregge gli scambi tipici dell'OCR dentro i numeri (O→0, I/l→1, S→5, B→8)
  function fixDigits(s) { return s.replace(/[OoQD]/g, '0').replace(/[Il|!]/g, '1').replace(/[Ss]/g, '5').replace(/B/g, '8').replace(/Z/g, '2'); }
  function mk(y, mo, d) {
    if (y < 100) y += 2000;
    if (y < 2000 || y > 2099 || mo < 1 || mo > 12) return '';
    const last = lastDayOfMonth(y, mo);
    if (d == null) d = last;
    if (d < 1 || d > last) return '';
    return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  }

  function findDates(text) {
    const T = String(text || '').toUpperCase();
    const out = [];
    const N = '[0-9OQDIl|!SBZ]'; // cifre più i caratteri che l'OCR confonde con le cifre
    const add = (iso, index, len, kind, re) => { if (iso) out.push({ iso, index, len, kind }); else re.lastIndex = index + 1; };
    let m;
    // 2027-12-31, 2027/12/31, 2027.12.31
    let re = new RegExp(`(?<![0-9])([0-9]${N}{3})\\s?[-/.]\\s?(${N}{1,2})\\s?[-/.]\\s?(${N}{1,2})(?![0-9])`, 'g');
    while ((m = re.exec(T))) { const [y, mo, d] = [m[1], m[2], m[3]].map(x => +fixDigits(x)); add(mk(y, mo, d), m.index, m[0].length, 'ymd', re); }
    // 31/12/2027, 31.12.27, 31-12-2027 (giorno prima del mese, uso europeo)
    re = new RegExp(`(?<![0-9])([0-9]${N}?)\\s?[-/.]\\s?(${N}{1,2})\\s?[-/.]\\s?(${N}{4}|${N}{2})(?![0-9])`, 'g');
    while ((m = re.exec(T))) { const [d, mo, y] = [m[1], m[2], m[3]].map(x => +fixDigits(x)); add(mk(y, mo, d), m.index, m[0].length, 'dmy', re); }
    // 2027-12 (formato ISO 15223 "anno-mese")
    re = new RegExp(`(?<![0-9])(20${N}{2})\\s?[-/.]\\s?(${N}{1,2})(?![0-9/.\\-])`, 'g');
    while ((m = re.exec(T))) add(mk(+fixDigits(m[1]), +fixDigits(m[2])), m.index, m[0].length, 'ym', re);
    // 12/2027
    re = new RegExp(`(?<![0-9/.\\-])([0-9]${N}?)\\s?[-/.]\\s?(20${N}{2})(?![0-9])`, 'g');
    while ((m = re.exec(T))) add(mk(+fixDigits(m[2]), +fixDigits(m[1])), m.index, m[0].length, 'my', re);
    // 31 DEC 2027, DEC 2027, 2027 DEC 31, 31-DIC-27
    const MN = Object.keys(MONTHS).sort((a, b) => b.length - a.length).join('|');
    re = new RegExp(`(?:([0-9]${N}?)[\\s\\-./]*)?\\b(${MN})[A-Z]*\\.?[\\s\\-./]*(20${N}{2}|${N}{2})(?![0-9])`, 'g');
    while ((m = re.exec(T))) add(mk(+fixDigits(m[3]), MONTHS[m[2]], m[1] ? +fixDigits(m[1]) : null), m.index, m[0].length, 'mon', re);
    re = new RegExp(`(20${N}{2})[\\s\\-./]*(${MN})[A-Z]*\\.?(?:[\\s\\-./]*(${N}{1,2}))?(?![0-9])`, 'g');
    while ((m = re.exec(T))) add(mk(+fixDigits(m[1]), MONTHS[m[2]], m[3] ? +fixDigits(m[3]) : null), m.index, m[0].length, 'mon', re);

    // Togli i doppioni sovrapposti (tieni la lettura più lunga)
    out.sort((a, b) => a.index - b.index || b.len - a.len);
    const res = [];
    for (const c of out) { const prev = res[res.length - 1]; if (prev && c.index < prev.index + prev.len) { if (c.len > prev.len) res[res.length - 1] = c; continue; } res.push(c); }
    return res;
  }

  // Restituisce la scadenza più probabile, le alternative e il lotto letto
  function extractLabelInfo(text, today) {
    today = today || todayISO();
    const T = String(text || '').toUpperCase();
    const dates = findDates(T);
    const lastPos = (re, s) => { re.lastIndex = 0; let m, pos = -1; while ((m = re.exec(s))) pos = m.index; return pos; };
    // Guarda le parole subito prima della data: conta quella più vicina
    const kind = c => {
      const before = T.slice(Math.max(0, c.index - 28), c.index);
      const e = lastPos(EXP_WORDS, before), f = lastPos(MFG_WORDS, before);
      return e < 0 && f < 0 ? '' : (e > f ? 'exp' : 'mfg');
    };
    const t = toDayNum(today);
    const scored = dates.map(c => {
      let score = 0;
      const k = kind(c), exp = k === 'exp', mfg = k === 'mfg';
      if (exp) score += 10;
      if (mfg) score -= 10;
      const dn = toDayNum(c.iso);
      if (dn > t) score += 2;
      if (dn > t + 365 * 12 || dn < t - 365 * 3) score -= 20; // date implausibili
      return Object.assign({}, c, { score, exp, mfg });
    }).filter(c => c.score > -20);
    // A parità di punteggio vince la data più lontana (la scadenza viene dopo la produzione)
    scored.sort((a, b) => b.score - a.score || b.iso.localeCompare(a.iso));
    const uniq = [];
    for (const c of scored) if (!uniq.some(u => u.iso === c.iso)) uniq.push(c);
    const best = uniq[0] && !uniq[0].mfg ? uniq[0] : null;

    let lot = '';
    const lm = /(?:\bLOT(?:TO)?|\bBATCH|\bCH\.?-?B\.?|\bCHARGE)\b\s*(?:N[O°º.]?\s*)?[:#.\-]?\s*([A-Z0-9][A-Z0-9\-\/.]{2,24})/.exec(T);
    if (lm) lot = lm[1].replace(/[.\-\/]+$/, '');
    return { exp: best ? best.iso : '', confident: !!(best && best.exp), candidates: uniq.slice(0, 4).map(c => c.iso), lot };
  }

  const api = { todayISO, toDayNum, fromDayNum, addDays, fmtIT, parseDate, parseGS1, parseScan, lotStatus, findDates, extractLabelInfo };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Core = api;
})(typeof self !== 'undefined' ? self : this);

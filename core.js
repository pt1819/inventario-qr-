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

  const api = { todayISO, toDayNum, fromDayNum, addDays, fmtIT, parseDate, parseGS1, parseScan, lotStatus };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Core = api;
})(typeof self !== 'undefined' ? self : this);

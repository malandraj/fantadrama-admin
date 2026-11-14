/* FantaDrama · script dell'agente per Leghe Fantacalcio (SOLA LETTURA).
 * Caricato dall'agente nella pagina della lega: intercetta gli Excel scaricati con i pulsanti
 * "Scarica Excel" / "Scarica ora" / "Esporta XLSX" senza salvarli sul computer, li legge con
 * import_fd.js e prepara i dati compressi da passare alla pagina admin. Non invia nulla a nessuno.
 */
(() => {
  if (window.fdAgente) return;
  const dl = [];
  const blobs = new Map();
  const origCreate = URL.createObjectURL;
  URL.createObjectURL = function (b) { const u = origCreate.call(URL, b); if (b instanceof Blob) blobs.set(u, b); return u; };
  const grab = async (a) => {
    const href = a.href, nome = a.download || decodeURIComponent(href.split('/').pop().split('?')[0]);
    let blob = blobs.get(href);
    if (!blob) blob = await (await fetch(href, { credentials: 'include' })).blob();
    dl.push({ nome, buf: new Uint8Array(await blob.arrayBuffer()), ora: new Date().toISOString() });
  };
  const origClick = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    if (this.hasAttribute('download') || (this.href || '').startsWith('blob:')) { grab(this); return; }
    return origClick.call(this);
  };
  document.addEventListener('click', (e) => {
    const a = e.target.closest && e.target.closest('a[download]');
    if (a) { e.preventDefault(); grab(a); }
  }, true);

  const gz = async (txt) => {
    const s = new Blob([txt]).stream().pipeThrough(new CompressionStream('gzip'));
    const u = new Uint8Array(await new Response(s).arrayBuffer());
    let bin = ''; for (let i = 0; i < u.length; i += 0x8000) bin += String.fromCharCode(...u.subarray(i, i + 0x8000));
    return btoa(bin);
  };

  window.fdAgente = {
    versione: 2,
    scaricati: () => dl.map((d) => ({ nome: d.nome, kb: Math.round(d.buf.length / 1024), ora: d.ora })),
    svuota: () => { dl.length = 0; return 0; },
    // legge tutti gli Excel intercettati: { dati, riepilogo } — i controlli veri li rifà il server
    leggi: async () => {
      const dati = { formazioni: [] }, riepilogo = [];
      for (const d of dl) {
        const tipo = await fdRiconosci(d.buf);
        if (tipo === 'formazioni') { const f = await fdImportFormazioni(d.buf, d.nome); dati.formazioni.push({ giornata: f.giornata, partite: f.partite }); riepilogo.push(`formazioni g${f.giornata}: ${f.partite.length} partite, controlli ${f.ok ? 'ok' : 'NON ok'}`); }
        else if (tipo === 'classifica') { dati.classifica = await fdImportClassifica(d.buf, d.nome); riepilogo.push(`classifica dopo la ${dati.classifica.giornata}ª`); }
        else if (tipo === 'calendario') { dati.calendario = await fdImportCalendario(d.buf, d.nome); riepilogo.push(`calendario: ${dati.calendario.length} partite, ${dati.calendario.filter((c) => c.risultato).length} risultati`); }
        else if (tipo === 'rose') { dati.rose = await fdImportRose(d.buf, d.nome); riepilogo.push(`rose: ${Object.keys(dati.rose).length} squadre`); }
        else riepilogo.push(`${d.nome}: non riconosciuto`);
      }
      window.__fdDati = dati;
      return riepilogo;
    },
    // dati letti → gzip+base64 a pezzi, da passare alla pagina admin
    pacchetto: async (pezzo = 12000) => {
      const b = await gz(JSON.stringify(window.__fdDati || {}));
      window.__fdPezzi = []; for (let i = 0; i < b.length; i += pezzo) window.__fdPezzi.push(b.slice(i, i + pezzo));
      return { pezzi: window.__fdPezzi.length, caratteri: b.length };
    },
    pezzo: (i) => window.__fdPezzi[i],
    // deposita i dati letti nella cassetta di Supabase (solo scrittura, chiave pubblica). Restituisce l'id da dare alla pagina admin.
    invia: async () => {
      const id = crypto.randomUUID ? crypto.randomUUID() : ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, (c) => (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16));
      const payload = await gz(JSON.stringify(window.__fdDati || {}));
      const r = await fetch('https://omjgkgzbzgzagtafxatx.supabase.co/rest/v1/import_inbox', {
        method: 'POST',
        headers: { apikey: 'sb_publishable_L_XtsavtG6tfv9Q67Zyokw_OZnO5JPY', 'content-type': 'application/json', prefer: 'return=minimal' },
        body: JSON.stringify({ id, payload, nota: 'agente ' + new Date().toISOString() }),
      });
      if (!r.ok) throw new Error('deposito fallito: ' + r.status + ' ' + (await r.text()).slice(0, 200));
      return { id, caratteri: payload.length };
    },
  };
})();
/* FantaDrama — importatore export "formazioni" di Leghe Fantacalcio, in JavaScript.
 * Stessa logica di scripts/import_league.py (parse_lineups), ma gira anche dentro la pagina
 * del sito: legge lo .xlsx (zip) con DecompressionStream, senza librerie esterne.
 * Uso: await fdImportFormazioni(Uint8Array, nomeFile) -> { giornata, partite: [...] }
 * SOLA LETTURA: non tocca la pagina, non invia nulla.
 */
(function (root) {
  "use strict";
  const BONUS_KEYS = ["G", "GS", "AU", "RS", "RSb", "RP", "A", "AS", "AG", "PI"];
  const GREY = "00A0A7B1";

  async function unzip(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eocd = buf.length - 22;
    while (eocd > 0 && dv.getUint32(eocd, true) !== 0x06054b50) eocd--;
    const cdOff = dv.getUint32(eocd + 16, true), n = dv.getUint16(eocd + 10, true);
    const files = {}; let p = cdOff;
    for (let i = 0; i < n; i++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), lho = dv.getUint32(p + 42, true);
      files[new TextDecoder().decode(buf.slice(p + 46, p + 46 + nl))] = { method, csize, lho };
      p += 46 + nl + el + cl;
    }
    return async (name) => {
      const f = files[name]; if (!f) return null;
      const lnl = dv.getUint16(f.lho + 26, true), lel = dv.getUint16(f.lho + 28, true);
      const data = buf.slice(f.lho + 30 + lnl + lel, f.lho + 30 + lnl + lel + f.csize);
      if (f.method === 0) return new TextDecoder().decode(data);
      const s = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
      return await new Response(s).text();
    };
  }

  const unesc = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (a, d) => String.fromCharCode(+d)).replace(/&amp;/g, "&");
  // come openpyxl: un colore a 6 cifre RGB diventa ARGB con alfa "00"
  const argb = (v) => (v && /^[0-9A-Fa-f]{6}$/.test(v) ? "00" + v.toUpperCase() : v ? v.toUpperCase() : v);
  const attr = (tag, name) => { const m = tag.match(new RegExp("\\s" + name + '="([^"]*)"')); return m ? m[1] : null; };
  const texts = (xml) => [...xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) => unesc(m[1])).join("");
  const colNum = (letters) => [...letters].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const colLet = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  async function readSheet(buf) {
    const read = await unzip(buf);
    const ssXml = await read("xl/sharedStrings.xml");
    const ss = ssXml ? [...ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => texts(m[1])) : [];
    const stXml = await read("xl/styles.xml") || "";
    const section = (tag) => { const m = stXml.match(new RegExp("<" + tag + "[^>]*>([\\s\\S]*?)</" + tag + ">")); return m ? m[1] : ""; };
    const fonts = [...section("fonts").matchAll(/<font>([\s\S]*?)<\/font>|<font\/>/g)].map((m) => { const c = (m[1] || "").match(/<color\s[^>]*>/); return c ? argb(attr(c[0], "rgb")) : null; });
    const fills = [...section("fills").matchAll(/<fill>([\s\S]*?)<\/fill>|<fill\/>/g)].map((m) => {
      const pf = (m[1] || "").match(/<patternFill[^>]*>/); const type = pf ? attr(pf[0], "patternType") : null;
      const fg = (m[1] || "").match(/<fgColor\s[^>]*>/);
      return { type: type && type !== "none" ? type : null, rgb: fg ? argb(attr(fg[0], "rgb")) : null };
    });
    const xfs = [...section("cellXfs").matchAll(/<xf\s[^>]*?\/?>/g)].map((m) => ({ font: +(attr(m[0], "fontId") || 0), fill: +(attr(m[0], "fillId") || 0) }));
    const sh = await read("xl/worksheets/sheet1.xml");
    const cells = {};
    for (const m of sh.matchAll(/<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const head = "<c " + m[1] + ">", r = attr(head, "r"), t = attr(head, "t"), s = +(attr(head, "s") || 0);
      const inner = m[2] || "";
      const vm = inner.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/);
      let v = null;
      if (t === "s" && vm) v = ss[+vm[1]];
      else if (t === "inlineStr") v = texts(inner);
      else if ((t === "str" || t === "e") && vm) v = unesc(vm[1]);
      else if (t === "b" && vm) v = vm[1] === "1";
      else if (vm) v = Number(vm[1]);
      const xf = xfs[s] || { font: 0, fill: 0 };
      cells[r] = { v, font: fonts[xf.font] || null, fill: fills[xf.fill] || { type: null, rgb: null } };
    }
    const at = (row, col) => cells[colLet(col) + row] || { v: null, font: null, fill: { type: null, rgb: null } };
    const maxRow = Math.max(0, ...Object.keys(cells).map((k) => +k.replace(/[A-Z]+/, "")));
    const colNum2 = (letters) => [...letters].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
    const maxCol = Math.max(0, ...Object.keys(cells).map((k) => colNum2(k.replace(/\d+/, ""))));
    return { at, maxRow, maxCol };
  }

  const num = (v) => (typeof v === "number" ? v : null);
  const clean = (s) => String(s).replace(/\s+/g, " ").trim();
  const slug = (s) => s.normalize("NFKD").replace(/[^\x00-\x7F]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  const isScore = (v) => typeof v === "string" && /^\d+-\d+$/.test(v);

  function parseSide(sheet, fname, start, c0) {
    const { at, maxRow } = sheet;
    const col = (o) => c0 + o;
    const players = [], modifiers = [];
    let section = null, r = start + 5;
    while (r <= maxRow) {
      const a = at(r, col(0)).v;
      if (typeof a === "string" && a.trim() === "Panchina") { section = "panchina"; r++; continue; }
      if (typeof a === "string" && a.trim() === "Modificatori") {
        r++;
        while (r <= maxRow) {
          const nm = at(r, col(0)).v;
          if (!nm || typeof nm !== "string" || !nm.startsWith("Modificatore")) break;
          modifiers.push({ nome: nm.trim(), valore: num(at(r, col(5)).v), ref: `${fname}!${colLet(col(5))}${r}` });
          r++;
        }
        break;
      }
      const nameCell = at(r, col(1)), name = nameCell.v;
      if (a == null && name == null) break;
      if (section == null) section = "titolare";
      if (isScore(at(r, 17).v)) break;
      if (name == null) { r++; continue; }
      const counted = !(nameCell.font === GREY);
      const fill = at(r, col(2)).fill;
      let card = null;
      if (fill && fill.type) { if (fill.rgb === "FFFACC15") card = "ammonito"; else if (fill.rgb === "FFDC2626") card = "espulso"; }
      const bonus = {};
      BONUS_KEYS.forEach((k, i) => { const v = num(at(r, col(6 + i)).v); if (v) bonus[k] = v; });
      const v = at(r, col(4)).v, fv = at(r, col(5)).v;
      players.push({
        nome: String(name).trim(), ruolo: a ? String(a).trim() : null, club: at(r, col(3)).v,
        voto: num(v), fv: num(fv), sv: v === "SV", bonus, stato: section, conteggiato: counted, cartellino: card,
        ref: `${fname}!${colLet(col(1))}${r}:${colLet(col(5))}${r}`,
      });
      r++;
    }
    return { players, modifiers };
  }

  async function fdImportFormazioni(buf, fname) {
    const sheet = await readSheet(buf);
    const title = sheet.at(1, 1).v;
    const mm = String(title || "").match(/Giornata\s+(\d+)/);
    if (!mm) throw new Error("Intestazione non riconosciuta: " + title);
    const md = +mm[1];
    const partite = [];
    for (let r = 1; r <= sheet.maxRow; r++) {
      const q = sheet.at(r, 17).v;
      if (!isScore(q)) continue;
      const home = clean(sheet.at(r, 1).v), away = clean(sheet.at(r, 18).v);
      const [gh, ga] = q.split("-").map(Number);
      const hs = parseSide(sheet, fname, r, 1), as = parseSide(sheet, fname, r, 18);
      partite.push({
        giornata: md,
        casa: { squadra: home, id: slug(home), gol: gh, punti: num(sheet.at(r + 2, 6).v), modulo: sheet.at(r + 2, 2).v, giocatori: hs.players, modificatori: hs.modifiers, ref: `${fname}!F${r + 2}` },
        trasferta: { squadra: away, id: slug(away), gol: ga, punti: num(sheet.at(r + 2, 23).v), modulo: sheet.at(r + 2, 19).v, giocatori: as.players, modificatori: as.modifiers, ref: `${fname}!W${r + 2}` },
        ref: `${fname}!Q${r}`,
      });
    }
    // controllo interno: per ogni formazione, fantavoti conteggiati + modificatori = totale, e 11 conteggiati
    const controlli = [];
    for (const m of partite) for (const s of ["casa", "trasferta"]) {
      const t = m[s];
      const tot = t.giocatori.filter((p) => p.conteggiato).reduce((a, p) => a + (p.fv || 0), 0) + t.modificatori.reduce((a, x) => a + (x.valore || 0), 0);
      const n = t.giocatori.filter((p) => p.conteggiato).length;
      controlli.push({ squadra: t.squadra, ok: Math.abs(tot - t.punti) < 0.01 && n === 11, totale: t.punti, calcolato: Math.round(tot * 100) / 100, conteggiati: n });
    }
    return { giornata: md, partite, controlli, ok: controlli.every((c) => c.ok) };
  }

  // Calendario (come parse_calendar): due blocchi di colonne, A-E e G-K
  async function fdImportCalendario(buf, fname) {
    const { at, maxRow } = await readSheet(buf);
    const out = [];
    for (const c0 of [1, 7]) {
      let cur = null;
      for (let r = 1; r <= maxRow; r++) {
        const v = at(r, c0).v;
        if (typeof v === "string" && v.includes("Giornata lega")) {
          cur = { giornata: parseInt(v.match(/^(\d+)/)[1], 10), serie_a: at(r, c0 + 2).v };
          continue;
        }
        if (cur && typeof v === "string" && at(r, c0 + 3).v) {
          const res = at(r, c0 + 4).v;
          out.push({ giornata: cur.giornata, giornata_serie_a: cur.serie_a, casa: clean(v), trasferta: clean(at(r, c0 + 3).v),
            risultato: isScore(res) ? res : null, ref: `${fname}!${colLet(c0)}${r}` });
        }
      }
    }
    return out.sort((a, b) => a.giornata - b.giornata);
  }

  // Classifica (come parse_standings)
  async function fdImportClassifica(buf, fname) {
    const { at, maxRow, maxCol } = await readSheet(buf);
    let hdr = null; const righe = [];
    for (let r = 1; r <= maxRow; r++) {
      const vals = []; for (let c = 1; c <= maxCol; c++) vals.push(at(r, c).v);
      if (vals.includes("Pos") && vals.includes("Squadra")) { hdr = {}; vals.forEach((v, i) => { if (v) hdr[v] = i; }); continue; }
      if (hdr && typeof vals[hdr["Pos"]] === "number") {
        const g = (k) => vals[hdr[k]];
        const team = clean(g("Squadra"));
        righe.push({ pos: Math.trunc(g("Pos")), squadra: team, id: slug(team), g: g("G"), v: g("V"), n: g("N"), p: g("P"), gf: g("Gf"), gs: g("Gs"),
          dr: g("Dr"), pt: g("Pt."), fantapunti: g("Pt. Totali"), ref: `${fname}!A${r}:L${r}` });
      }
    }
    return { giornata: righe.length ? Math.max(...righe.map((x) => x.g)) : null, righe, fonte: fname };
  }

  // Rose (come parse_rosters): coppie di colonne squadra / "costo"
  async function fdImportRose(buf, fname) {
    const { at, maxRow, maxCol } = await readSheet(buf);
    const rose = {};
    for (let c = 1; c <= maxCol; c++) {
      const team = at(1, c).v;
      if (team && at(1, c + 1).v === "costo") {
        const lst = [];
        for (let r = 2; r <= maxRow; r++) {
          const n = at(r, c).v;
          if (n && !["totale", "total", "crediti residui", "residuo"].includes(String(n).trim().toLowerCase()))
            lst.push({ nome: String(n).trim(), costo: at(r, c + 1).v, ref: `${fname}!${colLet(c)}${r}` });
        }
        rose[slug(clean(team))] = lst;
      }
    }
    return rose;
  }

  // Riconosce il tipo di export dal contenuto (non dal nome del file)
  async function fdRiconosci(buf) {
    const { at, maxRow, maxCol } = await readSheet(buf);
    const a1 = at(1, 1).v;
    if (typeof a1 === "string" && /Giornata\s+\d+/.test(a1)) return "formazioni";
    for (let r = 1; r <= Math.min(maxRow, 12); r++) {
      const vals = []; for (let c = 1; c <= maxCol; c++) vals.push(at(r, c).v);
      if (vals.includes("Pos") && vals.includes("Squadra")) return "classifica";
      if (vals.some((v) => typeof v === "string" && v.includes("Giornata lega"))) return "calendario";
    }
    for (let c = 1; c <= maxCol; c++) if (at(1, c).v === "costo") return "rose";
    return null;
  }

  /* Verifica i caricamenti contro i dati della lega e prepara la lega aggiornata.
   * up = { formazioni: [risultati di fdImportFormazioni], classifica, calendario, rose } (tutti facoltativi) */
  function fdVerifica(L, up) {
    const errori = [], avvisi = [], ok = [];
    const ids = new Set(L.squadre.map((t) => t.id));
    const merged = { ...L, importato: { ...L.importato } };
    const nuove = (up.formazioni || []).slice().sort((a, b) => a.giornata - b.giornata);
    const giornateNuove = nuove.map((f) => f.giornata);
    if (new Set(giornateNuove).size !== giornateNuove.length) errori.push("Hai caricato due volte la stessa giornata.");
    for (const f of nuove) {
      const bad = f.controlli.filter((c) => !c.ok);
      if (bad.length) errori.push(`Giornata ${f.giornata}: i totali non tornano per ${bad.map((c) => `${c.squadra} (file ${c.totale}, calcolato ${c.calcolato}, ${c.conteggiati} nel punteggio)`).join("; ")}.`);
      else ok.push(`Giornata ${f.giornata}: ${f.partite.length} partite, ${f.controlli.length} formazioni con i totali al centesimo.`);
      const sconosciute = f.partite.flatMap((m) => [m.casa, m.trasferta]).filter((x) => !ids.has(x.id)).map((x) => x.squadra);
      if (sconosciute.length) errori.push(`Giornata ${f.giornata}: squadre sconosciute nella lega (${[...new Set(sconosciute)].join(", ")}).`);
    }
    if (nuove.length) {
      const days = new Set(giornateNuove);
      merged.partite = L.partite.filter((m) => !days.has(m.giornata)).concat(...nuove.map((f) => f.partite)).sort((a, b) => a.giornata - b.giornata);
      merged.importato.giornate = [...new Set(merged.partite.map((m) => m.giornata))].sort((a, b) => a - b);
      merged.importato.partite = merged.partite.length;
      const last = Math.max(...L.importato.giornate);
      for (const g of giornateNuove) if (g > last + 1) avvisi.push(`Manca la giornata ${last + 1}: hai caricato la ${g}.`);
      for (const g of giornateNuove) if (L.importato.giornate.includes(g)) avvisi.push(`La giornata ${g} c'era già: verrà sostituita con quella caricata.`);
    }
    const played = [...new Set(merged.partite.map((m) => m.giornata))];
    const maxPlayed = Math.max(...played);
    if (up.calendario) {
      const cal = up.calendario;
      const nomi = new Set(L.squadre.map((t) => t.nome));
      const fuori = [...new Set(cal.flatMap((c) => [c.casa, c.trasferta]).filter((n) => !nomi.has(n)))];
      if (fuori.length) errori.push(`Calendario: squadre sconosciute (${fuori.join(", ")}).`);
      let diversi = 0, mancanti = 0;
      for (const m of merged.partite) {
        const c = cal.find((x) => x.giornata === m.giornata && x.casa === m.casa.squadra && x.trasferta === m.trasferta.squadra);
        if (!c) { diversi++; continue; }
        if (c.risultato == null) { if (giornateNuove.includes(m.giornata)) mancanti++; }
        else if (c.risultato !== `${m.casa.gol}-${m.trasferta.gol}`) diversi++;
      }
      if (diversi) errori.push(`Calendario: ${diversi} partite non coincidono con le formazioni.`);
      else ok.push(`Calendario: ${cal.length} partite, risultati coerenti con le formazioni.`);
      if (mancanti) avvisi.push(`Calendario senza i risultati della giornata appena caricata (${mancanti} partite): probabilmente è stato scaricato prima del calcolo. Va bene lo stesso.`);
      merged.calendario = cal;
    }
    if (up.classifica) {
      const cls = up.classifica;
      if (cls.giornata > maxPlayed) errori.push(`La classifica è aggiornata alla ${cls.giornata}ª, ma mancano le formazioni di quella giornata.`);
      else {
        const bad = [];
        for (const row of cls.righe) {
          let v = 0, n = 0, p = 0, gf = 0, ga = 0, pts = 0;
          for (const m of merged.partite) {
            if (m.giornata > cls.giornata) continue;
            for (const [a, b] of [["casa", "trasferta"], ["trasferta", "casa"]]) {
              if (m[a].id !== row.id) continue;
              gf += m[a].gol; ga += m[b].gol; pts += m[a].punti;
              if (m[a].gol > m[b].gol) v++; else if (m[a].gol === m[b].gol) n++; else p++;
            }
          }
          const mine = [v, n, p, gf, ga, v * 3 + n, Math.round(pts * 100) / 100].join(",");
          const off = [row.v, row.n, row.p, row.gf, row.gs, row.pt, row.fantapunti].join(",");
          if (mine !== off) bad.push(row.squadra);
        }
        if (bad.length) errori.push(`Classifica dopo la ${cls.giornata}ª: non coincide con i risultati per ${bad.join(", ")}.`);
        else ok.push(`Classifica dopo la ${cls.giornata}ª: coerente con i risultati di tutte le ${cls.righe.length} squadre.`);
      }
      merged.classifica = cls;
    }
    if (up.rose) {
      const fuori = Object.keys(up.rose).filter((k) => !ids.has(k));
      if (fuori.length) avvisi.push(`Rose: squadre non riconosciute (${fuori.join(", ")}).`);
      ok.push(`Rose: ${Object.keys(up.rose).length} squadre, ${Object.values(up.rose).reduce((a, r) => a + r.length, 0)} giocatori.`);
      merged.rose = up.rose;
    }
    let evento = null;
    if (nuove.length) evento = { giornata: Math.max(...giornateNuove), tipo: "post" };
    else if (up.classifica || up.calendario || up.rose) evento = { giornata: maxPlayed + 1, tipo: "vigilia" };
    if (evento && evento.tipo === "vigilia" && !(merged.calendario || []).some((c) => c.giornata === evento.giornata)) evento = null;
    return { ok, avvisi, errori, valido: !errori.length && (nuove.length || up.classifica || up.calendario || up.rose), merged, evento };
  }

  root.fdRiconosci = fdRiconosci;
  root.fdVerifica = fdVerifica;
  root.fdImportFormazioni = fdImportFormazioni;
  root.fdImportCalendario = fdImportCalendario;
  root.fdImportClassifica = fdImportClassifica;
  root.fdImportRose = fdImportRose;
  if (typeof module !== "undefined") module.exports = { fdImportFormazioni, fdImportCalendario, fdImportClassifica, fdImportRose, fdRiconosci, fdVerifica };
})(typeof window !== "undefined" ? window : globalThis);

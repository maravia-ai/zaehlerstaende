"use strict";

/* Zählerstände – Web-App mit Excel (OneDrive) als Datenquelle.
 *
 * Excel-Datei (siehe config.js):
 *   Blatt "Zähler":     A Gebäude | B Medium | C Zählernummer | D Einheit
 *   Tabelle "Ablesungen" (Blatt "Ablesungen"):
 *     A Datum | B Gebäude | C Medium | D Zählernummer | E Zählerstand | F Einheit |
 *     G Verbrauch | H Tage | I Verbrauch pro Tag | J Foto | K Notiz | L Erfasst am
 * Verbrauch = Zählerstand − letzter Zählerstand desselben Zählers (am oder vor dem Ablesedatum).
 * Bei geänderter Zählernummer (Zählerwechsel) wird kein Verbrauch berechnet.
 * Zähler mit mehreren Registern (z. B. Zweirichtungszähler 1.8.0/2.8.0) schreiben je Register eine Zeile
 * mit eigenem Medium (siehe config.js).
 */

const C = window.APP_CONFIG;
const SCOPES = ["Files.ReadWrite", "Calendars.ReadWrite"];
const GRAPH = "https://graph.microsoft.com/v1.0";
// Kennzeichnung des von der App angelegten Kalendertermins (eigene, feste GUID)
const EXT_PROP = "String {3b8e5d21-7c4f-4a9e-b6d2-91f0c8a7e5b3} Name ZaehlerstaendeErinnerung";
const DEMO = new URLSearchParams(location.search).has("demo");
const LS_ITEM = "zaehler.itemId";
const TABLE = "Ablesungen";
const SHEET_ZAEHLER = "Zähler";

const HEAD_ABL = ["Datum", "Gebäude", "Medium", "Zählernummer", "Zählerstand", "Einheit",
  "Verbrauch", "Tage", "Verbrauch pro Tag", "Foto", "Notiz", "Erfasst am"];
// Optionale Spalte M "Link": in allen Zeilen dieselbe Formel, macht die URL in "Foto" anklickbar.
// (Keine individuelle Formel je Zeile: Excel würde sie als berechnete Spalte auf alle Zeilen übertragen.)
const LINK_HEAD = "Link";
const LINK_FORMULA = '=IF(Ablesungen[[#This Row],[Foto]]="","",HYPERLINK(Ablesungen[[#This Row],[Foto]],"Foto öffnen"))';
const A = { datum: 0, geb: 1, med: 2, nr: 3, stand: 4, einheit: 5, verbrauch: 6, tage: 7, proTag: 8, foto: 9, notiz: 10, erfasst: 11 };

let pca = null;
let account = null;
let state = { meters: [], readings: [], blankRow: false, cols: 12, hasLink: false, loadedAt: null, error: null, itemId: null };

const $app = document.getElementById("app");
const $backdrop = document.getElementById("backdrop");
const $sheet = document.getElementById("sheet");
const $toast = document.getElementById("toast");

/* ---------- Hilfsfunktionen ---------- */

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86400000;

function dateToSerial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH) / DAY); }
function todaySerial() { const n = new Date(); return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()); }
function nowSerial() {
  const n = new Date();
  const ms = Date.UTC(n.getFullYear(), n.getMonth(), n.getDate(), n.getHours(), n.getMinutes(), n.getSeconds());
  return Math.round(((ms - EXCEL_EPOCH) / DAY) * 1e6) / 1e6;
}
function serialToUTCDate(s) { return new Date(EXCEL_EPOCH + Math.floor(s) * DAY); }
function isoDate(s) { return serialToUTCDate(s).toISOString().slice(0, 10); }
function isoToSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  return m ? dateToSerial(+m[1], +m[2], +m[3]) : null;
}
function fmtDate(s) {
  return serialToUTCDate(s).toLocaleDateString("de-DE", { timeZone: "UTC", day: "2-digit", month: "2-digit", year: "numeric" });
}
function monthKey(s) { return isoDate(s).slice(0, 7); }
function monthName(s) {
  return serialToUTCDate(s).toLocaleDateString("de-DE", { timeZone: "UTC", month: "long", year: "numeric" });
}
function fmtNum(v, digits = 3) {
  return Number(v).toLocaleString("de-DE", { minimumFractionDigits: 0, maximumFractionDigits: digits });
}
function round(v, digits) { const f = 10 ** digits; return Math.round(v * f) / f; }
function esc(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function plural(n, one, many) { return `${n} ${Math.abs(n) === 1 ? one : many}`; }
// "12345,678" / "12345.678" / "12.345,678" → 12345.678; leer → null; ungültig → NaN
function parseDecimal(s) {
  let t = String(s).trim().replace(/\s/g, "");
  if (t === "") return null;
  if (t.includes(",")) t = t.replace(/\./g, "").replace(",", ".");
  if (!/^\d+(\.\d+)?$/.test(t)) return NaN;
  return Number(t);
}
function meterKey(geb, med) { return geb + "|" + med; }

function toast(msg, ms = 2600) {
  $toast.textContent = msg;
  $toast.classList.add("show");
  clearTimeout(toast.t);
  toast.t = setTimeout(() => $toast.classList.remove("show"), ms);
}

/* ---------- Microsoft Graph ---------- */

async function getToken() {
  if (!account) throw new Error("Nicht angemeldet.");
  try {
    const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || e.errorCode === "no_tokens_found") {
      await pca.acquireTokenRedirect({ scopes: SCOPES, account });
      return new Promise(() => {}); // Seite wird umgeleitet
    }
    throw e;
  }
}

async function graph(method, path, body, contentType) {
  const token = await getToken();
  const raw = contentType && body instanceof Blob;
  const res = await fetch(GRAPH + path, {
    method,
    headers: { Authorization: "Bearer " + token, ...(body ? { "Content-Type": raw ? contentType : "application/json" } : {}) },
    body: body ? (raw ? body : JSON.stringify(body)) : undefined
  });
  if (res.status === 204) return null;
  const data = await res.json().catch(() => null);
  if (!res.ok) {
    const msg = data?.error?.message || res.statusText;
    const err = new Error(`Microsoft Graph (${res.status}): ${msg}`);
    err.status = res.status;
    err.code = data?.error?.code;
    throw err;
  }
  return data;
}

function drivePath(p) { return p.split("/").filter(Boolean).map(encodeURIComponent).join("/"); }
function wb(id) { return `/me/drive/items/${id}/workbook`; }
function ws(id, name) { return `${wb(id)}/worksheets('${encodeURIComponent(name.replace(/'/g, "''"))}')`; }
function tbl(id) { return `${wb(id)}/tables('${TABLE}')`; }

async function getItemId(force) {
  let id = null;
  try { id = force ? null : localStorage.getItem(LS_ITEM); } catch { /* ignorieren */ }
  if (id) return id;
  let item;
  try {
    item = await graph("GET", `/me/drive/root:/${drivePath(C.filePath)}?$select=id,name`);
  } catch (e) {
    if (e.status === 404) throw new Error(`Excel-Datei nicht gefunden:\nOneDrive/${C.filePath}\n\nBitte Datei dort ablegen oder den Pfad in config.js anpassen.`);
    throw e;
  }
  try { localStorage.setItem(LS_ITEM, item.id); } catch { /* ignorieren */ }
  return item.id;
}

async function readWorkbook(id) {
  const [z, t] = await Promise.all([
    graph("GET", `${ws(id, SHEET_ZAEHLER)}/range(address='A1:D50')?$select=values`),
    graph("GET", `${tbl(id)}/range?$select=values,formulas`)
  ]);
  return { zaehler: z.values, ablesungen: t.values, formeln: t.formulas };
}

async function readAll() {
  let id = await getItemId(false);
  try {
    return { id, ...(await readWorkbook(id)) };
  } catch (e) {
    if (e.status !== 404) throw e;
    // Datei evtl. verschoben/ersetzt → ID neu ermitteln
    id = await getItemId(true);
    try {
      return { id, ...(await readWorkbook(id)) };
    } catch (e2) {
      if (e2.status === 404) throw new Error(`Blatt "${SHEET_ZAEHLER}" oder Tabelle "${TABLE}" in der Excel-Datei nicht gefunden.\n${e2.message}`);
      throw e2;
    }
  }
}

/* ---------- Excel-Inhalte auswerten ---------- */

function parseZaehler(values) {
  const h = values[0].map(v => String(v).trim().toLowerCase());
  if (h[0] !== "gebäude" || h[1] !== "medium" || !h[2].startsWith("zählernummer")) {
    throw new Error(`Blatt "${SHEET_ZAEHLER}": Kopfzeile unerwartet. Erwartet: A "Gebäude", B "Medium", C "Zählernummer".`);
  }
  const byKey = new Map();
  values.slice(1).forEach((r, i) => {
    const geb = String(r[0]).trim(), med = String(r[1]).trim();
    if (geb && med) byKey.set(meterKey(geb, med), { row: i + 2, nummer: String(r[2] ?? "").trim() });
  });
  const meters = [];
  for (const g of C.gebaeude) {
    for (const m of g.zaehler) {
      const z = byKey.get(meterKey(g.name, m.medium));
      const regs = (m.register || [{ medium: m.medium, obis: "", label: "" }])
        .map(r => ({ med: r.medium, obis: r.obis || "", label: r.label || "" }));
      meters.push({ geb: g.name, med: m.medium, einheit: m.einheit, nachkomma: m.nachkomma, regs,
        nummer: z?.nummer || "", zRow: z?.row || null });
    }
  }
  return meters;
}

function parseAblesungen(values, formulas) {
  const h = values[0].map(v => String(v).trim());
  const bad = HEAD_ABL.findIndex((x, i) => h[i] !== x);
  if (bad >= 0) {
    throw new Error(`Tabelle "${TABLE}": Spalte ${String.fromCharCode(65 + bad)} heißt "${h[bad] ?? ""}", erwartet "${HEAD_ABL[bad]}".`);
  }
  const rows = values.slice(1);
  const blankRow = rows.length === 1 && rows[0].every(v => v === "" || v === null);
  const cols = values[0].length;
  const hasLink = h[HEAD_ABL.length] === LINK_HEAD;
  const readings = [];
  rows.forEach((r, i) => {
    if (typeof r[A.datum] !== "number" || typeof r[A.stand] !== "number") return;
    const num = v => (typeof v === "number" ? v : null);
    readings.push({
      idx: i,
      datum: Math.floor(r[A.datum]),
      geb: String(r[A.geb]).trim(),
      med: String(r[A.med]).trim(),
      nummer: String(r[A.nr] ?? "").trim(),
      stand: r[A.stand],
      verbrauch: num(r[A.verbrauch]),
      tage: num(r[A.tage]),
      proTag: num(r[A.proTag]),
      foto: String(formulas?.[i + 1]?.[A.foto] || r[A.foto] || "").trim(),
      notiz: String(r[A.notiz] ?? "").trim(),
      erfasst: num(r[A.erfasst]) ?? 0
    });
  });
  readings.sort((a, b) => a.datum - b.datum || a.erfasst - b.erfasst || a.idx - b.idx);
  return { readings, blankRow, cols, hasLink };
}

function readingsOf(geb, med, list = state.readings) { return list.filter(r => r.geb === geb && r.med === med); }

// Letzte Ablesung am oder vor dem Datum (bei gleichem Datum: zuletzt erfasste)
function previousReading(list, datum) {
  let prev = null;
  for (const r of list) if (r.datum <= datum) prev = r;
  return prev;
}

function avgPerDay(list) {
  const v = list.filter(r => r.proTag !== null && r.proTag >= 0).slice(-6);
  return v.length ? v.reduce((a, r) => a + r.proTag, 0) / v.length : null;
}

/* ---------- Neue Ablesung berechnen ---------- */

function evaluate(meter, list, datum, stand, nummer) {
  const prev = previousReading(list, datum);
  const res = { prev, verbrauch: null, tage: null, proTag: null, wechsel: false, warnings: [], later: list.some(r => r.datum > datum) };
  if (!prev) return res;
  if (prev.nummer && nummer && prev.nummer !== nummer) { res.wechsel = true; return res; }
  const d = meter.nachkomma;
  res.verbrauch = round(stand - prev.stand, d);
  res.tage = datum - prev.datum;
  if (res.tage > 0) res.proTag = round(res.verbrauch / res.tage, 3);
  if (res.verbrauch < 0) {
    res.warnings.push(`Der Stand ist kleiner als die letzte Ablesung (${fmtNum(prev.stand, d)} ${meter.einheit} am ${fmtDate(prev.datum)}). Tippfehler? Bei einem neuen Zähler bitte die Zählernummer ändern.`);
  } else {
    const avg = avgPerDay(list);
    if (avg && res.proTag !== null && res.tage >= 3 && res.proTag > avg * 3) {
      res.warnings.push(`Ungewöhnlich hoher Verbrauch: ${fmtNum(res.proTag, 3)} ${meter.einheit}/Tag, bisher Ø ${fmtNum(avg, 3)} ${meter.einheit}/Tag.`);
    }
  }
  return res;
}

/* ---------- Schreiben ---------- */

async function ensureFolder(path) {
  const parts = path.split("/").filter(Boolean);
  let parent = "";
  for (const name of parts) {
    const cur = parent ? parent + "/" + name : name;
    try {
      await graph("GET", `/me/drive/root:/${drivePath(cur)}?$select=id`);
    } catch (e) {
      if (e.status !== 404) throw e;
      const url = parent ? `/me/drive/root:/${drivePath(parent)}:/children` : `/me/drive/root/children`;
      await graph("POST", url, { name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" });
    }
    parent = cur;
  }
}

async function uploadPhoto(blob, meter, datum, suffix = "") {
  await ensureFolder(C.fotoOrdner);
  const safe = s => s.normalize("NFKD").replace(/[^\w-]+/g, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
  const name = `${isoDate(datum)}_${safe(meter.geb)}_${safe(meter.med)}${suffix ? "_" + safe(suffix) : ""}.jpg`;
  const item = await graph("PUT",
    `/me/drive/root:/${drivePath(C.fotoOrdner + "/" + name)}:/content?@microsoft.graph.conflictBehavior=rename`,
    blob, "image/jpeg");
  return item.webUrl;
}

// Zeile auf die Spaltenzahl der Tabelle bringen; Spalte "Link" bekommt die einheitliche Formel
function fullRow(values, cols, hasLink) {
  const r = values.slice();
  if (hasLink) r.push(LINK_FORMULA);
  while (r.length < cols) r.push("");
  return r;
}

async function appendRow(id, values, useBlankRow) {
  if (useBlankRow) await graph("PATCH", `${tbl(id)}/rows/itemAt(index=0)`, { values: [values] });
  else await graph("POST", `${tbl(id)}/rows/add`, { values: [values] });
}

async function writeNummer(id, zRow, nummer) {
  await graph("PATCH", `${ws(id, SHEET_ZAEHLER)}/range(address='C${zRow}')`, { values: [[nummer]] });
}

// Foto verkleinern (max. 1600 px, JPEG) – spart Speicher und Upload-Zeit
function shrinkImage(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const max = 1600, f = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement("canvas");
      c.width = Math.round(img.naturalWidth * f);
      c.height = Math.round(img.naturalHeight * f);
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      c.toBlob(b => (b ? resolve(b) : reject(new Error("Foto konnte nicht verarbeitet werden."))), "image/jpeg", 0.82);
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Foto konnte nicht gelesen werden.")); };
    img.src = url;
  });
}

/* ---------- Kalender (monatliche Erinnerung) ---------- */

async function syncMonthlyReminder() {
  const filter = `singleValueExtendedProperties/Any(ep: ep/id eq '${EXT_PROP}' and ep/value eq 'monatlich')`;
  const existing = await graph("GET", `/me/events?$filter=${encodeURIComponent(filter)}&$select=id&$top=50`);
  for (const ev of existing.value || []) await graph("DELETE", `/me/events/${ev.id}`);

  const n = new Date();
  let y = n.getFullYear(), m = n.getMonth() + 1;
  if (n.getDate() >= C.erinnerungTag) { m++; if (m > 12) { m = 1; y++; } }
  const day = Math.min(C.erinnerungTag, 28);
  const d = `${y}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  const [hh, mm] = C.erinnerungUhrzeit.split(":").map(Number);
  const endMin = hh * 60 + mm + 15;
  const t = (h, mi) => `${String(h).padStart(2, "0")}:${String(mi).padStart(2, "0")}:00`;

  await graph("POST", "/me/events", {
    subject: "Zählerstände ablesen",
    body: {
      contentType: "text",
      content: `Gas, Wasser und Strom ablesen:\n${C.gebaeude.map(g => "- " + g.name).join("\n")}\n\nApp: ${location.origin + location.pathname}\n\nAutomatisch erstellt von der Zählerstände-App.`
    },
    start: { dateTime: `${d}T${t(hh, mm)}`, timeZone: C.zeitzone },
    end: { dateTime: `${d}T${t(Math.floor(endMin / 60), endMin % 60)}`, timeZone: C.zeitzone },
    recurrence: {
      pattern: { type: "absoluteMonthly", interval: 1, dayOfMonth: day },
      range: { type: "noEnd", startDate: d }
    },
    isReminderOn: true,
    reminderMinutesBeforeStart: 0,
    showAs: "free",
    singleValueExtendedProperties: [{ id: EXT_PROP, value: "monatlich" }]
  });
  return d;
}

/* ---------- Demo-Modus (ohne Microsoft-Konto, Daten nur im Speicher) ---------- */

const demo = (() => {
  const T = todaySerial();
  const zaehler = [["Gebäude", "Medium", "Zählernummer", "Einheit"]];
  const abl = [HEAD_ABL.slice()];
  const base = { Gas: 18234.512, Wasser: 812.304, Strom: 45210.4, "Strom Einspeisung": 3120.0 };
  const rate = { Gas: 9.5, Wasser: 0.32, Strom: 7.4, "Strom Einspeisung": 8.1 };
  let nr = 1000;
  C.gebaeude.forEach((g, gi) => {
    for (const m of g.zaehler) {
      const num = `DEMO-${++nr}`;
      zaehler.push([g.name, m.medium, num, m.einheit]);
      for (const r of (m.register || [{ medium: m.medium }])) {
        let prev = null;
        for (let k = 3; k >= 1; k--) {
          const d = T - k * 30;
          const stand = round(base[r.medium] * (gi ? 0.5 : 1) + (3 - k) * 30 * rate[r.medium] * (gi ? 0.6 : 1), m.nachkomma);
          const v = prev ? round(stand - prev.stand, m.nachkomma) : "";
          abl.push([d, g.name, r.medium, num, stand, m.einheit, v, prev ? 30 : "", prev ? round(v / 30, 3) : "", "", "", d + 0.4]);
          prev = { stand };
        }
      }
    }
  });
  return { zaehler, abl };
})();

/* ---------- Laden ---------- */

async function load() {
  state.error = null;
  render();
  try {
    let z, a, f = null, id = null;
    if (DEMO) { z = demo.zaehler; a = demo.abl; }
    else ({ id, zaehler: z, ablesungen: a, formeln: f } = await readAll());
    state.itemId = id;
    state.meters = parseZaehler(z);
    Object.assign(state, parseAblesungen(a, f));
    state.loadedAt = new Date();
  } catch (e) {
    console.error(e);
    state.error = e.message || String(e);
  }
  render();
}

/* ---------- Darstellung ---------- */

function regTitle(reg) { return [reg.obis, reg.label].filter(Boolean).join(" "); }

function meterStatus(meter) {
  const mk = monthKey(todaySerial());
  const regs = meter.regs.map(reg => {
    const list = readingsOf(meter.geb, reg.med);
    return { reg, list, last: list[list.length - 1] || null, done: list.some(r => monthKey(r.datum) === mk) };
  });
  const lastDates = regs.filter(x => x.last).map(x => x.last.datum);
  return { regs, done: regs.every(x => x.done), lastDatum: lastDates.length ? Math.max(...lastDates) : null };
}

function rowHtml(meter) {
  const st = meterStatus(meter);
  const badge = st.done ? `<span class="badge done">erfasst</span>` : `<span class="badge open">offen</span>`;
  const d = meter.nachkomma;
  const multi = meter.regs.length > 1;
  const right = st.regs.map(({ reg, last }) => {
    const pre = multi ? `<span class="meta">${esc(reg.obis)}</span> ` : "";
    if (!last) return `<div class="meta">${multi ? esc(reg.obis) + ": " : ""}noch keine Ablesung</div>`;
    const v = `<div class="val">${pre}${fmtNum(last.stand, d)} ${esc(meter.einheit)}</div>`;
    const c = last.verbrauch !== null ? `<div class="meta">+${fmtNum(last.verbrauch, d)}${last.tage ? ` in ${last.tage} T.` : ""}</div>` : "";
    return v + (multi ? "" : c || `<div class="meta">&nbsp;</div>`);
  }).join("");
  return `<button class="row m-${esc(meter.med)}" data-key="${esc(meterKey(meter.geb, meter.med))}">
    <span class="dot"></span>
    <div><div class="name">${esc(meter.med)}${badge}</div>
      <div class="meta">${meter.nummer ? "Nr. " + esc(meter.nummer) : "Zählernummer fehlt"}${st.lastDatum ? " · " + fmtDate(st.lastDatum) : ""}${multi ? " · Zweirichtung" : ""}</div></div>
    <div class="right">${right}</div>
  </button>`;
}

function render() {
  if (!DEMO && (!C.clientId || C.clientId.startsWith("HIER"))) {
    $app.innerHTML = `<div class="center"><h1>Einrichtung fehlt</h1>
      <p>Bitte in <b>config.js</b> die Client-ID aus der Entra-App-Registrierung eintragen.</p>
      <p><a href="?demo">Demo ansehen</a></p></div>`;
    return;
  }
  if (!DEMO && !account) {
    $app.innerHTML = `<div class="center"><h1>Zählerstände</h1>
      <p>Anmelden, um die Zählerstände aus deiner Excel-Datei zu laden.</p>
      <div class="actions"><button class="btn primary" id="login">Mit Microsoft anmelden</button></div></div>`;
    document.getElementById("login").onclick = () => pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
    return;
  }

  let html = `<header><div><h1>Zählerstände</h1>
    <div class="sub">${DEMO ? "Demo-Daten · " : ""}${state.loadedAt ? "Stand " + state.loadedAt.toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" }) : "Lade ..."}</div></div>
    <button class="iconbtn" id="reload">Aktualisieren</button></header>`;

  if (state.error) html += `<div class="error">${esc(state.error)}</div>`;

  if (state.meters.length) {
    const T = todaySerial();
    const nDone = state.meters.filter(m => meterStatus(m).done).length;
    const nAll = state.meters.length;
    html += `<div class="summary">
      <div class="tile ${nDone === nAll ? "ok" : "warn"}"><b>${nDone} / ${nAll}</b><span>Zähler erfasst im ${esc(monthName(T))}</span></div>
      <div class="tile"><b>${state.readings.length}</b><span>Ablesewerte gesamt</span></div></div>`;

    for (const g of C.gebaeude) {
      html += `<h2>${esc(g.name)}</h2><div class="list">${state.meters.filter(m => m.geb === g.name).map(rowHtml).join("")}</div>`;
    }

    html += `<div class="actions">
      <button class="btn" id="sync">Monatliche Erinnerung einrichten</button>
      ${DEMO ? "" : `<button class="btn danger" id="logout">Abmelden</button>`}
    </div>
    <p class="sub" style="margin:12px 4px">Erinnerung am ${C.erinnerungTag}. jedes Monats, ${esc(C.erinnerungUhrzeit)} Uhr, im Microsoft-365-Kalender. Ablesungen zwischendurch sind jederzeit möglich.</p>`;
  } else if (!state.error) {
    html += `<div class="center">Lade Zählerstände ...</div>`;
  }

  $app.innerHTML = html;
  document.getElementById("reload").onclick = load;
  $app.querySelectorAll(".row").forEach(b => b.onclick = () => openDetail(b.dataset.key));
  const sync = document.getElementById("sync");
  if (sync) sync.onclick = async () => {
    if (DEMO) return toast("Im Demo-Modus nicht verfügbar.");
    sync.disabled = true; sync.textContent = "Richte Erinnerung ein ...";
    try { const d = await syncMonthlyReminder(); toast(`Monatliche Erinnerung ab ${fmtDate(isoToSerial(d))} eingetragen.`, 4000); }
    catch (e) { state.error = e.message; render(); }
    finally { sync.disabled = false; sync.textContent = "Monatliche Erinnerung einrichten"; }
  };
  const logout = document.getElementById("logout");
  if (logout) logout.onclick = () => pca.logoutRedirect({ account });
}

/* ---------- Detailansicht / Erfassen ---------- */

function closeDetail() { $backdrop.classList.remove("open"); }
$backdrop.addEventListener("click", e => { if (e.target === $backdrop) closeDetail(); });

function histHtml(meter, regs) {
  const d = meter.nachkomma;
  const multi = regs.length > 1;
  return regs.map(({ reg, list }) => {
    if (!list.length) return "";
    const rows = list.slice(-12).reverse().map(r => {
      const link = /^https?:\/\//.test(r.foto) ? r.foto : (/HYPERLINK\("([^"]+)"/i.exec(r.foto)?.[1] || "");
      return `<tr><td>${fmtDate(r.datum)}</td>
        <td class="r">${fmtNum(r.stand, d)}</td>
        <td class="r">${r.verbrauch !== null ? "+" + fmtNum(r.verbrauch, d) : (r.notiz.startsWith("Zählerwechsel") ? "Wechsel" : "–")}</td>
        <td class="r">${link ? `<a href="${esc(link)}" target="_blank" rel="noopener">Foto</a>` : (r.foto ? "Foto" : "")}</td></tr>`;
    }).join("");
    return `<h2>Letzte Ablesungen${multi ? " · " + esc(regTitle(reg)) : ""}</h2><div class="list"><table class="hist">${rows}</table></div>`;
  }).join("");
}

function openDetail(key) {
  const meter = state.meters.find(m => meterKey(m.geb, m.med) === key);
  if (!meter) return;
  const st = meterStatus(meter);
  const regs = st.regs; // [{ reg, list, last }]
  const d = meter.nachkomma;
  const multi = regs.length > 1;
  const photoBlobs = regs.map(() => null);
  let confirmWarn = false;

  const kv = regs.map(({ reg, list, last }) => {
    if (!last) return "";
    const avg = avgPerDay(list);
    const t = multi ? esc(regTitle(reg)) + ": " : "";
    return `<div><span>${t}Letzter Stand</span><span>${fmtNum(last.stand, d)} ${esc(meter.einheit)}</span></div>
      <div><span>${t}Abgelesen am</span><span>${fmtDate(last.datum)}</span></div>
      ${avg !== null ? `<div><span>${t}Ø pro Tag</span><span>${fmtNum(avg, 3)} ${esc(meter.einheit)}</span></div>` : ""}`;
  }).join("");

  const inputs = regs.map(({ reg, last }, i) => `
    <label class="field" for="stand${i}">${multi ? `Zählerstand ${esc(regTitle(reg))}` : "Zählerstand"}</label>
    <div class="unitwrap"><input type="text" class="big" id="stand${i}" inputmode="decimal" autocomplete="off" placeholder="${last ? fmtNum(last.stand, d) : "0"}"><span>${esc(meter.einheit)}</span></div>
    <div class="preview" id="pv${i}"></div>
    <label class="field">${multi ? `Foto ${esc(regTitle(reg))} (optional, als Nachweis)` : "Foto vom Zähler (optional, als Nachweis)"}</label>
    <div class="photo"><img id="thumb${i}" alt="" hidden><label class="btn" for="foto${i}" id="fotobtn${i}">Foto aufnehmen</label>
      <input type="file" id="foto${i}" class="fotoin" accept="image/*" capture="environment" style="display:none"></div>`).join("");

  $sheet.innerHTML = `<div class="grab"></div>
    <h3>${esc(meter.med)}</h3><div class="sub">${esc(meter.geb)}${meter.nummer ? " · Zähler-Nr. " + esc(meter.nummer) : ""}${multi ? " · Zweirichtungszähler" : ""}</div>
    ${kv ? `<div class="kv">${kv}</div>` : `<p class="hint" style="margin:12px 4px">Erste Ablesung für diesen Zähler. Der Verbrauch wird ab der zweiten Ablesung berechnet.</p>`}

    <div class="grid2" style="margin-top:14px">
      <div><label class="field" for="datum">Ablesedatum</label>
        <input type="date" id="datum" value="${isoDate(todaySerial())}" max="${isoDate(todaySerial())}"></div>
      <div><label class="field" for="nummer">Zählernummer</label>
        <input type="text" id="nummer" value="${esc(meter.nummer)}" autocomplete="off" autocapitalize="characters" placeholder="eintragen"></div>
    </div>
    <p class="hint">Zählernummer nur bei Zählerwechsel ändern. Dann wird kein Verbrauch berechnet; den Endstand des alten Zählers vorher unter der alten Nummer erfassen.</p>
    ${multi ? `<p class="hint">Am Zähler-Display: 1.8.0 = Bezug aus dem Netz, 2.8.0 = Einspeisung ins Netz. Beide Werte sind erforderlich.</p>` : ""}

    ${inputs}
    <div class="preview" id="pvall"></div>

    <label class="field" for="notiz">Notiz (optional)</label>
    <input type="text" id="notiz" autocomplete="off" placeholder="z. B. Zwischenablesung Mieterwechsel">

    <div class="actions" style="margin-top:16px">
      <button class="btn primary" id="save" disabled>Speichern</button>
      <button class="btn" id="close">Schließen</button>
    </div>
    ${histHtml(meter, regs)}`;

  const $ = s => $sheet.querySelector(s);
  const $datum = $("#datum"), $nummer = $("#nummer"), $save = $("#save"), $pvall = $("#pvall");
  const $stands = regs.map((_, i) => $(`#stand${i}`));
  const $pvs = regs.map((_, i) => $(`#pv${i}`));

  function current() {
    return {
      stands: $stands.map(x => parseDecimal(x.value)),
      datum: isoToSerial($datum.value),
      nummer: $nummer.value.trim()
    };
  }

  function update() {
    confirmWarn = false;
    $save.className = "btn primary";
    $save.textContent = "Speichern";
    $pvall.className = "preview"; $pvall.textContent = "";
    const { stands, datum, nummer } = current();
    $pvs.forEach(p => { p.className = "preview"; p.textContent = ""; });
    if (datum === null || datum > todaySerial()) { $pvall.textContent = "Bitte ein gültiges Datum (nicht in der Zukunft) wählen."; $save.disabled = true; return; }
    let complete = true, anyWarn = false;
    regs.forEach(({ list }, i) => {
      const stand = stands[i], $pv = $pvs[i];
      if (stand === null) { complete = false; return; }
      if (Number.isNaN(stand)) { complete = false; $pv.textContent = "Bitte nur Ziffern und ggf. ein Komma eingeben."; return; }
      const ev = evaluate(meter, list, datum, stand, nummer);
      const parts = [`Erfasst wird: ${fmtNum(stand, d)} ${meter.einheit}`];
      if (ev.wechsel) parts.push(`Zählerwechsel (alt: ${ev.prev.nummer}) – kein Verbrauch`);
      else if (ev.verbrauch !== null) {
        parts.push(`Seit ${fmtDate(ev.prev.datum)}: ${fmtNum(ev.verbrauch, d)} ${meter.einheit}` +
          (ev.proTag !== null ? ` (${ev.tage} T., ${fmtNum(ev.proTag, 3)} ${meter.einheit}/Tag)` : ""));
      }
      if (ev.later) parts.push("Hinweis: Es gibt bereits eine spätere Ablesung. Deren Verbrauchswert wird nicht neu berechnet.");
      if (ev.warnings.length) { $pv.className = "preview warn"; parts.push(...ev.warnings); anyWarn = true; }
      else if (ev.verbrauch !== null) $pv.className = "preview ok";
      $pv.innerHTML = parts.map(esc).join("<br>");
    });
    $save.disabled = !complete;
    if (!complete && multi && stands.some(s => s !== null)) $pvall.textContent = "Bitte beide Werte (1.8.0 und 2.8.0) eingeben.";
    if (complete && anyWarn) {
      $save.className = "btn warnbtn";
      $save.textContent = "Trotzdem speichern";
      confirmWarn = true;
    }
  }

  $stands.forEach(x => x.oninput = update);
  $datum.oninput = update;
  $nummer.oninput = update;
  $("#close").onclick = closeDetail;

  regs.forEach((_, i) => {
    $(`#foto${i}`).onchange = async e => {
      const f = e.target.files?.[0];
      if (!f) return;
      const $btn = $(`#fotobtn${i}`);
      $btn.textContent = "Verarbeite ...";
      try {
        photoBlobs[i] = await shrinkImage(f);
        const t = $(`#thumb${i}`);
        t.src = URL.createObjectURL(photoBlobs[i]);
        t.hidden = false;
        $btn.textContent = `Foto ändern (${Math.round(photoBlobs[i].size / 1024)} KB)`;
      } catch (err) {
        photoBlobs[i] = null;
        $btn.textContent = "Foto aufnehmen";
        toast(err.message, 4000);
      }
    };
  });

  $save.onclick = async () => {
    const { stands, datum, nummer } = current();
    if (datum === null || stands.some(s => s === null || Number.isNaN(s))) return;
    const wasWarn = confirmWarn;
    $save.disabled = true; $save.textContent = "Speichere ...";
    try {
      let blankRow = state.blankRow, cols = state.cols, hasLink = state.hasLink;
      if (!DEMO) {
        // Vor dem Schreiben frisch lesen, damit der Verbrauch auf dem aktuellen Stand berechnet wird
        const t = await graph("GET", `${tbl(state.itemId)}/range?$select=values,formulas`);
        const p = parseAblesungen(t.values, t.formulas);
        regs.forEach(x => { x.list = readingsOf(meter.geb, x.reg.med, p.readings); });
        ({ blankRow, cols, hasLink } = p);
      }
      const evs = regs.map((x, i) => evaluate(meter, x.list, datum, stands[i], nummer));
      if (evs.some(ev => ev.warnings.length) && !wasWarn) { update(); return; }

      const fotoUrls = regs.map(() => "");
      if (!DEMO) {
        for (let i = 0; i < regs.length; i++) {
          if (!photoBlobs[i]) continue;
          $save.textContent = multi ? `Lade Foto ${regs[i].reg.obis} hoch ...` : "Lade Foto hoch ...";
          fotoUrls[i] = await uploadPhoto(photoBlobs[i], meter, datum, multi ? regs[i].reg.obis : "");
        }
      }
      const notizIn = $("#notiz").value.trim();
      const rows = regs.map(({ reg }, i) => {
        const ev = evs[i];
        let notiz = notizIn;
        if (ev.wechsel) notiz = `Zählerwechsel (alt: ${ev.prev.nummer})${notiz ? ". " + notiz : ""}`;
        return [datum, meter.geb, reg.med, nummer, stands[i], meter.einheit,
          ev.verbrauch ?? "", ev.tage ?? "", ev.proTag ?? "", fotoUrls[i], notiz, nowSerial()];
      });

      $save.textContent = "Speichere ...";
      if (DEMO) {
        rows.forEach((row, i) => { if (photoBlobs[i]) row[A.foto] = "Foto (Demo)"; demo.abl.push(row); });
        if (nummer !== meter.nummer) {
          const z = demo.zaehler.find(r => r[0] === meter.geb && r[1] === meter.med);
          if (z) z[2] = nummer;
        }
      } else {
        for (let i = 0; i < rows.length; i++) await appendRow(state.itemId, fullRow(rows[i], cols, hasLink), blankRow && i === 0);
        if (nummer !== meter.nummer && meter.zRow) await writeNummer(state.itemId, meter.zRow, nummer);
      }
      closeDetail();
      await load();
      toast(`${meter.med} ${meter.geb}: ${stands.map(s => fmtNum(s, d)).join(" / ")} ${meter.einheit} gespeichert.`);
    } catch (e) {
      $save.disabled = false; $save.className = "btn primary"; $save.textContent = "Speichern";
      $pvall.className = "preview warn";
      $pvall.textContent = e.message;
    }
  };

  update();
  $backdrop.classList.add("open");
}

/* ---------- Start ---------- */

async function start() {
  if (DEMO) { render(); return load(); }
  if (!C.clientId || C.clientId.startsWith("HIER")) return render();
  pca = new msal.PublicClientApplication({
    auth: {
      clientId: C.clientId,
      authority: `https://login.microsoftonline.com/${C.tenantId}`,
      redirectUri: location.origin + location.pathname
    },
    cache: { cacheLocation: "localStorage" }
  });
  try {
    await pca.initialize();
    const res = await pca.handleRedirectPromise();
    account = res?.account || pca.getAllAccounts()[0] || null;
  } catch (e) {
    $app.innerHTML = `<div class="error">Anmeldung fehlgeschlagen:\n${esc(e.message || e)}</div>`;
    return;
  }
  render();
  if (account) load();
}

start();

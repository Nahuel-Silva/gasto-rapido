/* Gasto Rápido — carga rápida de gastos que se guardan en tu Google Sheets. */
"use strict";

/* ---------- helpers ---------- */
const $ = (s) => document.querySelector(s);
const norm = (s) => (s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim();
const pad = (n) => String(n).padStart(2, "0");
const iso = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const today = () => { const d = new Date(); d.setHours(12, 0, 0, 0); return d; };
const fmtMoney = (n) => (n < 0 ? "-$ " : "$ ") + Math.abs(Math.round(n * 100) / 100).toLocaleString("es-AR", { maximumFractionDigits: 2 });
const fmtDate = (s) => { const [y, m, d] = s.split("-"); return `${d}/${m}/${y}`; };
const uid = () => "w" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const cap = (s) => (s ? s.charAt(0).toUpperCase() + s.slice(1) : s);
const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

function parseAmount(raw) {
  if (raw == null) return NaN;
  if (typeof raw === "number") return raw;
  let s = String(raw).trim().replace(/\s/g, "").replace(/\$/g, "").replace(/ARS/i, "");
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  if (s.startsWith("-")) { neg = true; s = s.slice(1); } else if (s.endsWith("-")) { neg = true; s = s.slice(0, -1); }
  if (s.startsWith("+")) s = s.slice(1);
  let mult = 1; if (/k$/i.test(s)) { mult = 1000; s = s.slice(0, -1); }
  if (!/^[\d.,]+$/.test(s) || !/\d/.test(s)) return NaN;
  const hasDot = s.includes("."), hasCom = s.includes(",");
  if (hasDot && hasCom) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (hasCom) {
    const p = s.split(",");
    s = p.length === 2 && p[1].length !== 3 ? p[0] + "." + p[1] : s.replace(/,/g, "");
  } else if (hasDot) {
    const p = s.split(".");
    if (p.length > 2 || (p.length === 2 && p[1].length === 3 && mult === 1)) s = s.replace(/\./g, "");
  }
  const n = parseFloat(s) * mult;
  return isNaN(n) ? NaN : neg ? -n : n;
}
function parseDateAny(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate(), 12);
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 12);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], 12);
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?/);
  if (m) {
    let y = m[3] ? +m[3] : today().getFullYear(); if (y < 100) y += 2000;
    const d = new Date(y, +m[2] - 1, +m[1], 12);
    if (!m[3] && d > new Date(today().getTime() + 86400000 * 2)) d.setFullYear(y - 1);
    return d.getMonth() === +m[2] - 1 ? d : null;
  }
  return null;
}

/* ---------- local storage (todo con try/catch) ---------- */
const LS = { conn: "gr-conexion", queue: "gr-pendientes", cache: "gr-planilla", prefs: "gr-prefs", tab: "gr-tab" };
const lsGet = (k, def) => { try { const v = JSON.parse(localStorage.getItem(k) || "null"); return v ?? def; } catch (e) { return def; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* sin almacenamiento */ } };

let conn = lsGet(LS.conn, null);               // {url, clave}
let queue = lsGet(LS.queue, []);               // movimientos pendientes
let remote = lsGet(LS.cache, null);            // config leída de la planilla
let prefs = lsGet(LS.prefs, { num: "ar", defMedio: "" });
const saveQueue = () => lsSet(LS.queue, queue);

const D = window.GR_DEFAULTS;
const CATS = () => (remote && remote.categorias && remote.categorias.length ? remote.categorias : D.categorias);
const CUENTAS = () => (remote && remote.cuentas && remote.cuentas.length ? remote.cuentas : D.cuentas);
const DEF_MEDIO = () => prefs.defMedio || (remote && remote.cuentaDefecto) || D.cuentaDefecto;

let RULES = null;
function rules() {
  if (!RULES) {
    const src = remote && remote.reglas && remote.reglas.length ? remote.reglas : D.reglas;
    RULES = src.map(([k, c]) => ({ k, c, nk: norm(k) })).filter((r) => r.nk);
  }
  return RULES;
}
function categorize(desc) {
  const d = norm(desc); let best = null;
  for (const r of rules()) if (d.includes(r.nk) && (!best || r.nk.length >= best.nk.length)) best = r;
  return best ? best.c : "";
}

/* ---------- API de la planilla ---------- */
async function api(accion, data = {}) {
  if (!conn) throw new Error("Primero conectá tu planilla.");
  let res;
  try {
    res = await fetch(conn.url, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ ...data, accion, clave: conn.clave }),
      redirect: "follow"
    });
  } catch (e) {
    throw new Error("No hay conexión con la planilla. Revisá internet: lo cargado queda guardado acá.");
  }
  let out;
  try { out = await res.json(); } catch (e) { throw new Error("La dirección del script no respondió bien. ¿Está implementado como aplicación web?"); }
  if (!out.ok) throw new Error(out.error || "Error desconocido");
  return out;
}
function setPill(cls, txt) { const p = $("#sync"); p.className = "pill " + cls; p.querySelector("span").textContent = txt; }

async function loadRemote(quiet) {
  if (!conn) { setPill("local", "Sin conectar"); $("#setup").hidden = false; return; }
  setPill("", "Conectando…");
  try {
    const cfg = await api("config");
    remote = { categorias: cfg.categorias, cuentas: cfg.cuentas, cuentaDefecto: cfg.cuentaDefecto, reglas: cfg.reglas, fijos: cfg.fijos, planilla: cfg.planilla, url: cfg.url };
    lsSet(LS.cache, remote); RULES = null;
    setPill("ok", "Conectado a " + (cfg.planilla || "tu planilla"));
    $("#setup").hidden = true;
    fillSelects(); render(); renderFijos(); renderRules();
    loadResumen();
  } catch (e) {
    setPill("local", "Sin conexión con la planilla");
    if (!quiet) toast(e.message);
    if (/clave|dirección|implementad/i.test(e.message)) $("#setup").hidden = false;
  }
}

/* ---------- items ---------- */
function mk(o) {
  return { id: uid(), createdAt: Date.now(), fecha: o.fecha, tipo: o.tipo || "Gasto", desc: o.desc || "", monto: Math.round(o.monto * 100) / 100, medio: o.medio || DEF_MEDIO(), cat: o.cat ?? "", catSrc: o.catSrc || "auto", origen: o.origen || "página" };
}
function addItems(list, origen) {
  if (!list.length) { toast("No encontré movimientos para agregar"); return; }
  list.forEach((it) => { it.origen = origen; queue.push(it); });
  saveQueue(); render();
  toast(`${list.length} ${list.length === 1 ? "movimiento agregado" : "movimientos agregados"} a la lista`);
}

/* ---------- escribir ---------- */
const MEDIO_ALIAS = { santander: "Santander", sant: "Santander", stdr: "Santander", icbc: "ICBC", efectivo: "Efectivo", efvo: "Efectivo", efe: "Efectivo", cash: "Efectivo", naranja: "Naranja", nx: "Naranja", mp: "Mercado Pago", mercadopago: "Mercado Pago" };
function medioAlias(n) {
  const m = MEDIO_ALIAS[n]; if (m && CUENTAS().includes(m)) return m;
  return CUENTAS().find((c) => norm(c) === n) || (m || null);
}
function parseLine(line) {
  let t = line.trim(); if (!t) return null;
  let tipo = "Gasto", d = today(), medio = null, forced = null;
  if (t.startsWith("+")) { tipo = "Ingreso"; t = t.slice(1).trim(); }
  t = t.replace(/\bmercado\s+pago\b/i, () => { medio = "Mercado Pago"; return " "; });
  t = t.replace(/#(\S+)/g, (m, c) => {
    const n = norm(c); const hit = CATS().find((k) => norm(k) === n) || CATS().find((k) => norm(k).startsWith(n));
    if (hit) forced = hit; return " ";
  });
  const toks = t.split(/\s+/).filter(Boolean);
  let amtIdx = -1, amt = NaN;
  toks.forEach((w, i) => { if (/^\$?\d[\d.,]*k?$/i.test(w) && !/^\d{1,2}\/\d{1,2}/.test(w) && !isNaN(parseAmount(w))) { amtIdx = i; amt = parseAmount(w); } });
  const words = [];
  toks.forEach((w, i) => {
    const n = norm(w);
    if (i === amtIdx) return;
    if (n === "hoy") { d = today(); return; }
    if (n === "ayer") { d = today(); d.setDate(d.getDate() - 1); return; }
    if (n === "anteayer" || n === "antier") { d = today(); d.setDate(d.getDate() - 2); return; }
    if (/^\d{1,2}\/\d{1,2}(\/\d{2,4})?$/.test(w)) { const p = parseDateAny(w); if (p) { d = p; return; } }
    const md = medioAlias(n); if (md) { medio = md; return; }
    words.push(w);
  });
  let desc = cap(words.join(" ").replace(/\s+/g, " ").trim());
  if (isNaN(amt) || amt <= 0) return { error: true, line };
  if (/^(ingreso|sueldo|cobro|cobre|cobré|reintegro|aguinaldo)\b/i.test(desc)) tipo = "Ingreso";
  if (!desc) desc = tipo === "Ingreso" ? "Ingreso" : "Sin descripción";
  return mk({ fecha: iso(d), tipo, desc, monto: amt, medio: medio || DEF_MEDIO(), cat: tipo === "Ingreso" ? "" : forced || categorize(desc), catSrc: forced ? "manual" : "auto" });
}
function quickPreview() {
  const lines = $("#quick").value.split("\n").filter((l) => l.trim());
  const box = $("#qprev"); box.innerHTML = "";
  lines.slice(0, 40).forEach((l) => {
    const r = parseLine(l); const el = document.createElement("div");
    if (!r || r.error) { el.className = "pv err"; el.innerHTML = `<span class="d">${esc(l)}</span><small>falta el monto</small>`; }
    else { el.className = "pv"; el.innerHTML = `<span class="d">${esc(r.desc)} <small>· ${fmtDate(r.fecha).slice(0, 5)} · ${esc(r.medio)} · ${esc(r.tipo === "Ingreso" ? "INGRESO" : r.cat || "sin categoría")}</small></span><span class="m">${r.tipo === "Ingreso" ? "+" : ""}${fmtMoney(r.monto)}</span>`; }
    box.appendChild(el);
  });
}
function quickAdd() {
  const ok = [], bad = [];
  $("#quick").value.split("\n").forEach((l) => { if (!l.trim()) return; const r = parseLine(l); if (!r || r.error) bad.push(l); else ok.push(r); });
  if (ok.length) addItems(ok, "página");
  $("#quick").value = bad.join("\n"); quickPreview();
  if (bad.length) toast(`${bad.length} ${bad.length === 1 ? "línea quedó" : "líneas quedaron"} sin monto`);
}

/* ---------- extracto ---------- */
let X = { aoa: null };
const RX = { fecha: /fecha/, desc: /descrip|concepto|detalle|movimiento|referencia|leyenda|operaci/, deb: /d[eé]bito|debe|egreso|retiro|cargo/, cred: /cr[eé]dito|haber|dep[oó]sito|abono/, monto: /importe|monto|valor|pesos|\$/ };
function detect(aoa) {
  for (let i = 0; i < Math.min(aoa.length, 40); i++) {
    const row = aoa[i].map((c) => norm(String(c ?? "")));
    const f = row.findIndex((c) => RX.fecha.test(c));
    if (f < 0) continue;
    const find = (rx) => row.findIndex((c, j) => j !== f && rx.test(c) && !/saldo/.test(c));
    const m = { fecha: f, desc: find(RX.desc), deb: find(RX.deb), cred: find(RX.cred), monto: find(RX.monto) };
    if (m.deb >= 0 || m.cred >= 0 || m.monto >= 0) {
      if (m.desc < 0) {
        let best = -1, len = 0;
        row.forEach((c, j) => { if (![m.fecha, m.deb, m.cred, m.monto].includes(j)) { const L = aoa.slice(i + 1, i + 15).reduce((a, r) => a + String(r[j] ?? "").length, 0); if (L > len) { len = L; best = j; } } });
        m.desc = best;
      }
      return { head: i, map: m };
    }
  }
  const rows = aoa.slice(0, 60); const ncol = Math.max(0, ...rows.map((r) => r.length));
  let f = -1, mo = -1, de = -1, bf = 0, bm = 0, bd = 0;
  for (let j = 0; j < ncol; j++) {
    let dc = 0, nc = 0, tl = 0;
    rows.forEach((r) => { const v = r[j]; if (parseDateAny(v)) dc++; else if (v !== "" && v != null && !isNaN(parseAmount(v))) nc++; else tl += String(v ?? "").length; });
    if (dc > bf) { bf = dc; f = j; } if (nc > bm) { bm = nc; mo = j; } if (tl > bd) { bd = tl; de = j; }
  }
  if (f >= 0 && mo >= 0 && bf >= 2) return { head: -1, map: { fecha: f, desc: de, deb: -1, cred: -1, monto: mo } };
  return null;
}
function rowsToItems() {
  const { aoa, head, map } = X; if (!aoa || !map) return [];
  const medio = $("#xmedio").value; const out = [];
  const posAs = X.pos || "auto";
  const body = aoa.slice(head + 1);
  const nums = body.map((r) => parseAmount(r[map.monto])).filter((n) => !isNaN(n));
  const mixed = nums.some((n) => n < 0) && nums.some((n) => n > 0);
  body.forEach((r) => {
    const d = parseDateAny(r[map.fecha]); if (!d) return;
    let desc = String(r[map.desc] ?? "").replace(/\s+/g, " ").trim();
    let tipo, n;
    if (map.deb >= 0 || map.cred >= 0) {
      const db = map.deb >= 0 ? Math.abs(parseAmount(r[map.deb])) : NaN, cr = map.cred >= 0 ? Math.abs(parseAmount(r[map.cred])) : NaN;
      if (db > 0) { tipo = "Gasto"; n = db; } else if (cr > 0) { tipo = "Ingreso"; n = cr; } else return;
    } else {
      n = parseAmount(r[map.monto]); if (isNaN(n) || n === 0) return;
      if (posAs === "gasto") tipo = "Gasto";
      else if (posAs === "ingreso") tipo = n < 0 ? "Gasto" : "Ingreso";
      else tipo = mixed ? (n < 0 ? "Gasto" : "Ingreso") : "Gasto";
      n = Math.abs(n);
    }
    if (/^saldo/i.test(desc)) return;
    desc = cap(desc.toLowerCase()).slice(0, 80) || "Movimiento";
    out.push(mk({ fecha: iso(d), tipo, desc, monto: n, medio, cat: tipo === "Ingreso" ? "" : categorize(desc), catSrc: "auto" }));
  });
  return out;
}
function renderMap() {
  const box = $("#xmap"), info = $("#xinfo");
  if (!X.aoa) { box.hidden = true; info.innerHTML = ""; $("#xadd").disabled = true; $("#xadd").textContent = "Agregar movimientos"; return; }
  if (!X.map) { box.hidden = true; $("#xadd").disabled = true; info.innerHTML = `<div class="notice warn">No reconocí las columnas de fecha y monto. Probá descargar el extracto en Excel o CSV.</div>`; return; }
  const hdr = X.head >= 0 ? X.aoa[X.head] : X.aoa[0].map((_, j) => "Columna " + (j + 1));
  const opts = (sel) => `<option value="-1">—</option>` + hdr.map((h, j) => `<option value="${j}" ${sel === j ? "selected" : ""}>${esc(String(h ?? "").slice(0, 30) || "Columna " + (j + 1))}</option>`).join("");
  const single = X.map.deb < 0 && X.map.cred < 0;
  box.innerHTML = `<label>Fecha<select data-k="fecha">${opts(X.map.fecha)}</select></label>
   <label>Descripción<select data-k="desc">${opts(X.map.desc)}</select></label>
   <label>Monto<select data-k="monto">${opts(X.map.monto)}</select></label>
   <label>Débito<select data-k="deb">${opts(X.map.deb)}</select></label>
   <label>Crédito<select data-k="cred">${opts(X.map.cred)}</select></label>
   ${single ? `<label>Montos positivos<select id="xpos"><option value="auto">Detectar</option><option value="gasto">Son gastos</option><option value="ingreso">Son ingresos</option></select></label>` : ""}`;
  box.hidden = false;
  box.querySelectorAll("select[data-k]").forEach((s) => (s.onchange = () => { X.map[s.dataset.k] = +s.value; renderMap(); }));
  const xp = $("#xpos"); if (xp) { xp.value = X.pos || "auto"; xp.onchange = () => { X.pos = xp.value; renderMap(); }; }
  const items = rowsToItems();
  const g = items.filter((i) => i.tipo === "Gasto"), sin = g.filter((i) => !i.cat).length;
  info.innerHTML = items.length
    ? `<div class="notice">Encontré <b>${items.length}</b> movimientos (${g.length} gastos, ${items.length - g.length} ingresos)${sin ? `, ${sin} sin categoría` : ""}. Los revisás en la lista antes de guardar.</div>`
    : `<div class="notice warn">No encontré filas con fecha y monto. Revisá qué columna es cada cosa.</div>`;
  $("#xadd").disabled = !items.length; $("#xadd").textContent = items.length ? `Agregar ${items.length} movimientos` : "Agregar movimientos";
}
function loadAoa(aoa) {
  X = { aoa: aoa.filter((r) => r.some((c) => String(c ?? "").trim() !== "")), head: -1, map: null, pos: "auto" };
  const d = detect(X.aoa); if (d) { X.head = d.head; X.map = d.map; }
  renderMap();
}
function parseTextTable(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  const sep = lines.some((l) => l.includes("\t")) ? "\t" : lines.filter((l) => l.includes(";")).length > lines.length / 2 ? ";" : null;
  if (sep) return lines.map((l) => l.split(sep).map((c) => c.replace(/^"|"$/g, "").trim()));
  if (lines.filter((l) => l.split(",").length >= 3).length > lines.length / 2) return lines.map((l) => l.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map((c) => c.replace(/^"|"$/g, "").trim()));
  return lines.map((l) => {
    const m = l.match(/^\s*(\d{1,2}[\/\-.]\d{1,2}(?:[\/\-.]\d{2,4})?)\s+(.*?)\s+(-?\$?\s?-?[\d.]+,\d{2}-?)(?:\s+-?\$?\s?[\d.]+,\d{2})?\s*$/);
    return m ? [m[1], m[2], m[3]] : [l];
  });
}
async function readFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".txt")) {
    let t = await file.text();
    if (/Ã/.test(t)) t = new TextDecoder("windows-1252").decode(await file.arrayBuffer());
    $("#xpaste").value = ""; loadAoa(parseTextTable(t)); return;
  }
  if (typeof XLSX === "undefined") { toast("No se pudo cargar el lector de Excel. Probá con CSV o pegando las filas."); return; }
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  loadAoa(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" }));
}

/* ---------- lista ---------- */
const selOpts = (list, val) => list.map((v) => `<option ${v === val ? "selected" : ""}>${esc(v)}</option>`).join("");
const dupKey = (it) => `${it.fecha}|${Math.round(it.monto * 100)}|${it.medio}`;
function render() {
  const ae = document.activeElement;
  const focusId = ae && ae.closest && ae.closest(".item")?.dataset.id, focusF = ae && ae.dataset && ae.dataset.f;
  const list = $("#list");
  const seen = new Map(); queue.forEach((it) => seen.set(dupKey(it), (seen.get(dupKey(it)) || 0) + 1));
  let tg = 0, ti = 0, unk = 0;
  if (!queue.length) {
    list.innerHTML = `<div class="empty"><h3>Nada pendiente</h3><p>Escribí tus gastos a la izquierda o subí el extracto del banco. Acá los revisás y con un toque se guardan en tu planilla.</p><div class="ex">cabi 3500<br>vea 143.860 santander<br>ayer kfc 18880</div></div>`;
  } else {
    list.innerHTML = queue.map((it) => {
      const ing = it.tipo === "Ingreso"; if (ing) ti += it.monto; else tg += it.monto;
      const noCat = !ing && !it.cat; if (noCat) unk++;
      const auto = categorize(it.desc);
      const learn = !ing && it.cat && it.catSrc !== "auto" && auto !== it.cat;
      const word = (norm(it.desc).match(/[a-zñ]{3,}/) || [""])[0];
      const meta = [seen.get(dupKey(it)) > 1 ? `<span class="chip warn">¿Repetido?</span>` : "",
        learn && word ? `<button class="chip learn" data-learn="${esc(word)}" data-cat="${esc(it.cat)}">Recordar: «${esc(word)}» → ${esc(it.cat)}</button>` : ""].join("");
      return `<div class="item ${ing ? "ing" : ""}" data-id="${it.id}">
        <input class="f-date" type="date" value="${it.fecha}" data-f="fecha" aria-label="Fecha">
        <input class="f-desc" value="${esc(it.desc)}" data-f="desc" aria-label="Descripción">
        <input class="f-amt amt" inputmode="decimal" value="${(ing ? "+" : "") + it.monto.toLocaleString("es-AR", { maximumFractionDigits: 2 })}" data-f="monto" aria-label="Monto">
        <select class="f-medio" data-f="medio" aria-label="Medio">${selOpts(CUENTAS(), it.medio)}</select>
        <select class="f-cat ${noCat ? "cat-unk" : ""}" data-f="cat" aria-label="Categoría">${ing ? `<option value="">INGRESO</option>` : `<option value="" ${noCat ? "selected" : ""}>SIN CATEGORÍA</option>` + selOpts(CATS(), it.cat)}</select>
        <button class="x" data-del aria-label="Borrar">×</button>
        <div class="meta"><button class="tipo ${ing ? "ing" : ""}" data-tipo>${ing ? "Ingreso" : "Gasto"}</button>${meta}</div>
      </div>`;
    }).join("");
  }
  $("#tg").textContent = fmtMoney(tg); $("#ti").textContent = fmtMoney(ti);
  const nc = $("#navcount"); nc.hidden = !queue.length; nc.textContent = queue.length;
  $("#qcount").textContent = queue.length ? `${queue.length} ${queue.length === 1 ? "movimiento" : "movimientos"} sin guardar` : "Sin movimientos pendientes";
  $("#save").disabled = !queue.length; $("#copy").disabled = !queue.length; $("#clear").hidden = !queue.length;
  $("#save").textContent = queue.length ? `Guardar ${queue.length} en la planilla` : "Guardar en la planilla";
  $("#tbar").hidden = !unk;
  $("#unkinfo").textContent = unk ? `${unk} sin categoría: elegila en la lista (o se guarda como SIN CATEGORÍA y la ponés después).` : "";
  if (focusId && focusF) { const el = list.querySelector(`.item[data-id="${focusId}"] [data-f="${focusF}"]`); if (el) el.focus({ preventScroll: true }); }
}
const findItem = (el) => { const id = el.closest(".item")?.dataset.id; return queue.find((i) => i.id === id); };
$("#list").addEventListener("change", (e) => {
  const it = findItem(e.target); if (!it) return; const f = e.target.dataset.f; let v = e.target.value;
  if (f === "monto") { const n = Math.abs(parseAmount(v)); if (!n) { render(); return; } v = Math.round(n * 100) / 100; }
  if (f === "fecha" && !v) { render(); return; }
  it[f] = v;
  if (f === "cat") it.catSrc = "manual";
  if (f === "desc" && it.catSrc === "auto" && it.tipo !== "Ingreso") it.cat = categorize(v);
  saveQueue(); render();
});
$("#list").addEventListener("click", (e) => {
  const it = findItem(e.target); if (!it) return;
  if (e.target.closest("[data-del]")) {
    const i = queue.indexOf(it); queue.splice(i, 1); saveQueue(); render();
    toast("Movimiento borrado", () => { queue.splice(i, 0, it); saveQueue(); render(); });
  } else if (e.target.closest("[data-tipo]")) {
    it.tipo = it.tipo === "Ingreso" ? "Gasto" : "Ingreso"; it.cat = it.tipo === "Ingreso" ? "" : categorize(it.desc); it.catSrc = "auto"; saveQueue(); render();
  } else if (e.target.dataset.learn) {
    addRule(e.target.dataset.learn, e.target.dataset.cat);
  }
});

/* ---------- guardar ---------- */
$("#save").onclick = async () => {
  if (!conn) { go("ajustes"); toast("Conectá tu planilla para guardar"); return; }
  const b = $("#save"); b.disabled = true; b.innerHTML = `<span class="spin"></span> Guardando…`;
  const batch = queue.slice(0, 500);
  try {
    const out = await api("agregar", { items: batch.map((it) => ({ id: it.id, fecha: it.fecha, tipo: it.tipo, desc: it.desc, monto: it.monto, medio: it.medio, cat: it.cat, origen: it.origen })) });
    const done = new Set(out.ids || []);
    queue = queue.filter((it) => !done.has(it.id)); saveQueue(); render();
    toast(`Guardados ${out.agregados} en la planilla${out.repetidos ? ` (${out.repetidos} ya estaban)` : ""}`);
    loadResumen();
  } catch (e) { toast(e.message); render(); }
};
$("#copy").onclick = async () => {
  const us = prefs.num === "us";
  const text = queue.map((it) => [fmtDate(it.fecha), it.tipo, it.desc.replace(/\t/g, " "), us ? String(it.monto) : String(it.monto).replace(".", ","), it.medio, it.tipo === "Ingreso" ? "" : it.cat || ""].join("\t")).join("\n");
  try { await navigator.clipboard.writeText(text); toast("Copiado. Pegalo en la hoja Movimientos, columna A"); }
  catch (e) { $("#xpaste").value = text; toast("No pude copiar: quedó en la caja de pegado del extracto"); }
};
$("#clear").onclick = () => { const b = queue; queue = []; saveQueue(); render(); toast("Lista vaciada", () => { queue = b; saveQueue(); render(); }); };

/* ---------- resumen del mes ---------- */
async function loadResumen() {
  if (!conn) return;
  try {
    const r = await api("resumen", { mes: iso(today()).slice(0, 7) });
    $("#mes").hidden = false;
    const [y, m] = r.mes.split("-");
    $("#mtitle").textContent = `${cap(MESES[+m - 1])} ${y} en tu planilla`;
    $("#mi").textContent = fmtMoney(r.ingresos); $("#mg").textContent = fmtMoney(r.gastos); $("#mr").textContent = fmtMoney(r.ingresos - r.gastos);
    $("#mr").style.color = r.ingresos - r.gastos < 0 ? "var(--bad)" : "var(--good)";
    const top = (r.categorias || []).slice(0, 7); const max = Math.max(1, ...top.map((c) => c.monto));
    $("#mbars").innerHTML = top.length ? top.map((c) => `<div class="bar"><span class="n">${esc(c.cat)}</span><span class="t"><i style="width:${(c.monto / max) * 100}%"></i></span><span class="v">${fmtMoney(c.monto)}</span></div>`).join("") : `<p class="hint">Todavía no hay gastos este mes.</p>`;
    $("#mlast").innerHTML = (r.ultimos || []).map((u) => `<div><span class="d">${u.fecha ? fmtDate(u.fecha).slice(0, 5) : ""}</span><span class="s">${esc(u.desc)} <small class="hint">· ${esc(u.tipo === "Ingreso" ? "INGRESO" : u.cat || "")}</small></span><span class="m" style="color:${u.tipo === "Ingreso" ? "var(--good)" : "inherit"}">${u.tipo === "Ingreso" ? "+" : ""}${fmtMoney(u.monto)}</span></div>`).join("") || `<p class="hint">Sin movimientos.</p>`;
    renderTarjeta(r.tarjeta);
    if (remote && remote.url) { const a = $("#mopen"); a.href = remote.url; a.hidden = false; }
  } catch (e) { /* el resumen es opcional */ }
}
$("#mref").onclick = () => loadRemote(false);

/* ---------- tarjeta ---------- */
let lastTarjeta = null;
function mesLabel(ym) { const [y, m] = ym.split("-"); return `${MESES[+m - 1].slice(0, 3)} ${y}`; }
function addMonths(ym, n) { const [y, m] = ym.split("-").map(Number); const d = new Date(y, m - 1 + n, 1); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`; }
function renderTarjeta(t) {
  lastTarjeta = t || null;
  const box = $("#tcard");
  if (!t) { box.hidden = true; return; }
  box.hidden = false;
  $("#tv1").textContent = fmtMoney(t.esteMes); $("#tv2").textContent = fmtMoney(t.proximo); $("#tv3").textContent = fmtMoney(t.deuda);
  $("#tpay").disabled = !(t.porDonde && t.porDonde.Tarjeta > 0);
  $("#tpay").textContent = t.porDonde && t.porDonde.Tarjeta > 0 ? `Cargar pago del resumen (${fmtMoney(t.porDonde.Tarjeta)})` : "Nada que pagar este mes";
  $("#tlist").innerHTML = (t.compras || []).map((c) => {
    const estado = c.actual >= 1 ? `cuota ${c.actual}/${c.cuotas}` : `arranca ${mesLabel(c.primerMes)}`;
    return `<div><span class="q">${estado}</span><span class="s">${esc(c.compra)} <small class="hint">· ${esc(c.donde)}</small></span><span class="m">${fmtMoney(c.precio)}</span></div>`;
  }).join("") || `<p class="hint">No tenés cuotas activas.</p>`;
}
$("#tpay").onclick = () => {
  if (!lastTarjeta || !(lastTarjeta.porDonde.Tarjeta > 0)) return;
  addItems([mk({ fecha: iso(today()), tipo: "Gasto", desc: "Pago resumen tarjeta", monto: lastTarjeta.porDonde.Tarjeta, medio: DEF_MEDIO(), cat: CATS().includes("TARJETA") ? "TARJETA" : "", catSrc: "manual" })], "página");
  toast("Agregado a la lista. Ajustá el monto si el resumen vino distinto y guardalo");
};
function fillMeses() {
  const now = iso(today()).slice(0, 7);
  $("#kmes").innerHTML = [0, 1, 2].map((n) => { const ym = addMonths(now, n); return `<option value="${ym}" ${n === 1 ? "selected" : ""}>${cap(mesLabel(ym))}${n === 0 ? " (este mes)" : n === 1 ? " (el que viene)" : ""}</option>`; }).join("");
}
function cuotaPreview() {
  const compra = $("#kcompra").value.trim(); const monto = Math.abs(parseAmount($("#kmonto").value));
  const n = +$("#kcuotas").value; const porCuota = $("#ktipo").value === "cuota";
  const box = $("#kprev");
  if (!monto || !n) { box.hidden = true; $("#kadd").disabled = true; return null; }
  const precio = Math.round((porCuota ? monto : monto / n) * 100) / 100;
  const total = Math.round(precio * n * 100) / 100;
  const ini = $("#kmes").value, fin = addMonths(ini, n - 1);
  box.hidden = false;
  box.innerHTML = n === 1
    ? `Un pago de <b>${fmtMoney(precio)}</b> en el resumen de ${mesLabel(ini)}.`
    : `<b>${n} cuotas de ${fmtMoney(precio)}</b> (total ${fmtMoney(total)}), de ${mesLabel(ini)} a ${mesLabel(fin)}.`;
  $("#kadd").disabled = !compra;
  return { compra, precio, cuotas: n, primerMes: ini, donde: $("#kdonde").value };
}
["#kcompra", "#kmonto"].forEach((s) => $(s).addEventListener("input", cuotaPreview));
["#ktipo", "#kcuotas", "#kmes", "#kdonde"].forEach((s) => $(s).addEventListener("change", cuotaPreview));
$("#kadd").onclick = async () => {
  const c = cuotaPreview(); if (!c || !c.compra) return;
  if (!conn) { go("ajustes"); toast("Conectá tu planilla para guardar compras con tarjeta"); return; }
  const b = $("#kadd"); b.disabled = true; b.innerHTML = `<span class="spin"></span> Guardando…`;
  try {
    await api("cuota", c);
    toast(`Guardada: ${c.compra}, ${c.cuotas === 1 ? "1 pago" : c.cuotas + " cuotas"} de ${fmtMoney(c.precio)}`);
    $("#kcompra").value = ""; $("#kmonto").value = ""; cuotaPreview(); loadResumen();
  } catch (e) {
    toast(/desconocida/i.test(e.message) ? "Actualizá el script de la planilla para usar esta función (ver README)" : e.message);
  }
  b.textContent = "Guardar compra"; cuotaPreview();
};

/* ---------- reglas ---------- */
async function addRule(k, c) {
  k = (k || "").trim().toLowerCase(); if (!k || !c) return;
  const local = () => {
    remote = remote || { categorias: CATS(), cuentas: CUENTAS(), cuentaDefecto: DEF_MEDIO(), reglas: D.reglas.slice(), fijos: [] };
    remote.reglas = (remote.reglas || []).filter(([kk]) => norm(kk) !== norm(k)).concat([[k, c]]);
    lsSet(LS.cache, remote); RULES = null;
    let n = 0; queue.forEach((it) => { if (it.tipo !== "Ingreso" && it.catSrc === "auto") { const c2 = categorize(it.desc); if (c2 !== it.cat) { it.cat = c2; n++; } } });
    saveQueue(); render(); renderRules(); return n;
  };
  if (conn) {
    try { await api("regla", { palabra: k, categoria: c }); const n = local(); toast(`Regla guardada en la planilla: «${k}» → ${c}${n ? ` (actualicé ${n})` : ""}`); }
    catch (e) { toast(e.message); }
  } else { const n = local(); toast(`Regla guardada en este dispositivo: «${k}» → ${c}${n ? ` (actualicé ${n})` : ""}`); }
}
function renderRules() {
  const q = norm($("#rsearch").value);
  const all = rules().slice().sort((a, b) => a.nk.localeCompare(b.nk));
  const shown = all.filter((r) => !q || r.nk.includes(q) || norm(r.c).includes(q)).slice(0, 250);
  $("#rules").innerHTML = shown.map((r) => `<div class="rule"><span class="k">${esc(r.k)}</span><span>${esc(r.c)}</span><span></span></div>`).join("") || `<div class="rule"><span class="hint">Ninguna regla coincide.</span></div>`;
}
$("#radd").onclick = () => { addRule($("#rk").value, $("#rc").value); $("#rk").value = ""; };
$("#rsearch").oninput = renderRules;

/* ---------- fijos ---------- */
function renderFijos() {
  const fijos = (remote && remote.fijos) || [];
  const box = $("#fijos");
  if (!fijos.length) { box.innerHTML = `<p class="hint">${conn ? "No hay fijos cargados en la hoja Fijos." : "Conectá tu planilla para ver tus fijos."}</p>`; $("#fadd").disabled = true; return; }
  box.innerHTML = fijos.map((f, i) => `<label><input type="checkbox" data-i="${i}" ${f.activo ? "checked" : ""}><span>${esc(f.concepto)} <small>· día ${f.dia} · ${esc(f.medio)}</small></span><span class="m">${fmtMoney(f.monto)}</span></label>`).join("");
  updFijosBtn();
}
function updFijosBtn() {
  const fijos = (remote && remote.fijos) || [];
  const sel = [...document.querySelectorAll("#fijos input:checked")].map((c) => fijos[+c.dataset.i]);
  const n = sel.reduce((a, f) => a + f.monto, 0);
  $("#fadd").disabled = !sel.length;
  $("#fadd").textContent = sel.length ? `Agregar ${sel.length} fijos de ${MESES[today().getMonth()]} (${fmtMoney(n)})` : "Agregar fijos";
}
$("#fijos").addEventListener("change", updFijosBtn);
$("#fadd").onclick = () => {
  const fijos = (remote && remote.fijos) || []; const t = today();
  const last = new Date(t.getFullYear(), t.getMonth() + 1, 0).getDate();
  const sel = [...document.querySelectorAll("#fijos input:checked")].map((c) => fijos[+c.dataset.i]);
  addItems(sel.map((f) => mk({ fecha: iso(new Date(t.getFullYear(), t.getMonth(), Math.min(f.dia, last), 12)), tipo: "Gasto", desc: f.concepto, monto: f.monto, medio: f.medio, cat: f.cat, catSrc: "manual" })), "fijos");
};

/* ---------- conexión y ajustes ---------- */
function fillSelects() {
  ["#xmedio", "#sdef"].forEach((s) => ($(s).innerHTML = selOpts(CUENTAS(), DEF_MEDIO())));
  $("#rc").innerHTML = selOpts(CATS(), CATS().includes("SALIDAS") ? "SALIDAS" : CATS()[0]);
  $("#snum").value = prefs.num || "ar";
}
$("#cgo").onclick = async () => {
  const url = $("#curl").value.trim(), clave = $("#cclave").value.trim();
  if (!/^https:\/\/script\.google(usercontent)?\.com\/.+\/exec/.test(url)) { $("#cmsg").textContent = "La dirección tiene que ser la del script y terminar en /exec."; return; }
  if (!clave) { $("#cmsg").textContent = "Falta la clave."; return; }
  conn = { url, clave }; lsSet(LS.conn, conn);
  $("#cmsg").textContent = "Probando la conexión…";
  await loadRemote(true);
  if ($("#setup").hidden) { $("#cmsg").textContent = "Conectado a " + ((remote && remote.planilla) || "tu planilla") + "."; toast("¡Conectado! Ya podés guardar en tu planilla"); go("cargar"); }
  else $("#cmsg").textContent = "No pude conectar. Revisá la dirección y la clave (menú Gasto Rápido → Ver clave actual).";
};
$("#sconn").onclick = () => go("ajustes");
$("#sdef").onchange = (e) => { prefs.defMedio = e.target.value; lsSet(LS.prefs, prefs); quickPreview(); };
$("#snum").onchange = (e) => { prefs.num = e.target.value; lsSet(LS.prefs, prefs); };

/* ---------- secciones ---------- */
const VIEWS = ["cargar", "mes", "tarjeta", "ajustes"];
function showView() {
  const v = VIEWS.includes(location.hash.slice(1)) ? location.hash.slice(1) : "cargar";
  VIEWS.forEach((x) => { $("#v-" + x).hidden = x !== v; });
  document.querySelectorAll(".nav a").forEach((a) => { if (a.dataset.v === v) a.setAttribute("aria-current", "page"); else a.removeAttribute("aria-current"); });
  if (v === "ajustes") { if (conn && !$("#curl").value) $("#curl").value = conn.url; if (!conn) $("#curl").focus(); }
  if ((v === "mes" || v === "tarjeta") && conn) loadResumen();
}
function go(v) { if (location.hash !== "#" + v) location.hash = v; else showView(); window.scrollTo(0, 0); }
window.addEventListener("hashchange", () => { showView(); window.scrollTo(0, 0); });

/* ---------- pestañas y eventos ---------- */
document.querySelectorAll(".tab").forEach((t) => (t.onclick = () => {
  document.querySelectorAll(".tab").forEach((x) => { x.setAttribute("aria-selected", x === t); $("#" + x.dataset.p).hidden = x !== t; });
  try { localStorage.setItem(LS.tab, t.id); } catch (e) { /* nada */ }
}));
try { const tb = localStorage.getItem(LS.tab); if (tb && $("#" + tb)) $("#" + tb).click(); } catch (e) { /* nada */ }
$("#quick").addEventListener("input", quickPreview);
$("#quick").addEventListener("keydown", (e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); quickAdd(); } });
$("#qadd").onclick = quickAdd;
$("#xmedio").onchange = renderMap;
$("#xfile").onchange = (e) => { const f = e.target.files[0]; if (f) readFile(f).catch(() => toast("No pude leer ese archivo")); e.target.value = ""; };
const drop = $("#drop");
["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) readFile(f).catch(() => toast("No pude leer ese archivo")); });
let pasteT;
$("#xpaste").addEventListener("input", (e) => { clearTimeout(pasteT); pasteT = setTimeout(() => { const t = e.target.value; if (!t.trim()) { X = { aoa: null }; renderMap(); return; } loadAoa(parseTextTable(t)); }, 250); });
$("#xadd").onclick = () => { addItems(rowsToItems(), "extracto"); X = { aoa: null }; $("#xpaste").value = ""; renderMap(); };

let tt;
function toast(msg, undo) {
  const t = $("#toast"); $("#tmsg").textContent = msg; const u = $("#tundo");
  u.hidden = !undo; u.onclick = () => { undo && undo(); t.hidden = true; };
  t.hidden = false; clearTimeout(tt); tt = setTimeout(() => (t.hidden = true), undo ? 6000 : 4000);
}

showView(); fillSelects(); fillMeses(); render(); renderRules(); renderFijos(); quickPreview(); loadRemote(true);

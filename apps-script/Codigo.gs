/**
 * Gasto Rápido — API de tu planilla de finanzas.
 *
 * La página Gasto Rápido le manda los movimientos a este script y el script
 * los escribe en la hoja «Movimientos». También le pasa a la página tus
 * categorías, cuentas, reglas y gastos fijos, y el resumen del mes.
 *
 * Instalación (una sola vez):
 *  1. En la planilla: Extensiones → Apps Script. Borrá lo que haya y pegá este archivo.
 *  2. Guardá (ícono de disquete) y recargá la planilla. Aparece el menú «Gasto Rápido».
 *  3. Menú Gasto Rápido → Generar clave. Aceptá los permisos que pide Google.
 *  4. En Apps Script: Implementar → Nueva implementación → tipo «Aplicación web».
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier usuario. → Implementar.
 *  5. Copiá la URL que termina en /exec y pegala en la página junto con la clave.
 */

const HOJA_MOV = 'Movimientos';
const COL = { fecha: 1, tipo: 2, desc: 3, monto: 4, medio: 5, cat: 6, catFinal: 7, mes: 8, cuenta: 9, id: 10, origen: 11 };
const PROP_CLAVE = 'GASTO_RAPIDO_CLAVE';
const TRANSF = 'TRANSFERENCIAS'; // movimientos entre tus cuentas: cambian saldos, no son ingreso ni gasto

/* ------------------------------------------------------------------ menú */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gasto Rápido')
    .addItem('Generar clave nueva', 'generarClave')
    .addItem('Ver clave actual', 'verClave')
    .addItem('Preparar planilla (fórmulas)', 'prepararPlanilla')
    .addToUi();
}

function generarClave() {
  prepararPlanilla_();
  const clave = Utilities.getUuid().replace(/-/g, '').slice(0, 24);
  PropertiesService.getScriptProperties().setProperty(PROP_CLAVE, clave);
  SpreadsheetApp.getUi().alert(
    'Tu clave de Gasto Rápido',
    'Copiala y pegala en la página (Ajustes → Conexión):\n\n' + clave +
      '\n\nSi generás otra, esta deja de funcionar y vas a tener que pegar la nueva en cada dispositivo.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function verClave() {
  const clave = PropertiesService.getScriptProperties().getProperty(PROP_CLAVE);
  SpreadsheetApp.getUi().alert(clave ? 'Tu clave es:\n\n' + clave : 'Todavía no generaste una clave. Usá «Generar clave nueva».');
}

/** Fórmulas automáticas de Movimientos: una sola fórmula por columna que cubre todas las filas. */
const FORMULAS = {
  G2: '=MAP(B2:B,C2:C,D2:D,F2:F,LAMBDA(tipo,desc,monto,cat,IF(monto="","",IF(tipo="Ingreso",IF(cat<>"",cat,"INGRESO"),IF(cat<>"",cat,' +
      'IFERROR(INDEX(Reglas!$B:$B,1/(1/SUMPRODUCT(MAX(ISNUMBER(SEARCH(Reglas!$A$2:$A$400,desc))*(Reglas!$A$2:$A$400<>"")*ROW(Reglas!$A$2:$A$400))))),"SIN CATEGORÍA"))))))',
  H2: '=ARRAYFORMULA(IF((A2:A="")+(D2:D=""),"",DATE(YEAR(A2:A),MONTH(A2:A),1)))',
  I2: '=ARRAYFORMULA(IF(D2:D="","",IF(E2:E="",Config!$E$5,E2:E)))'
};

function prepararPlanilla_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_MOV);
  Object.keys(FORMULAS).forEach((a1) => {
    const r = sh.getRange(a1);
    if (!r.getFormula()) {
      sh.getRange(a1 + ':' + a1.charAt(0)).clearContent(); // la fórmula necesita la columna vacía (ej. G2:G)
      r.setFormula(FORMULAS[a1]);
    }
  });
  sh.getRange('H2:H').setNumberFormat('mmm yyyy');
}

function prepararPlanilla() {
  prepararPlanilla_();
  SpreadsheetApp.getUi().alert('Listo: las columnas G, H e I de Movimientos se calculan solas.');
}

/* ------------------------------------------------------------------ web */

function doGet() {
  return json_({ ok: true, app: 'gasto-rapido', mensaje: 'El script funciona. Pegá esta dirección en la página Gasto Rápido.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'El pedido no es válido.' });
  }
  const clave = PropertiesService.getScriptProperties().getProperty(PROP_CLAVE);
  if (!clave) return json_({ ok: false, error: 'Falta generar la clave: en la planilla, menú Gasto Rápido → Generar clave nueva.' });
  if (String(req.clave || '') !== clave) return json_({ ok: false, error: 'La clave no coincide. Revisala en Ajustes.' });

  try {
    switch (req.accion) {
      case 'config': return json_(Object.assign({ ok: true }, leerConfig_()));
      case 'agregar': return json_(Object.assign({ ok: true }, agregar_(req.items || [])));
      case 'regla': return json_(Object.assign({ ok: true }, agregarRegla_(req.palabra, req.categoria)));
      case 'resumen': return json_(Object.assign({ ok: true }, resumen_(req.mes)));
      case 'cuota': return json_(Object.assign({ ok: true }, agregarCuota_(req)));
      case 'saldo': return json_(Object.assign({ ok: true }, ajustarSaldo_(req)));
      case 'fijo': return json_(Object.assign({ ok: true }, agregarFijo_(req)));
      default: return json_({ ok: false, error: 'Acción desconocida: ' + req.accion });
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ lectura */

function leerConfig_() {
  const ss = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName('Config');
  const lista = (col) => cfg.getRange(5, col, 40, 1).getValues()
    .map((r) => String(r[0]).trim()).filter((v) => v);
  const categorias = lista(1).filter((c) => c !== 'SIN CATEGORÍA');
  const cuentas = lista(3);
  const cuentaDefecto = String(cfg.getRange('E5').getValue() || cuentas[0] || '');

  const rg = ss.getSheetByName('Reglas');
  const nr = Math.max(rg.getLastRow() - 1, 0);
  const reglas = nr ? rg.getRange(2, 1, nr, 2).getValues()
    .filter((r) => String(r[0]).trim() && String(r[1]).trim())
    .map((r) => [String(r[0]).trim(), String(r[1]).trim()]) : [];

  const fj = ss.getSheetByName('Fijos');
  // Monto vacío = servicio de monto variable (luz, agua, internet…): se escribe cada mes.
  const fijos = fj.getRange(5, 1, 15, 6).getValues()
    .filter((r) => String(r[0]).trim())
    .map((r) => ({
      concepto: String(r[0]).trim(),
      cat: String(r[1] || 'FIJOS').trim(),
      monto: Number(r[2]) > 0 ? Number(r[2]) : 0,
      variable: !(Number(r[2]) > 0),
      medio: String(r[3] || cuentaDefecto).trim(),
      dia: Number(r[4]) > 0 ? Math.min(31, Math.round(Number(r[4]))) : 0,
      activo: String(r[5]).trim() === '' ? true : /^s/i.test(String(r[5]))
    }));

  return { categorias, cuentas, cuentaDefecto, reglas, fijos, planilla: ss.getName(), url: ss.getUrl() };
}

/** Primera fila vacía de la columna A (ignora las columnas calculadas). */
function primeraFilaLibre_(sh) {
  const n = sh.getMaxRows() - 1;
  const a = sh.getRange(2, COL.fecha, n, 1).getValues();
  for (let i = a.length - 1; i >= 0; i--) {
    if (a[i][0] !== '' && a[i][0] !== null) return i + 3;
  }
  return 2;
}

/* ------------------------------------------------------------------ escritura */

function agregar_(items) {
  if (!Array.isArray(items) || !items.length) return { agregados: 0, repetidos: 0, ids: [] };
  if (items.length > 500) throw new Error('Mandá como máximo 500 movimientos por vez.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(HOJA_MOV);
    const { categorias, cuentas } = leerConfig_();
    const fila = primeraFilaLibre_(sh);

    const existentes = new Set(
      fila > 2 ? sh.getRange(2, COL.id, fila - 2, 1).getValues().map((r) => String(r[0])).filter((v) => v) : []);

    const nuevos = [];
    const repetidos = [];
    items.forEach((it) => {
      const id = String((it && it.id) || '').slice(0, 40);
      if (!id || existentes.has(id)) { repetidos.push(id); return; }
      const fecha = parseFecha_(it.fecha);
      const monto = Math.abs(Number(it.monto));
      if (!fecha || !isFinite(monto) || monto <= 0) throw new Error('Movimiento inválido: ' + JSON.stringify(it).slice(0, 120));
      const tipo = it.tipo === 'Ingreso' ? 'Ingreso' : 'Gasto';
      const medio = cuentas.indexOf(String(it.medio)) >= 0 ? String(it.medio) : '';
      // Los ingresos solo llevan categoría si es TRANSFERENCIAS (plata que viene de otra cuenta tuya).
      const cat = categorias.indexOf(String(it.cat)) >= 0 && (tipo === 'Gasto' || String(it.cat) === TRANSF) ? String(it.cat) : '';
      const desc = String(it.desc || '').replace(/^[=+\-@]+/, '').slice(0, 120);
      existentes.add(id);
      nuevos.push({ fila: [fecha, tipo, desc, Math.round(monto * 100) / 100, medio, cat], extra: [id, String(it.origen || 'página').slice(0, 20)], id });
    });

    if (nuevos.length) {
      const falta = fila + nuevos.length - 1 - sh.getMaxRows();
      if (falta > 0) sh.insertRowsAfter(sh.getMaxRows(), falta + 200);
      sh.getRange(fila, COL.fecha, nuevos.length, 6).setValues(nuevos.map((n) => n.fila));
      sh.getRange(fila, COL.id, nuevos.length, 2).setValues(nuevos.map((n) => n.extra));
      sh.getRange(fila, COL.fecha, nuevos.length, 1).setNumberFormat('dd/mm/yyyy');
      sh.getRange(fila, COL.monto, nuevos.length, 1).setNumberFormat('$ #,##0.00');
    }
    SpreadsheetApp.flush();
    return { agregados: nuevos.length, repetidos: repetidos.length, ids: nuevos.map((n) => n.id).concat(repetidos), desdeFila: fila };
  } finally {
    lock.releaseLock();
  }
}

function agregarRegla_(palabra, categoria) {
  palabra = String(palabra || '').trim().toLowerCase().slice(0, 40);
  categoria = String(categoria || '').trim();
  const { categorias } = leerConfig_();
  if (!palabra || categorias.indexOf(categoria) < 0) throw new Error('Regla inválida.');
  const rg = SpreadsheetApp.getActive().getSheetByName('Reglas');
  const n = Math.max(rg.getLastRow() - 1, 0);
  const vals = n ? rg.getRange(2, 1, n, 2).getValues() : [];
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === palabra) {
      rg.getRange(i + 2, 2).setValue(categoria);
      return { regla: [palabra, categoria], actualizada: true };
    }
  }
  let fila = 2;
  for (let i = vals.length - 1; i >= 0; i--) { if (String(vals[i][0]).trim()) { fila = i + 3; break; } }
  rg.getRange(fila, 1, 1, 2).setValues([[palabra, categoria]]);
  return { regla: [palabra, categoria], actualizada: false };
}

/* ------------------------------------------------------------------ tarjeta y cuotas */

const CUOTAS = { hoja: 'Cuotas', fila: 5, filas: 40 }; // A Compra · B Dónde · C Precio por cuota · D Cant. cuotas · E Primer mes

function agregarCuota_(req) {
  const compra = String(req.compra || '').replace(/^[=+\-@]+/, '').trim().slice(0, 60);
  const donde = req.donde === 'Aparte' ? 'Aparte' : 'Tarjeta';
  const precio = Math.round(Math.abs(Number(req.precio)) * 100) / 100;
  const cuotas = Math.round(Number(req.cuotas));
  const m = /^(\d{4})-(\d{2})$/.exec(String(req.primerMes || ''));
  if (!compra || !(precio > 0) || !(cuotas >= 1 && cuotas <= 72) || !m) throw new Error('Compra en cuotas inválida.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CUOTAS.hoja);
    const a = sh.getRange(CUOTAS.fila, 1, CUOTAS.filas, 1).getValues();
    let libre = -1;
    for (let i = 0; i < a.length; i++) { if (String(a[i][0]).trim() === '') { libre = i; break; } }
    if (libre < 0) throw new Error('La hoja Cuotas está llena: borrá compras viejas ya pagadas para hacer lugar.');
    const fila = CUOTAS.fila + libre;
    sh.getRange(fila, 1, 1, 5).setValues([[compra, donde, precio, cuotas, new Date(Number(m[1]), Number(m[2]) - 1, 1, 12)]]);
    sh.getRange(fila, 5).setNumberFormat('mmm yyyy');
    SpreadsheetApp.flush();
    return { fila, compra, donde, precio, cuotas, primerMes: m[0] };
  } finally {
    lock.releaseLock();
  }
}

/** Qué vence en el mes, el siguiente, deuda pendiente y compras activas. */
function tarjeta_(y, mo) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CUOTAS.hoja);
  const vals = sh.getRange(CUOTAS.fila, 1, CUOTAS.filas, 5).getValues();
  const idx = (d) => d.getFullYear() * 12 + d.getMonth();
  const hoy = y * 12 + mo;
  let esteMes = 0, proximo = 0, deuda = 0;
  const porDonde = { Tarjeta: 0, Aparte: 0 };
  const compras = [];
  vals.forEach((r) => {
    const compra = String(r[0]).trim(), donde = String(r[1] || 'Tarjeta'), precio = Number(r[2]) || 0, n = Number(r[3]) || 0, f = r[4];
    if (!compra || !precio || !n || !(f instanceof Date)) return;
    const ini = idx(f), fin = ini + n - 1;
    if (hoy >= ini && hoy <= fin) { esteMes += precio; porDonde[donde === 'Aparte' ? 'Aparte' : 'Tarjeta'] += precio; }
    if (hoy + 1 >= ini && hoy + 1 <= fin) proximo += precio;
    const restantes = Math.max(0, fin - Math.max(hoy, ini - 1));
    deuda += restantes * precio;
    if (fin >= hoy) {
      compras.push({ compra, donde, precio, cuotas: n, actual: Math.min(n, Math.max(0, hoy - ini + 1)), primerMes: fmtMes_(Math.floor(ini / 12), ini % 12) });
    }
  });
  return { esteMes: round2_(esteMes), proximo: round2_(proximo), deuda: round2_(deuda), porDonde, compras };
}

/* ------------------------------------------------------------------ saldos */

// Hoja Saldos: fila 4 = nombre de cada cuenta (B, D, F…), filas desde la 6 = un mes por fila.
// Por cuenta: columna «Inicio real (opcional)» y al lado «Saldo final» (fórmula).
const SALDOS = { hoja: 'Saldos', filaCuentas: 4, filaDatos: 6 };

function mapaSaldos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SALDOS.hoja);
  const ncol = sh.getLastColumn();
  const cab = sh.getRange(SALDOS.filaCuentas, 1, 1, ncol).getValues()[0];
  const cuentas = [];
  for (let c = 2; c <= ncol; c += 2) {
    const nombre = String(cab[c - 1] || '').trim();
    if (nombre && nombre.toUpperCase() !== 'TOTAL') cuentas.push({ nombre, colReal: c, colFinal: c + 1 });
  }
  const n = sh.getLastRow() - SALDOS.filaDatos + 1;
  const meses = sh.getRange(SALDOS.filaDatos, 1, Math.max(n, 1), 1).getValues().map((r) => r[0]);
  return { sh, cuentas, meses };
}

function filaMes_(meses, y, mo) {
  for (let i = 0; i < meses.length; i++) {
    const d = meses[i];
    if (d instanceof Date && d.getFullYear() === y && d.getMonth() === mo) return SALDOS.filaDatos + i;
  }
  return -1;
}

/** Saldo de cada cuenta al día de hoy (saldo final del mes en curso, con lo cargado hasta ahora). */
function saldos_(y, mo) {
  const { sh, cuentas, meses } = mapaSaldos_();
  const fila = filaMes_(meses, y, mo);
  if (fila < 0) return { cuentas: [], total: 0, aviso: 'La hoja Saldos no tiene una fila para este mes.' };
  const fil = sh.getRange(fila, 1, 1, sh.getLastColumn()).getValues()[0];
  const lista = cuentas.map((c) => ({
    cuenta: c.nombre,
    saldo: round2_(Number(fil[c.colFinal - 1]) || 0),
    ajustado: fil[c.colReal - 1] !== '' && fil[c.colReal - 1] !== null
  }));
  return { cuentas: lista, total: round2_(lista.reduce((a, c) => a + c.saldo, 0)) };
}

/** Corrige el saldo de una cuenta con lo que dice el banco hoy. */
function ajustarSaldo_(req) {
  const cuenta = String(req.cuenta || '').trim();
  const real = Number(req.monto);
  if (!cuenta || !isFinite(real)) throw new Error('Saldo inválido.');
  const hoy = new Date();
  const y = hoy.getFullYear(), mo = hoy.getMonth();
  const { sh, cuentas, meses } = mapaSaldos_();
  const c = cuentas.find((x) => x.nombre === cuenta);
  if (!c) throw new Error('La cuenta «' + cuenta + '» no está en la hoja Saldos.');
  const fila = filaMes_(meses, y, mo);
  if (fila < 0) throw new Error('La hoja Saldos no tiene una fila para este mes.');

  // Movimiento neto de la cuenta en el mes, para que el saldo final quede igual a lo que dice el banco.
  const mov = SpreadsheetApp.getActive().getSheetByName(HOJA_MOV);
  const ult = primeraFilaLibre_(mov) - 1;
  let neto = 0;
  if (ult >= 2) {
    mov.getRange(2, 1, ult - 1, COL.cuenta).getValues().forEach((r) => {
      const f = r[COL.fecha - 1];
      if (!(f instanceof Date) || f.getFullYear() !== y || f.getMonth() !== mo) return;
      if (String(r[COL.cuenta - 1] || r[COL.medio - 1]) !== cuenta) return;
      const m = Number(r[COL.monto - 1]) || 0;
      neto += r[COL.tipo - 1] === 'Ingreso' ? m : -m;
    });
  }
  const inicio = round2_(real - neto);
  sh.getRange(fila, c.colReal).setValue(inicio);
  SpreadsheetApp.flush();
  return { cuenta, saldo: round2_(real), inicioDelMes: inicio };
}

/* ------------------------------------------------------------------ resumen */

function resumen_(mes) {
  const hoy = new Date();
  const m = /^(\d{4})-(\d{2})$/.exec(String(mes || ''));
  const y = m ? Number(m[1]) : hoy.getFullYear();
  const mo = m ? Number(m[2]) - 1 : hoy.getMonth();

  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_MOV);
  const ult = primeraFilaLibre_(sh) - 1;
  if (ult < 2) return { mes: fmtMes_(y, mo), ingresos: 0, gastos: 0, categorias: [], ultimos: [], tarjeta: tarjeta_(y, mo), saldos: saldos_(y, mo) };
  const datos = sh.getRange(2, 1, ult - 1, COL.origen).getValues();

  let ingresos = 0, gastos = 0;
  const porCat = {};
  const descMes = {};
  const ultimoPorDesc = {};
  const porCuenta = {};
  datos.forEach((r) => {
    const f = r[COL.fecha - 1];
    if (f instanceof Date && r[COL.tipo - 1] !== 'Ingreso') {
      const k = String(r[COL.desc - 1] || '').trim().toLowerCase();
      if (k) ultimoPorDesc[k] = { monto: Number(r[COL.monto - 1]) || 0, fecha: Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd') };
    }
    if (!(f instanceof Date) || f.getFullYear() !== y || f.getMonth() !== mo) return;
    const monto = Number(r[COL.monto - 1]) || 0;
    const cuenta = String(r[COL.cuenta - 1] || r[COL.medio - 1] || '');
    const esTransf = String(r[COL.catFinal - 1] || r[COL.cat - 1]) === TRANSF;
    if (esTransf) {
      porCuenta[cuenta] = (porCuenta[cuenta] || 0) + (r[COL.tipo - 1] === 'Ingreso' ? monto : -monto);
    } else if (r[COL.tipo - 1] === 'Ingreso') {
      ingresos += monto;
      porCuenta[cuenta] = (porCuenta[cuenta] || 0) + monto;
    } else {
      gastos += monto;
      const dk = String(r[COL.desc - 1] || '').trim().toLowerCase();
      if (dk && !descMes[dk]) descMes[dk] = Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const c = String(r[COL.catFinal - 1] || r[COL.cat - 1] || 'SIN CATEGORÍA');
      porCat[c] = (porCat[c] || 0) + monto;
      porCuenta[cuenta] = (porCuenta[cuenta] || 0) - monto;
    }
  });
  const categorias = Object.keys(porCat).map((k) => ({ cat: k, monto: round2_(porCat[k]) }))
    .sort((a, b) => b.monto - a.monto);
  const ultimos = datos.slice(-8).reverse().map((r) => ({
    fecha: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    tipo: r[1], desc: r[2], monto: Number(r[3]) || 0, medio: r[4], cat: r[COL.catFinal - 1]
  }));
  return { mes: fmtMes_(y, mo), ingresos: round2_(ingresos), gastos: round2_(gastos), categorias, porCuenta, ultimos, tarjeta: tarjeta_(y, mo), saldos: saldos_(y, mo), fijosCargados: fijosCargados_(descMes), fijosUltimo: fijosUltimo_(ultimoPorDesc) };
}

/** Agrega un gasto fijo nuevo a la hoja Fijos (filas 5 a 19). */
function agregarFijo_(req) {
  const concepto = String(req.concepto || '').replace(/^[=+\-@]+/, '').trim().slice(0, 40);
  const variable = !!req.variable;
  const monto = variable ? '' : Math.round(Math.abs(Number(req.monto)) * 100) / 100;
  const dia = Number(req.dia) > 0 ? Math.min(31, Math.round(Number(req.dia))) : '';
  const { categorias, cuentas, cuentaDefecto } = leerConfig_();
  const cat = categorias.indexOf(String(req.cat)) >= 0 ? String(req.cat) : 'FIJOS';
  const medio = cuentas.indexOf(String(req.medio)) >= 0 ? String(req.medio) : cuentaDefecto;
  if (!concepto || (!variable && !(monto > 0))) throw new Error('Gasto fijo inválido.');
  const fj = SpreadsheetApp.getActive().getSheetByName('Fijos');
  const a = fj.getRange(5, 1, 15, 1).getValues();
  for (let i = 0; i < a.length; i++) {
    if (String(a[i][0]).trim().toLowerCase() === concepto.toLowerCase()) {
      fj.getRange(5 + i, 2, 1, 5).setValues([[cat, monto, medio, dia, 'Sí']]);
      return { concepto, actualizado: true };
    }
  }
  const libre = a.findIndex((r) => String(r[0]).trim() === '');
  if (libre < 0) throw new Error('La hoja Fijos está llena (15 filas). Borrá alguno que ya no uses.');
  fj.getRange(5 + libre, 1, 1, 6).setValues([[concepto, cat, monto, medio, dia, 'Sí']]);
  return { concepto, actualizado: false };
}

/** Último monto pagado de cada fijo (para mostrar de referencia en los de monto variable). */
function fijosUltimo_(ultimoPorDesc) {
  const fj = SpreadsheetApp.getActive().getSheetByName('Fijos');
  const out = {};
  fj.getRange(5, 1, 15, 1).getValues().forEach((r) => {
    const k = String(r[0] || '').trim().toLowerCase();
    if (k && ultimoPorDesc[k]) out[k] = ultimoPorDesc[k];
  });
  return out;
}

/** Qué fijos de la hoja Fijos ya tienen un gasto con la misma descripción en el mes. */
function fijosCargados_(descMes) {
  const fj = SpreadsheetApp.getActive().getSheetByName('Fijos');
  const out = {};
  fj.getRange(5, 1, 15, 1).getValues().forEach((r) => {
    const k = String(r[0] || '').trim().toLowerCase();
    if (k && descMes[k]) out[k] = descMes[k];
  });
  return out;
}

/* ------------------------------------------------------------------ util */

function parseFecha_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
  return isNaN(d.getTime()) ? null : d;
}
function fmtMes_(y, mo) { return y + '-' + ('0' + (mo + 1)).slice(-2); }
function round2_(n) { return Math.round(n * 100) / 100; }

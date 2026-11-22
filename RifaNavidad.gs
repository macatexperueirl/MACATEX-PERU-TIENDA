/**
 * MACATEXPERU - Rifa Navideña 2026
 * Agregar este archivo al proyecto de Apps Script que ya sirve el catálogo.
 * Luego añadir en doGet(e), después de obtener "tipo":
 * if (tipo === 'rifa') return salida(obtenerRifaPublica_(), e);
 *
 * La reserva se registra mediante POST desde un formulario HTML oculto (iframe).
 * El listado público solo devuelve número y estado; nunca datos personales.
 */
const RIFA_CFG = {
  HOJA: 'RIFA_2026',
  TOTAL: 300,
  PRECIO: 10,
  MINUTOS_RESERVA: 30,
  YAPE: '930523654'
};

function prepararRifaNavidena2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RIFA_CFG.HOJA);
  if (!sh) sh = ss.insertSheet(RIFA_CFG.HOJA);
  const headers = ['Número', 'Estado', 'Nombre', 'WhatsApp', 'Fecha reserva', 'Reserva hasta', 'Fecha pago', 'Observaciones'];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    const rows = Array.from({length: RIFA_CFG.TOTAL}, (_, i) => [String(i + 1).padStart(3, '0'), 'DISPONIBLE', '', '', '', '', '', '']);
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sh.autoResizeColumns(1, headers.length);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getDisplayValues()[0];
    if (headers.some((h, i) => current[i] !== h)) {
      throw new Error('La hoja RIFA_2026 ya existe con encabezados diferentes. No se modificó para proteger los datos.');
    }
  }
  SpreadsheetApp.getUi().alert('RIFA_2026 está lista. Revisa los 300 números del 001 al 300.');
}

/** Endpoint público JSONP: devuelve únicamente número y estado. */
function obtenerRifaPublica_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RIFA_CFG.HOJA);
    if (!sh) throw new Error('Primero ejecuta prepararRifaNavidena2026().');
    liberarReservasVencidas_(sh);
    const values = sh.getDataRange().getValues();
    return values.slice(1).filter(r => r[0]).map(r => ({
      numero: String(r[0]).padStart(3, '0'),
      estado: String(r[1] || 'DISPONIBLE').toUpperCase()
    }));
  } finally {
    lock.releaseLock();
  }
}

/** Devuelve el resultado del POST al sitio principal dentro del iframe. */
function respuestaReservaRifa_(datos) {
  const payload = Object.assign({ type: 'MACATEX_RIFA_RESULT' }, datos);
  const mensaje = JSON.stringify(payload).replace(/</g, '\\u003c');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body><script>' +
    'window.parent.postMessage(' + mensaje + ', "*");' +
    'if (window.top !== window.parent) window.top.postMessage(' + mensaje + ', "*");' +
    '</scr' + 'ipt></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Recibe la reserva y comunica explícitamente si fue aceptada o rechazada. */
function doPost(e) {
  const p = (e && e.parameter) || {};
  const requestId = String(p.requestId || '');
  try {
    if (String(p.accion || '') !== 'reservar_rifa') {
      throw new Error('Solicitud no válida.');
    }
    const numero = String(p.numero || '').trim();
    const nombre = String(p.nombre || '').trim().slice(0, 100);
    const whatsapp = String(p.whatsapp || '').replace(/[^0-9+]/g, '').slice(0, 20);
    if (!/^\d{3}$/.test(numero) || Number(numero) < 1 || Number(numero) > RIFA_CFG.TOTAL) {
      throw new Error('Número inválido.');
    }
    if (nombre.length < 2 || !/^\+?\d{9,15}$/.test(whatsapp)) {
      throw new Error('Completa correctamente tu nombre y WhatsApp.');
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    let resultado;
    try {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RIFA_CFG.HOJA);
      if (!sh) throw new Error('La hoja de la rifa aún no está preparada.');
      liberarReservasVencidas_(sh);
      const last = sh.getLastRow();
      const data = sh.getRange(2, 1, Math.max(last - 1, 1), 8).getValues();
      const idx = data.findIndex(r => String(r[0]).padStart(3, '0') === numero);
      if (idx < 0) throw new Error('No encontramos ese número.');
      if (String(data[idx][1]).toUpperCase() !== 'DISPONIBLE') {
        throw new Error('Ese número ya no está disponible. Elige otro.');
      }
      const now = new Date();
      const expires = new Date(now.getTime() + RIFA_CFG.MINUTOS_RESERVA * 60000);
      const row = idx + 2;
      sh.getRange(row, 2, 1, 7).setValues([[
        'RESERVADO', nombre, whatsapp, now, expires, '', ''
      ]]);
      sh.getRange(row, 5, 1, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
      resultado = {
        ok: true,
        requestId: requestId,
        numero: numero,
        expiresAt: expires.toISOString(),
        message: 'Reserva aceptada por 30 minutos.'
      };
    } finally {
      lock.releaseLock();
    }
    return respuestaReservaRifa_(resultado);
  } catch (err) {
    return respuestaReservaRifa_({
      ok: false,
      requestId: requestId,
      message: err && err.message ? err.message : 'No se pudo registrar la reserva.'
    });
  }
}

/** Libera automáticamente reservas vencidas sin tocar números vendidos. */
function liberarReservasVencidas_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const range = sh.getRange(2, 1, last - 1, 8);
  const rows = range.getValues();
  const now = new Date();
  let changed = false;
  rows.forEach(r => {
    if (String(r[1]).toUpperCase() === 'RESERVADO' && r[5] instanceof Date && r[5].getTime() <= now.getTime()) {
      r[1] = 'DISPONIBLE';
      r[2] = ''; r[3] = ''; r[4] = ''; r[5] = ''; r[6] = '';
      r[7] = 'Reserva vencida automáticamente';
      changed = true;
    }
  });
  if (changed) range.setValues(rows);
}

/** Ejecutar manualmente desde Apps Script para liberar vencidos y actualizar el listado. */
function actualizarRifaNavidena2026() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RIFA_CFG.HOJA);
  if (!sh) throw new Error('No existe la hoja RIFA_2026.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { liberarReservasVencidas_(sh); }
  finally { lock.releaseLock(); }
  SpreadsheetApp.getUi().alert('Se revisaron las reservas vencidas.');
}


/** Confirma manualmente el pago y marca un número como VENDIDO. */
function confirmarNumeroRifa() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('MACATEX - Confirmar venta de rifa', 'Escribe el número de tres dígitos (001 al 300) cuyo Yape ya verificaste:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const numero = r.getResponseText().trim();
  if (!/^\d{3}$/.test(numero) || Number(numero) < 1 || Number(numero) > RIFA_CFG.TOTAL) {
    ui.alert('Número inválido. Usa el formato 001 al 300.');
    return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RIFA_CFG.HOJA);
    if (!sh) throw new Error('No existe la hoja RIFA_2026.');
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
    const idx = values.findIndex(row => String(row[0]).padStart(3, '0') === numero);
    if (idx < 0) throw new Error('No se encontró el número.');
    if (String(values[idx][1]).toUpperCase() !== 'RESERVADO') throw new Error('El número no está reservado. Estado actual: ' + values[idx][1]);
    const row = idx + 2;
    sh.getRange(row, 2).setValue('VENDIDO');
    sh.getRange(row, 7).setValue(new Date());
    sh.getRange(row, 8).setValue('Pago Yape verificado manualmente');
    ui.alert('Número ' + numero + ' marcado como VENDIDO.');
  } finally {
    lock.releaseLock();
  }
}

/** Liberación manual excepcional de un número reservado. */
function liberarNumeroRifa() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('MACATEX - Liberar número', 'Escribe el número reservado de tres dígitos (001 al 300):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const numero = r.getResponseText().trim();
  if (!/^\d{3}$/.test(numero) || Number(numero) < 1 || Number(numero) > RIFA_CFG.TOTAL) {
    ui.alert('Número inválido. Usa el formato 001 al 300.');
    return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RIFA_CFG.HOJA);
    if (!sh) throw new Error('No existe la hoja RIFA_2026.');
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
    const idx = values.findIndex(row => String(row[0]).padStart(3, '0') === numero);
    if (idx < 0) throw new Error('No se encontró el número.');
    if (String(values[idx][1]).toUpperCase() !== 'RESERVADO') throw new Error('Solo se puede liberar un número RESERVADO.');
    sh.getRange(idx + 2, 2, 1, 7).setValues([['DISPONIBLE', '', '', '', '', '', 'Liberado manualmente']]);
    ui.alert('Número ' + numero + ' liberado.');
  } finally {
    lock.releaseLock();
  }
}

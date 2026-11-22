/**
 * MACATEXPERU - Tickets digitales verificables de la Rifa Navideña 2026
 * Copiar este archivo al mismo proyecto de Apps Script que contiene RifaNavidad.gs.
 *
 * Integración requerida en doGet(e), antes de la salida general:
 * if (tipo === 'verificar_rifa') return verificarTicketRifa_(e);
 *
 * Agregar al menú onOpen():
 * .addItem('Generar ticket PDF de rifa', 'generarTicketRifaManual')
 *
 * Flujo: primero confirmar Yape con confirmarNumeroRifa(); luego generar el PDF
 * desde el menú. El archivo queda en una carpeta privada de Drive para descargar
 * y enviarlo manualmente al comprador por WhatsApp.
 */

const TICKET_RIFA_CFG = {
  HOJA: 'RIFA_2026',
  FOLDER_PROPERTY: 'MACATEX_RIFA_TICKETS_FOLDER_ID',
  HEADERS: ['Código verificación', 'Archivo PDF Drive']
};

function asegurarColumnasTicketRifa_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TICKET_RIFA_CFG.HOJA);
  if (!sh) throw new Error('No existe la hoja RIFA_2026.');
  sh.getRange(1, 9, 1, 2).setValues([TICKET_RIFA_CFG.HEADERS]);
  sh.getRange(1, 9, 1, 2).setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
  return sh;
}

function generarTicketRifaManual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('MACATEXPERU - Generar ticket', 'Ingresa el número VENDIDO (001 al 300):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const numero = resp.getResponseText().trim();
  if (!/^\d{3}$/.test(numero) || Number(numero) < 1 || Number(numero) > 300) {
    ui.alert('Número inválido. Usa el formato 001 al 300.');
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = asegurarColumnasTicketRifa_();
    const last = sh.getLastRow();
    const rows = sh.getRange(2, 1, last - 1, 10).getValues();
    const idx = rows.findIndex(r => String(r[0]).padStart(3, '0') === numero);
    if (idx < 0) throw new Error('No se encontró el número ' + numero + '.');
    const row = idx + 2;
    const data = rows[idx];
    if (String(data[1]).toUpperCase() !== 'VENDIDO') {
      throw new Error('El número debe estar VENDIDO. Primero confirma el pago Yape.');
    }
    if (!String(data[2] || '').trim()) throw new Error('El registro no tiene nombre del comprador.');
    const codigo = String(data[8] || Utilities.getUuid().replace(/-/g, '').toUpperCase());
    const existingId = String(data[9] || '');
    if (existingId) {
      try {
        const existing = DriveApp.getFileById(existingId);
        ui.alert('El ticket ya fue generado.\n\nCódigo: ' + codigo + '\nArchivo: ' + existing.getUrl());
        return;
      } catch (e) {
        // Si el archivo fue eliminado, se genera nuevamente con el mismo código.
      }
    }

    const verifyUrl = ScriptApp.getService().getUrl() + '?tipo=verificar_rifa&codigo=' + encodeURIComponent(codigo);
    const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(verifyUrl) + '&size=260&margin=2&format=png';
    const qrResponse = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
    if (qrResponse.getResponseCode() !== 200) throw new Error('No se pudo generar el QR. Intenta nuevamente.');
    const qrBase64 = Utilities.base64Encode(qrResponse.getBlob().getBytes());
    const nombre = String(data[2]).trim();
    const fechaPago = data[6] instanceof Date
      ? Utilities.formatDate(data[6], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
      : String(data[6] || 'Registrado');
    const html = crearHtmlTicketRifa_(numero, nombre, fechaPago, codigo, qrBase64);
    const pdf = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF)
      .setName('MACATEX_RIFA_2026_' + numero + '.pdf');

    const props = PropertiesService.getScriptProperties();
    let folderId = props.getProperty(TICKET_RIFA_CFG.FOLDER_PROPERTY);
    let folder;
    try { folder = folderId ? DriveApp.getFolderById(folderId) : null; } catch (e) { folder = null; }
    if (!folder) {
      folder = DriveApp.createFolder('MACATEX - Tickets Rifa Navideña 2026');
      props.setProperty(TICKET_RIFA_CFG.FOLDER_PROPERTY, folder.getId());
    }
    const file = folder.createFile(pdf);
    // La carpeta y el PDF permanecen privados. El operador descarga el PDF y lo adjunta en WhatsApp.
    sh.getRange(row, 9, 1, 2).setValues([[codigo, file.getId()]]);
    ui.alert('Ticket generado correctamente.\n\nNúmero: ' + numero +
      '\nComprador: ' + nombre + '\nCódigo: ' + codigo +
      '\n\nAbre el archivo para descargarlo y enviarlo manualmente por WhatsApp:\n' + file.getUrl());
  } finally {
    lock.releaseLock();
  }
}

function crearHtmlTicketRifa_(numero, nombre, fechaPago, codigo, qrBase64) {
  const safe = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  return `<!doctype html><html><head><meta charset="utf-8"><style>
    @page{size:A5 portrait;margin:12mm}*{box-sizing:border-box}
    body{font-family:Arial,sans-serif;color:#242424;margin:0;background:#fffaf1}
    .ticket{border:4px solid #b5162c;border-radius:18px;padding:22px;background:#fffdf8;position:relative}
    .top{text-align:center;border-bottom:3px solid #c18a22;padding-bottom:14px}
    .brand{font-size:28px;font-weight:900;color:#a40f24;letter-spacing:1px}
    .sub{font-size:12px;font-weight:bold;color:#075a3e;margin-top:4px}
    .ribbon{background:#b5162c;color:white;text-align:center;font-size:23px;font-weight:900;padding:12px;margin:18px -22px}
    .green{background:#075a3e;color:white;border-radius:20px;text-align:center;padding:8px;font-weight:bold}
    .number{border:3px solid #b5162c;border-radius:14px;text-align:center;padding:12px;margin:18px auto;max-width:260px}
    .number small{display:block;font-weight:bold;color:#555;font-size:12px}
    .number strong{font-size:54px;color:#b5162c;letter-spacing:3px}
    .details{background:#fff4df;border:1px solid #e7c98c;border-radius:12px;padding:14px;line-height:1.8;font-size:14px}
    .verify{display:flex;align-items:center;gap:14px;margin-top:18px;padding:12px;border:1px dashed #b78012;border-radius:12px}
    .verify img{width:112px;height:112px}
    .verify strong{color:#075a3e}
    .code{font-size:10px;color:#555;word-break:break-all;margin-top:6px}
    .footer{text-align:center;color:#666;font-size:10px;margin-top:16px;line-height:1.5}
    .price{text-align:center;color:#b5162c;font-size:24px;font-weight:900;margin:14px}
  </style></head><body><div class="ticket">
    <div class="top"><div class="brand">🎁 MACATEXPERU</div><div class="sub">Canastas Navideñas y Regalos Corporativos</div></div>
    <div class="ribbon">GRAN RIFA NAVIDEÑA 2026</div>
    <div class="green">¡3 PREMIOS, 3 GANADORES!</div>
    <div class="number"><small>NÚMERO DE PARTICIPACIÓN</small><strong>${safe(numero)}</strong></div>
    <div class="details">
      <b>Participante:</b> ${safe(nombre)}<br>
      <b>Fecha de pago:</b> ${safe(fechaPago)}<br>
      <b>Estado:</b> PAGO VERIFICADO<br>
      <b>Premios:</b> Súper Canasta Navideña CAN-015 · Juego de ollas Record · Mini licuadora portátil Oster
    </div>
    <div class="price">S/ 10.00 · PAGADO</div>
    <div class="green">Sorteo: domingo 27 de diciembre de 2026 · 8:00 p. m.</div>
    <div class="verify"><img src="data:image/png;base64,${qrBase64}"><div><strong>VERIFICA TU TICKET</strong><br><span style="font-size:12px">Escanea el QR para consultar el registro oficial de MACATEXPERU.</span><div class="code">Código: ${safe(codigo)}</div></div></div>
    <div class="footer">Transmisión por Facebook Live: MACATEXPERU · Recojo presencial de premios en Santa Anita.<br>Conserva este ticket hasta el sorteo. Este documento acredita el registro de participación; no es comprobante tributario.</div>
  </div></body></html>`;
}

function verificarTicketRifa_(e) {
  const codigo = String((e && e.parameter && e.parameter.codigo) || '').trim();
  const escape = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  let html;
  if (!/^[A-F0-9]{32}$/.test(codigo)) {
    html = paginaVerificacionRifa_('TICKET NO VÁLIDO', 'El código no tiene un formato reconocido.', false);
  } else {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TICKET_RIFA_CFG.HOJA);
    if (!sh || sh.getLastRow() < 2) {
      html = paginaVerificacionRifa_('NO DISPONIBLE', 'No se pudo consultar el registro en este momento.', false);
    } else {
      const values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getDisplayValues();
      const row = values.find(r => String(r[8] || '') === codigo);
      if (!row || String(row[1]).toUpperCase() !== 'VENDIDO') {
        html = paginaVerificacionRifa_('TICKET NO REGISTRADO', 'No encontramos un ticket pagado con este código.', false);
      } else {
        const fullName = String(row[2] || '').trim();
        const parts = fullName.split(/\s+/);
        const masked = parts.length > 1 ? parts[0] + ' ' + parts[1].charAt(0) + '.' : parts[0];
        html = paginaVerificacionRifa_('TICKET VÁLIDO', 'Número ' + escape(String(row[0]).padStart(3, '0')) +
          ' · Participante: ' + escape(masked) + ' · Pago registrado: ' + escape(row[6] || 'Sí'), true);
      }
    }
  }
  return HtmlService.createHtmlOutput(html).setTitle('Verificación de ticket MACATEXPERU')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function paginaVerificacionRifa_(titulo, detalle, valido) {
  const color = valido ? '#087443' : '#b5162c';
  return '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><meta charset="utf-8"></head>' +
    '<body style="margin:0;padding:28px 16px;background:#fff8ee;font-family:Arial,sans-serif;color:#242424">' +
    '<main style="max-width:520px;margin:30px auto;background:white;border:2px solid ' + color + ';border-radius:18px;padding:28px;text-align:center">' +
    '<div style="font-size:25px;font-weight:900;color:#a40f24">MACATEXPERU</div><p style="color:#075a3e">Rifa Navideña 2026</p>' +
    '<h1 style="color:' + color + '">' + titulo + '</h1><p style="font-size:17px;line-height:1.6">' + detalle + '</p>' +
    '<p style="font-size:12px;color:#777">Consulta del registro oficial de MACATEXPERU.</p></main></body></html>';
}

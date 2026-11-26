/**
 * MACATEX - ADMINISTRACIÓN DE PEDIDOS Y PAGOS
 * Uso: copiar este archivo a Apps Script como MACATEX_Admin.gs.
 * No reemplaza Código.gs ni modifica el módulo de catálogo o las funciones de rifas.
 *
 * Flujo:
 * 1) Registrar pago -> queda PENDIENTE DE VERIFICACIÓN.
 * 2) Verificar pago -> actualiza el acumulado recibido y el saldo.
 * 3) Con adelanto verificado, autorizar compra de insumos.
 * 4) Al recibir los insumos, habilitar el pedido para el flujo existente
 *    aceptarPedido() -> reserva stock; entregarPedido() -> descuenta stock.
 */

function macatexConsultarPedidos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('PEDIDOS');
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Todavía no hay pedidos registrados.');
    return;
  }
  const h = macatexAdminHeaders_(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const idx = n => h[n];
  const lines = data.slice(-15).reverse().map(r => {
    const id = r[idx('Pedido ID')] || '';
    const cliente = r[idx('Cliente')] || '';
    const estado = r[idx('Estado pedido')] || '';
    const total = Number(r[idx('Total (S/)')]) || 0;
    const saldo = Number(r[idx('Saldo pendiente (S/)')]) || 0;
    return id + ' | ' + cliente + ' | ' + estado + ' | Total S/ ' +
      total.toFixed(2) + ' | Saldo S/ ' + saldo.toFixed(2);
  });
  SpreadsheetApp.getUi().alert('MACATEX - Últimos pedidos', lines.join('\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function macatexRegistrarPago() {
  const ui = SpreadsheetApp.getUi();
  const pedidoId = macatexPrompt_(ui, 'Registrar pago', 'Escribe el Pedido ID:');
  if (pedidoId === null) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidos = ss.getSheetByName('PEDIDOS');
  if (!pedidos) throw new Error('No existe la hoja PEDIDOS.');
  const hp = macatexAdminHeaders_(pedidos);
  const fila = macatexBuscarPedido_(pedidos, hp, pedidoId);
  if (!fila) throw new Error('No se encontró el pedido ' + pedidoId + '.');
  const pedido = pedidos.getRange(fila, 1, 1, pedidos.getLastColumn()).getValues()[0];
  const total = Number(pedido[hp['Total (S/)']]) || 0;
  const recibido = Number(pedido[hp['Adelanto recibido (S/)']]) || 0;
  const saldo = Number(pedido[hp['Saldo pendiente (S/)']]) || Math.max(0, total - recibido);
  const adelantoReq = Number(pedido[hp['Adelanto requerido (S/)']]) || Math.round(total * 50) / 100;

  const concepto = macatexPrompt_(ui, 'Concepto del pago', 'Escribe ADELANTO o SALDO:');
  if (concepto === null) return;
  const tipo = concepto.trim().toUpperCase();
  if (!['ADELANTO', 'SALDO'].includes(tipo)) throw new Error('El concepto debe ser ADELANTO o SALDO.');
  if (tipo === 'SALDO' && recibido < adelantoReq) {
    throw new Error('Primero debe estar verificado el adelanto requerido de S/ ' + adelantoReq.toFixed(2) + '.');
  }
  if (tipo === 'SALDO' && String(pedido[hp['Estado pedido']] || '').toUpperCase() !== 'ENTREGADO') {
    throw new Error('El saldo final solo se registra cuando el pedido ya fue entregado.');
  }

  const importeTxt = macatexPrompt_(ui, 'Importe recibido', 'Importe del pago en soles (ejemplo: 125.00):');
  if (importeTxt === null) return;
  const importe = Number(importeTxt.replace(',', '.'));
  if (!Number.isFinite(importe) || importe <= 0) throw new Error('El importe debe ser mayor que cero.');
  const pendienteAdelanto = Math.max(0, adelantoReq - recibido);
  const maximo = tipo === 'ADELANTO' ? pendienteAdelanto : saldo;
  if (importe > maximo + 0.001) {
    throw new Error('El importe supera lo pendiente para ' + tipo + ': S/ ' + maximo.toFixed(2) + '.');
  }

  const medio = macatexPrompt_(ui, 'Medio de pago', 'Escribe YAPE, PLIN, TRANSFERENCIA o EFECTIVO:');
  if (medio === null) return;
  if (!medio.trim()) throw new Error('Debes indicar el medio de pago.');
  const referencia = macatexPrompt_(ui, 'Referencia', 'Número de operación o referencia (si no aplica, escribe N/A):');
  if (referencia === null) return;
  const observacion = macatexPrompt_(ui, 'Observaciones', 'Observación opcional:') || '';

  const pagos = macatexAdminHojaPagos_(ss);
  const hpago = macatexAdminHeaders_(pagos);
  const pagoId = 'PAG-' + Utilities.getUuid().split('-')[0].toUpperCase();
  const datos = {
    'Pago ID': pagoId,
    'Pedido ID': pedidoId,
    'Fecha de pago': new Date(),
    'Concepto': tipo,
    'Importe (S/)': Math.round(importe * 100) / 100,
    'Medio de pago': medio.trim().toUpperCase(),
    'Referencia operación': referencia.trim(),
    'Estado verificación': 'PENDIENTE DE VERIFICACIÓN',
    'Observaciones': observacion,
    'Responsable verificación': '',
    'Fecha verificación': ''
  };
  pagos.appendRow(macatexAdminFila_(pagos, hpago, datos));
  ui.alert('Pago registrado para revisión', 'Código de pago: ' + pagoId +
    '\nPedido: ' + pedidoId + '\nImporte: S/ ' + importe.toFixed(2) +
    '\nEstado: PENDIENTE DE VERIFICACIÓN\n\nEl importe todavía no se suma al dinero recibido.', ui.ButtonSet.OK);
}

function macatexVerificarPago() {
  const ui = SpreadsheetApp.getUi();
  const pagoId = macatexPrompt_(ui, 'Verificar pago', 'Escribe el Pago ID (PAG-...):');
  if (pagoId === null) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pagos = ss.getSheetByName('PAGOS');
  const pedidos = ss.getSheetByName('PEDIDOS');
  if (!pagos || !pedidos) throw new Error('Falta la hoja PAGOS o PEDIDOS.');
  const hpay = macatexAdminHeaders_(pagos);
  const hp = macatexAdminHeaders_(pedidos);
  const filaPago = macatexBuscarPago_(pagos, hpay, pagoId);
  if (!filaPago) throw new Error('No se encontró el pago ' + pagoId + '.');
  const pago = pagos.getRange(filaPago, 1, 1, pagos.getLastColumn()).getValues()[0];
  if (String(pago[hpay['Estado verificación']] || '').toUpperCase() !== 'PENDIENTE DE VERIFICACIÓN') {
    throw new Error('Este pago no está pendiente de verificación o ya fue procesado.');
  }
  const pedidoId = String(pago[hpay['Pedido ID']] || '').trim();
  const filaPedido = macatexBuscarPedido_(pedidos, hp, pedidoId);
  if (!filaPedido) throw new Error('No se encontró el pedido relacionado: ' + pedidoId);
  const pedido = pedidos.getRange(filaPedido, 1, 1, pedidos.getLastColumn()).getValues()[0];
  const total = Number(pedido[hp['Total (S/)']]) || 0;
  const recibidoActual = Number(pedido[hp['Adelanto recibido (S/)']]) || 0;
  const importe = Number(pago[hpay['Importe (S/)']]) || 0;
  const concepto = String(pago[hpay['Concepto']] || '').toUpperCase();
  const adelantoReq = Number(pedido[hp['Adelanto requerido (S/)']]) || Math.round(total * 50) / 100;
  const saldoActual = Math.max(0, total - recibidoActual);
  if (importe <= 0 || recibidoActual + importe > total + 0.001) throw new Error('El importe no es válido o supera el saldo total.');
  if (concepto === 'ADELANTO' && recibidoActual + importe > adelantoReq + 0.001) {
    throw new Error('El pago supera el adelanto requerido. Regístralo con el concepto correcto.');
  }
  if (concepto === 'SALDO') {
    if (recibidoActual < adelantoReq) throw new Error('El adelanto aún no está completo.');
    if (String(pedido[hp['Estado pedido']] || '').toUpperCase() !== 'ENTREGADO') {
      throw new Error('No se puede verificar el saldo antes de la entrega.');
    }
  }

  const confirmar = ui.alert('Confirmación de verificación',
    '¿Has comprobado el abono directamente en la cuenta o billetera de MACATEX?\n\n' +
    'Pedido: ' + pedidoId + '\nPago: ' + pagoId + '\nImporte: S/ ' + importe.toFixed(2) +
    '\nMedio: ' + pago[hpay['Medio de pago']] +
    '\nReferencia: ' + pago[hpay['Referencia operación']],
    ui.ButtonSet.YES_NO);
  if (confirmar !== ui.Button.YES) return;

  const recibidoNuevo = Math.round((recibidoActual + importe) * 100) / 100;
  const saldoNuevo = Math.round((total - recibidoNuevo) * 100) / 100;
  const estadoPago = recibidoNuevo >= total - 0.001
    ? 'PAGADO'
    : recibidoNuevo >= adelantoReq - 0.001 ? 'ADELANTO VERIFICADO' : 'ADELANTO PARCIAL';

  pagos.getRange(filaPago, hpay['Estado verificación'] + 1).setValue('VERIFICADO');
  pagos.getRange(filaPago, hpay['Responsable verificación'] + 1)
    .setValue(Session.getActiveUser().getEmail() || 'Administrador');
  pagos.getRange(filaPago, hpay['Fecha verificación'] + 1).setValue(new Date());

  pedidos.getRange(filaPedido, hp['Adelanto recibido (S/)'] + 1).setValue(recibidoNuevo);
  pedidos.getRange(filaPedido, hp['Saldo pendiente (S/)'] + 1).setValue(saldoNuevo);
  pedidos.getRange(filaPedido, hp['Estado pago'] + 1).setValue(estadoPago);
  pedidos.getRange(filaPedido, hp['Fecha actualización'] + 1).setValue(new Date());
  if (recibidoNuevo >= adelantoReq - 0.001 &&
      String(pedido[hp['Estado pedido']] || '').toUpperCase() === 'SOLICITUD RECIBIDA') {
    pedidos.getRange(filaPedido, hp['Estado pedido'] + 1).setValue('ADELANTO VERIFICADO');
  }
  SpreadsheetApp.flush();
  ui.alert('Pago verificado', 'Pedido: ' + pedidoId +
    '\nRecibido acumulado: S/ ' + recibidoNuevo.toFixed(2) +
    '\nSaldo pendiente: S/ ' + saldoNuevo.toFixed(2) +
    '\nEstado de pago: ' + estadoPago, ui.ButtonSet.OK);
}

function macatexAutorizarCompraInsumos() {
  const ui = SpreadsheetApp.getUi();
  const pedidoId = macatexPrompt_(ui, 'Autorizar compra de insumos', 'Escribe el Pedido ID:');
  if (pedidoId === null) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PEDIDOS');
  const h = macatexAdminHeaders_(sh);
  const fila = macatexBuscarPedido_(sh, h, pedidoId);
  if (!fila) throw new Error('No se encontró el pedido ' + pedidoId + '.');
  const r = sh.getRange(fila, 1, 1, sh.getLastColumn()).getValues()[0];
  const req = Number(r[h['Adelanto requerido (S/)']]) || Math.round(Number(r[h['Total (S/)']]) * 50) / 100;
  const rec = Number(r[h['Adelanto recibido (S/)']]) || 0;
  if (rec < req - 0.001) throw new Error('No se puede autorizar la compra: adelanto verificado insuficiente. Requerido S/ ' + req.toFixed(2) + ', recibido S/ ' + rec.toFixed(2) + '.');
  if (!['ADELANTO VERIFICADO', 'COMPRA DE INSUMOS AUTORIZADA'].includes(String(r[h['Estado pedido']] || '').toUpperCase())) {
    throw new Error('El estado actual no permite autorizar la compra.');
  }
  sh.getRange(fila, h['Estado pedido'] + 1).setValue('COMPRA DE INSUMOS AUTORIZADA');
  sh.getRange(fila, h['Fecha actualización'] + 1).setValue(new Date());
  ui.alert('Compra autorizada', 'El adelanto está verificado. Ya puedes adquirir los insumos del pedido ' + pedidoId + '.', ui.ButtonSet.OK);
}

function macatexRegistrarInsumosRecibidos() {
  const ui = SpreadsheetApp.getUi();
  const pedidoId = macatexPrompt_(ui, 'Registrar recepción de insumos', 'Escribe el Pedido ID:');
  if (pedidoId === null) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedidos = ss.getSheetByName('PEDIDOS');
  const detalle = ss.getSheetByName('DETALLE_PEDIDOS');
  if (!pedidos || !detalle) throw new Error('Falta PEDIDOS o DETALLE_PEDIDOS.');
  const hp = macatexAdminHeaders_(pedidos);
  const hd = macatexAdminHeaders_(detalle);
  const fila = macatexBuscarPedido_(pedidos, hp, pedidoId);
  if (!fila) throw new Error('No se encontró el pedido ' + pedidoId + '.');
  const estado = String(pedidos.getRange(fila, hp['Estado pedido'] + 1).getValue()).toUpperCase();
  if (estado !== 'COMPRA DE INSUMOS AUTORIZADA') throw new Error('Primero debes autorizar la compra de insumos.');
  const idxId = hd['Pedido ID'];
  const idxEstado = hd['Estado línea'];
  let count = 0;
  if (idxId === undefined || idxEstado === undefined) throw new Error('Faltan encabezados en DETALLE_PEDIDOS.');
  if (detalle.getLastRow() > 1) {
    const values = detalle.getRange(2, 1, detalle.getLastRow() - 1, detalle.getLastColumn()).getValues();
    values.forEach((r, i) => {
      if (String(r[idxId]).trim() === pedidoId) {
        detalle.getRange(i + 2, idxEstado + 1).setValue('PENDIENTE');
        count++;
      }
    });
  }
  if (!count) throw new Error('No se encontraron líneas del pedido.');
  pedidos.getRange(fila, hp['Estado pedido'] + 1).setValue('PENDIENTE');
  pedidos.getRange(fila, hp['Fecha actualización'] + 1).setValue(new Date());
  ui.alert('Insumos recibidos', 'Se habilitó el pedido ' + pedidoId + ' para el flujo existente de inventario. Ahora puedes usar MACATEX > Aceptar y reservar pedido.', ui.ButtonSet.OK);
}

function macatexCerrarVenta() {
  const ui = SpreadsheetApp.getUi();
  const pedidoId = macatexPrompt_(ui, 'Cerrar venta', 'Escribe el Pedido ID:');
  if (pedidoId === null) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PEDIDOS');
  const h = macatexAdminHeaders_(sh);
  const fila = macatexBuscarPedido_(sh, h, pedidoId);
  if (!fila) throw new Error('No se encontró el pedido ' + pedidoId + '.');
  const r = sh.getRange(fila, 1, 1, sh.getLastColumn()).getValues()[0];
  if (String(r[h['Estado pedido']] || '').toUpperCase() !== 'ENTREGADO') throw new Error('La venta solo se puede cerrar después de registrar la entrega.');
  if ((Number(r[h['Saldo pendiente (S/)']]) || 0) > 0.001) throw new Error('No se puede cerrar: todavía existe saldo pendiente.');
  if (String(r[h['Estado pago']] || '').toUpperCase() !== 'PAGADO') throw new Error('El estado de pago todavía no es PAGADO.');
  sh.getRange(fila, h['Estado pedido'] + 1).setValue('CERRADO');
  sh.getRange(fila, h['Fecha actualización'] + 1).setValue(new Date());
  ui.alert('Venta cerrada', 'El pedido ' + pedidoId + ' está entregado y pagado en su totalidad.', ui.ButtonSet.OK);
}

/** Llamar una sola vez desde onOpen() existente para añadir estas opciones al menú MACATEX. */
function macatexAgregarMenuPedidos_(menu) {
  menu.addSeparator()
    .addItem('Consultar últimos pedidos', 'macatexConsultarPedidos')
    .addItem('Registrar pago recibido', 'macatexRegistrarPago')
    .addItem('Verificar pago recibido', 'macatexVerificarPago')
    .addItem('Autorizar compra de insumos', 'macatexAutorizarCompraInsumos')
    .addItem('Registrar insumos recibidos', 'macatexRegistrarInsumosRecibidos')
    .addItem('Cerrar venta', 'macatexCerrarVenta');
}

function macatexAdminHojaPagos_(ss) {
  const headers = ['Pago ID', 'Pedido ID', 'Fecha de pago', 'Concepto', 'Importe (S/)',
    'Medio de pago', 'Referencia operación', 'Estado verificación', 'Observaciones',
    'Responsable verificación', 'Fecha verificación'];
  let sh = ss.getSheetByName('PAGOS');
  if (!sh) sh = ss.insertSheet('PAGOS');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(x => String(x).trim());
    headers.forEach(x => {
      if (!current.includes(x)) sh.getRange(1, sh.getLastColumn() + 1).setValue(x)
        .setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    });
  }
  return sh;
}

function macatexAdminHeaders_(sh) {
  if (!sh || sh.getLastColumn() < 1) throw new Error('No se encontró una hoja con encabezados.');
  const names = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(x => String(x).trim());
  const map = {};
  names.forEach((n, i) => { if (n) map[n] = i; });
  return map;
}

function macatexBuscarPedido_(sh, h, id) {
  if (h['Pedido ID'] === undefined || sh.getLastRow() < 2) return 0;
  const values = sh.getRange(2, h['Pedido ID'] + 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const i = values.findIndex(x => String(x).trim() === id);
  return i < 0 ? 0 : i + 2;
}

function macatexBuscarPago_(sh, h, id) {
  if (h['Pago ID'] === undefined || sh.getLastRow() < 2) return 0;
  const values = sh.getRange(2, h['Pago ID'] + 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const i = values.findIndex(x => String(x).trim() === id);
  return i < 0 ? 0 : i + 2;
}

function macatexAdminFila_(sh, h, values) {
  const row = Array(sh.getLastColumn()).fill('');
  Object.keys(values).forEach(k => { if (h[k] !== undefined) row[h[k]] = values[k]; });
  return row;
}

function macatexPrompt_(ui, title, message) {
  const r = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const value = r.getResponseText().trim();
  if (!value) throw new Error('El dato solicitado no puede estar vacío.');
  return value;
}

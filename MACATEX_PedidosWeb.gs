/**
 * MACATEX PERÚ - Registro web de solicitudes de pedido.
 *
 * INSTALACIÓN:
 * 1. Crear un archivo nuevo en el proyecto Apps Script llamado PedidosWeb.gs.
 * 2. Copiar este contenido completo.
 * 3. Guardar y actualizar la implementación de aplicación web a una nueva versión.
 *
 * Este módulo no modifica el inventario ni reserva unidades.
 * La solicitud queda pendiente hasta que MACATEX confirme condiciones y verifique el adelanto.
 */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return respuestaPedidoWeb_({ ok: false, error: 'El sistema está ocupado. Intenta nuevamente.' });
    }

    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
    if (!raw) throw new Error('No se recibió información del pedido.');

    const payload = JSON.parse(raw);
    const pedidoId = String(payload.pedidoId || '').trim();
    const cliente = textoPedidoWeb_(payload.cliente, 120);
    const whatsapp = String(payload.whatsapp || '').replace(/\D/g, '').slice(0, 15);
    const ciudad = textoPedidoWeb_(payload.ciudad, 120);
    const fechaEntrega = String(payload.fechaEntrega || '').trim();
    const observaciones = textoPedidoWeb_(payload.observaciones, 1000);
    const items = Array.isArray(payload.items) ? payload.items : [];

    if (!/^MAC-[A-Z0-9-]{8,40}$/.test(pedidoId)) throw new Error('Código de solicitud inválido.');
    if (!cliente || !whatsapp || !ciudad) throw new Error('Faltan datos obligatorios del cliente.');
    if (!items.length || items.length > 40) throw new Error('El pedido debe contener entre 1 y 40 líneas.');

    let total = 0;
    const lineas = items.map((item) => {
      const codigo = textoPedidoWeb_(item.codigo, 40);
      const nombre = textoPedidoWeb_(item.nombre, 160);
      const cantidad = Number(item.cantidad);
      const precio = Number(item.precioUnitario);
      if (!codigo || !nombre || !Number.isInteger(cantidad) || cantidad < 1 || cantidad > 500) {
        throw new Error('Hay un producto o cantidad inválida.');
      }
      if (!Number.isFinite(precio) || precio <= 0 || precio > 100000) {
        throw new Error('Hay un precio unitario inválido.');
      }
      const subtotal = Math.round(cantidad * precio * 100) / 100;
      total += subtotal;
      return { codigo, nombre, cantidad, precio, subtotal };
    });

    total = Math.round(total * 100) / 100;
    const adelanto = Math.round(total * 50) / 100;
    const saldo = Math.round((total - adelanto) * 100) / 100;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No se encontró el archivo de Google Sheets vinculado.');

    const pedidos = hojaPedidoWeb_(ss, 'PEDIDOS', [
      'Pedido ID', 'Fecha registro', 'Cliente', 'WhatsApp', 'Canal',
      'Estado pedido', 'Estado pago', 'Total (S/)', 'Fecha entrega',
      'Observaciones', 'Adelanto requerido (S/)', 'Adelanto recibido (S/)',
      'Saldo pendiente (S/)', 'Medio de pago', 'Fecha actualización', 'Distrito/ciudad'
    ]);
    const detalle = hojaPedidoWeb_(ss, 'DETALLE_PEDIDOS', [
      'Pedido ID', 'Código producto', 'Nombre comercial', 'Cantidad',
      'Precio unitario (S/)', 'Subtotal (S/)', 'Estado línea'
    ]);
    hojaPedidoWeb_(ss, 'PAGOS', [
      'Pago ID', 'Pedido ID', 'Fecha de pago', 'Concepto',
      'Importe (S/)', 'Medio de pago', 'Referencia operación',
      'Estado verificación', 'Observaciones'
    ]);

    const hp = pedidos.getRange(1, 1, 1, pedidos.getLastColumn()).getDisplayValues()[0];
    const idCol = hp.indexOf('Pedido ID');
    if (idCol < 0) throw new Error('No se encontró la columna Pedido ID en PEDIDOS.');
    const existentes = pedidos.getLastRow() > 1
      ? pedidos.getRange(2, idCol + 1, pedidos.getLastRow() - 1, 1).getDisplayValues().flat()
      : [];
    if (existentes.includes(pedidoId)) {
      return respuestaPedidoWeb_({ ok: true, duplicate: true, pedidoId, total, adelanto, saldo });
    }

    const ahora = new Date();
    const datosPedido = {
      'Pedido ID': pedidoId,
      'Fecha registro': ahora,
      'Cliente': cliente,
      'WhatsApp': whatsapp,
      'Canal': 'WEB',
      'Estado pedido': 'SOLICITUD RECIBIDA',
      'Estado pago': 'PENDIENTE DE COTIZACIÓN',
      'Total (S/)': total,
      'Fecha entrega': fechaEntrega,
      'Observaciones': observaciones,
      'Adelanto requerido (S/)': adelanto,
      'Adelanto recibido (S/)': 0,
      'Saldo pendiente (S/)': saldo,
      'Medio de pago': '',
      'Fecha actualización': ahora,
      'Distrito/ciudad': ciudad
    };
    pedidos.appendRow(hp.map(h => Object.prototype.hasOwnProperty.call(datosPedido, h)
      ? valorSeguroPedidoWeb_(datosPedido[h]) : ''));

    const hd = detalle.getRange(1, 1, 1, detalle.getLastColumn()).getDisplayValues()[0];
    lineas.forEach(linea => {
      const fila = {
        'Pedido ID': pedidoId,
        'Código producto': linea.codigo,
        'Nombre comercial': linea.nombre,
        'Cantidad': linea.cantidad,
        'Precio unitario (S/)': linea.precio,
        'Subtotal (S/)': linea.subtotal,
        'Estado línea': 'SOLICITADA'
      };
      detalle.appendRow(hd.map(h => Object.prototype.hasOwnProperty.call(fila, h)
        ? valorSeguroPedidoWeb_(fila[h]) : ''));
    });

    SpreadsheetApp.flush();
    return respuestaPedidoWeb_({ ok: true, duplicate: false, pedidoId, total, adelanto, saldo });
  } catch (err) {
    return respuestaPedidoWeb_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function hojaPedidoWeb_(ss, nombre, encabezados) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);

  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    sh.getRange(1, 1, 1, encabezados.length)
      .setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  } else {
    const actuales = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
      .map(h => String(h).trim());
    encabezados.forEach(h => {
      if (!actuales.includes(h)) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h)
          .setFontWeight('bold').setBackground('#1E3A8A').setFontColor('#FFFFFF');
      }
    });
  }
  return sh;
}

function textoPedidoWeb_(valor, max) {
  return String(valor == null ? '' : valor).replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, max);
}

function valorSeguroPedidoWeb_(valor) {
  if (typeof valor !== 'string') return valor;
  return /^[=+@\-]/.test(valor) ? "'" + valor : valor;
}

function respuestaPedidoWeb_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

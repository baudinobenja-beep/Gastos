/**
 * Recibe los gastos que manda la app del celular y los guarda en esta planilla.
 * Las fotos de los tickets van a la carpeta "Gastos - tickets" de tu Google Drive.
 *
 * Instalación (una sola vez):
 * 1. En la planilla: Extensiones > Apps Script. Borrá lo que haya y pegá este código.
 * 2. Cambiá CLAVE por una palabra tuya (la misma que vas a poner en la app).
 * 3. Implementar > Nueva implementación > Tipo: Aplicación web.
 *    Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
 * 4. Autorizá los permisos y copiá la dirección que termina en /exec.
 */
const CLAVE = 'cambiar-esta-clave';
const HOJA = 'Gastos';
const CARPETA = 'Gastos - tickets';
const COLUMNAS = [
  'codigo', 'fecha', 'mes', 'monto', 'tipo', 'categoria', 'rubro', 'campo',
  'medioPago', 'comprobante', 'cuit', 'comercio', 'nota', 'fotoUrl',
  'envio', 'envioFecha', 'reintegro', 'creado', 'actualizado'
];
const TITULOS = [
  'Código', 'Fecha', 'Mes', 'Monto', 'Tipo', 'Categoría', 'Actividad', 'Campo destino',
  'Medio de pago', 'Comprobante', 'CUIT', 'Comercio', 'Nota', 'Foto del ticket',
  'Envío', 'Fecha de envío', 'Reintegro', 'Creado', 'Actualizado'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const pedido = JSON.parse(e.postData.contents);
    if (pedido.clave !== CLAVE) return responder({ ok: false, error: 'clave incorrecta' });
    const g = pedido.gasto || {};
    if (!/^G-\d+$/.test(g.codigo || '')) return responder({ ok: false, error: 'gasto sin código' });

    const hoja = obtenerHoja();
    const fila = buscarFila(hoja, g.codigo);

    if (pedido.accion === 'borrar') {
      if (fila) hoja.deleteRow(fila);
      return responder({ ok: true });
    }

    if (g.fotoBase64) {
      const bytes = Utilities.base64Decode(g.fotoBase64);
      const archivo = obtenerCarpeta().createFile(
        Utilities.newBlob(bytes, g.fotoTipo || 'image/jpeg', g.codigo + '.jpg'));
      g.fotoUrl = archivo.getUrl();
    } else if (fila && !g.fotoUrl) {
      g.fotoUrl = hoja.getRange(fila, COLUMNAS.indexOf('fotoUrl') + 1).getValue();
    }
    g.actualizado = new Date().toISOString();

    const valores = [COLUMNAS.map(c => (g[c] === undefined || g[c] === null) ? '' : g[c])];
    if (fila) hoja.getRange(fila, 1, 1, COLUMNAS.length).setValues(valores);
    else hoja.appendRow(valores[0]);

    return responder({ ok: true, fotoUrl: g.fotoUrl || '' });
  } catch (err) {
    return responder({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return responder({ ok: true, mensaje: 'La planilla de gastos está lista para recibir datos.' });
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(HOJA);
    hoja.appendRow(TITULOS);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, TITULOS.length).setFontWeight('bold');
    hoja.getRange('D:D').setNumberFormat('$ #,##0.00');
  }
  return hoja;
}

function buscarFila(hoja, codigo) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return 0;
  const codigos = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  for (let i = 0; i < codigos.length; i++) if (codigos[i][0] === codigo) return i + 2;
  return 0;
}

function obtenerCarpeta() {
  const it = DriveApp.getFoldersByName(CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

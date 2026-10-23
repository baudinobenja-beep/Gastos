/**
 * Recibe los gastos que manda la app del celular y los guarda en la hoja "Gastos".
 * Cuando la app termina de sincronizar, arma una hoja ordenada para cada mes que
 * tocó (por ejemplo "Octubre 2026") con el resumen y el detalle de empresa y personal.
 * Las fotos de los tickets se quedan en el celular: no se copian a la planilla.
 *
 * Instalación (una sola vez):
 * 1. En la planilla: Extensiones > Apps Script. Borrá lo que haya y pegá este código.
 * 2. Cambiá CLAVE por una palabra tuya (la misma que vas a poner en la app).
 * 3. Implementar > Nueva implementación > Tipo: Aplicación web.
 *    Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
 * 4. Autorizá los permisos y copiá la dirección que termina en /exec.
 * 5. Si querés rearmar un mes a mano: recargá la planilla y usá el menú Gastos.
 *
 * Para actualizar el código más adelante: pegalo, guardá, y en
 * Implementar > Gestionar implementaciones > lápiz > Versión: Nueva versión > Implementar.
 * La dirección /exec no cambia.
 */
const CLAVE = 'cambiar-esta-clave';
const HOJA = 'Gastos';
const ZONA = 'America/Argentina/Buenos_Aires';
const MONEDA = '$ #,##0.00';
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
  'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
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

/* ---------- recibir gastos de la app ---------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const pedido = JSON.parse(e.postData.contents);
    if (pedido.clave !== CLAVE) return responder({ ok: false, error: 'clave incorrecta' });
    if (pedido.accion === 'armar') {
      const meses = [...new Set((pedido.meses || []).map(textoMes).filter(Boolean))];
      const hojas = meses.map(armarHojaMes).filter(Boolean).map(h => h.getName());
      return responder({ ok: true, hojas: hojas });
    }
    const g = pedido.gasto || {};
    if (!/^G-\d+$/.test(g.codigo || '')) return responder({ ok: false, error: 'gasto sin código' });

    const hoja = obtenerHoja();
    const fila = buscarFila(hoja, g.codigo);
    const mesAnterior = fila ? textoMes(hoja.getRange(fila, COLUMNAS.indexOf('mes') + 1).getValue()) : '';

    if (pedido.accion === 'borrar') {
      if (fila) hoja.deleteRow(fila);
      return responder({ ok: true, meses: [mesAnterior] });
    }

    // fotos: no se guardan; se conserva el link de las que ya se habían copiado
    if (fila) g.fotoUrl = hoja.getRange(fila, COLUMNAS.indexOf('fotoUrl') + 1).getValue();
    else g.fotoUrl = '';
    g.actualizado = new Date().toISOString();

    const valores = [COLUMNAS.map(c => (g[c] === undefined || g[c] === null) ? '' : g[c])];
    if (fila) hoja.getRange(fila, 1, 1, COLUMNAS.length).setValues(valores);
    else hoja.appendRow(valores[0]);

    // el mes viejo también se rearma si al editar le cambiaron la fecha
    return responder({ ok: true, meses: [textoMes(g.mes), mesAnterior] });
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
    hoja.getRange('C:C').setNumberFormat('@'); // keep "2026-10" as text, not a date
    hoja.getRange('D:D').setNumberFormat(MONEDA);
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

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- hoja ordenada de cada mes ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gastos')
    .addItem('Armar hoja de este mes', 'armarEsteMes')
    .addItem('Armar hoja del mes anterior', 'armarMesAnterior')
    .addToUi();
}

function armarEsteMes() { armarHojaMes(mesDe(new Date(), 0)); }
function armarMesAnterior() { armarHojaMes(mesDe(new Date(), -1)); }

// "2026-10" del mes de esa fecha, corrido `desplazamiento` meses
function mesDe(fecha, desplazamiento) {
  const y = Number(Utilities.formatDate(fecha, ZONA, 'yyyy'));
  const m = Number(Utilities.formatDate(fecha, ZONA, 'M')) - 1 + desplazamiento;
  const d = new Date(y, m, 1);
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
}

function nombreHojaMes(ym) {
  const [y, m] = ym.split('-').map(Number);
  return MESES[m - 1] + ' ' + y;
}

function textoMes(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}/.test(s) ? s.slice(0, 7) : '';
}

function comoFecha(v) {
  if (v instanceof Date) return v;
  const p = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return p ? new Date(Number(p[1]), Number(p[2]) - 1, Number(p[3])) : '';
}

function leerGastosDelMes(ym) {
  const hoja = obtenerHoja();
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  return hoja.getRange(2, 1, ultima - 1, COLUMNAS.length).getValues()
    .map(fila => {
      const g = {};
      COLUMNAS.forEach((c, i) => { g[c] = fila[i]; });
      g.fecha = comoFecha(g.fecha);
      g.mes = textoMes(g.mes) || (g.fecha ? Utilities.formatDate(g.fecha, ZONA, 'yyyy-MM') : '');
      g.monto = Number(g.monto) || 0;
      return g;
    })
    .filter(g => g.mes === ym)
    .sort((a, b) => (a.fecha - b.fecha) || String(a.codigo).localeCompare(String(b.codigo)));
}

function sumarPor(gastos, campo) {
  const totales = {};
  gastos.forEach(g => {
    const k = String(g[campo] || 'Sin dato');
    totales[k] = (totales[k] || 0) + g.monto;
  });
  return Object.keys(totales).sort((a, b) => totales[b] - totales[a]).map(k => [k, totales[k]]);
}

function sumar(gastos) { return gastos.reduce((t, g) => t + g.monto, 0); }

function armarHojaMes(ym) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const nombre = nombreHojaMes(ym);
  let hoja = libro.getSheetByName(nombre);
  const gastos = leerGastosDelMes(ym);
  if (!hoja && !gastos.length) return null; // no se crean hojas de meses sin gastos
  if (hoja) hoja.clear();
  else hoja = libro.insertSheet(nombre, 1); // los meses nuevos quedan primero, al lado de "Gastos"

  const empresa = gastos.filter(g => g.tipo === 'Empresa');
  const personal = gastos.filter(g => g.tipo !== 'Empresa');
  const deBolsillo = empresa.filter(g => g.medioPago === 'Transferencia personal');
  const sinReintegrar = deBolsillo.filter(g => g.reintegro !== 'Reintegrado');

  hoja.getRange(1, 1).setValue('Gastos de ' + nombre)
    .setFontSize(16).setFontWeight('bold');
  hoja.getRange(2, 1).setValue('Actualizado el ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm')
    + ' · ' + gastos.length + ' gastos').setFontColor('#666666');

  // resumen arriba, en bloques uno al lado del otro
  const bloques = [
    ['Resumen', [
      ['Empresa', sumar(empresa)],
      ['Personal', sumar(personal)],
      ['Total del mes', sumar(gastos)],
      ['Pagado de tu bolsillo (empresa)', sumar(deBolsillo)],
      ['Falta que te reintegren', sumar(sinReintegrar)]
    ]],
    ['Empresa por categoría', sumarPor(empresa, 'categoria')],
    ['Empresa por actividad', sumarPor(empresa, 'rubro')],
    ['Empresa por campo', sumarPor(empresa, 'campo')],
    ['Empresa por medio de pago', sumarPor(empresa, 'medioPago')],
    ['Personal por categoría', sumarPor(personal, 'categoria')]
  ];
  let alto = 0;
  bloques.forEach(([titulo, filas], i) => {
    const col = 1 + i * 3;
    const datos = filas.length ? filas : [['Sin gastos', 0]];
    hoja.getRange(4, col, 1, 2).setValues([[titulo, '']]).setFontWeight('bold').setBackground('#e8f0e8');
    hoja.getRange(5, col, datos.length, 2).setValues(datos);
    hoja.getRange(5, col + 1, datos.length, 1).setNumberFormat(MONEDA);
    alto = Math.max(alto, datos.length + 1);
  });

  let fila = 4 + alto + 2;
  fila = escribirTabla(hoja, fila, 'Gastos de empresa',
    ['Fecha', 'Código', 'Categoría', 'Actividad', 'Campo destino', 'Medio de pago', 'Comprobante',
      'CUIT', 'Comercio', 'Nota', 'Envío', 'Reintegro', 'Monto'],
    empresa.map(g => [g.fecha, g.codigo, g.categoria, g.rubro, g.campo, g.medioPago, g.comprobante,
      g.cuit, g.comercio, g.nota, g.envio === 'Pendiente' ? 'Sin enviar' : g.envio, g.reintegro, g.monto]));
  escribirTabla(hoja, fila + 2, 'Gastos personales',
    ['Fecha', 'Código', 'Categoría', 'Medio de pago', 'Comercio', 'Nota', 'Monto'],
    personal.map(g => [g.fecha, g.codigo, g.categoria, g.medioPago, g.comercio, g.nota, g.monto]));

  hoja.autoResizeColumns(1, 18);
  return hoja;
}

// Escribe título, encabezados, filas y total; devuelve la última fila usada.
function escribirTabla(hoja, fila, titulo, encabezados, filas) {
  const ancho = encabezados.length;
  hoja.getRange(fila, 1).setValue(titulo).setFontSize(13).setFontWeight('bold');
  hoja.getRange(fila + 1, 1, 1, ancho).setValues([encabezados])
    .setFontWeight('bold').setBackground('#2f5d50').setFontColor('#ffffff');
  let ultima = fila + 1;
  if (filas.length) {
    hoja.getRange(fila + 2, 1, filas.length, ancho).setValues(filas);
    hoja.getRange(fila + 2, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    hoja.getRange(fila + 2, ancho, filas.length, 1).setNumberFormat(MONEDA);
    ultima = fila + 1 + filas.length;
  } else {
    hoja.getRange(fila + 2, 1).setValue('Sin gastos este mes').setFontColor('#666666');
    ultima = fila + 2;
  }
  const total = filas.reduce((t, f) => t + (Number(f[ancho - 1]) || 0), 0);
  hoja.getRange(ultima + 1, ancho - 1, 1, 2).setValues([['Total', total]]).setFontWeight('bold');
  hoja.getRange(ultima + 1, ancho).setNumberFormat(MONEDA);
  return ultima + 1;
}

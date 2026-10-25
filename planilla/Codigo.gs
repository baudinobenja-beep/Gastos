/**
 * Recibe los gastos que manda la app del celular y los guarda en la hoja de su mes
 * (por ejemplo "Octubre 2026"). Cada hoja de mes tiene arriba el resumen con fórmulas
 * a la vista, abajo el cuadro de gastos con filtros, ordenado por fecha, y a la
 * derecha del cuadro la lista "Para reintegrar" (empresa pagado de tu bolsillo).
 * Las fotos de los tickets se quedan en el celular: no se copian a la planilla.
 *
 * Instalación (una sola vez):
 * 1. En la planilla: Extensiones > Apps Script. Borrá lo que haya y pegá este código.
 * 2. Cambiá CLAVE por una palabra tuya (la misma que vas a poner en la app).
 * 3. Implementar > Nueva implementación > Tipo: Aplicación web.
 *    Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
 * 4. Autorizá los permisos y copiá la dirección que termina en /exec.
 *
 * Para actualizar el código más adelante: pegalo, guardá, y en
 * Implementar > Gestionar implementaciones > lápiz > Versión: Nueva versión > Implementar.
 * La dirección /exec no cambia.
 */
const CLAVE = 'cambiar-esta-clave';
const ZONA = 'America/Argentina/Buenos_Aires';
const MONEDA = '$ #,##0.00';
const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
  'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
const HOJA_VIEJA = 'Gastos'; // versión anterior: todos los gastos juntos

// Cuadro de gastos de cada mes: encabezado en la fila 20, gastos desde la 21
const FILA_ENC = 20;
const FILA_DATOS = FILA_ENC + 1;
const CAMPOS = ['codigo', 'fecha', 'tipo', 'monto', 'categoria', 'rubro', 'campo',
  'medioPago', 'comprobante', 'cuit', 'comercio', 'nota', 'envio', 'reintegro'];
const TITULOS = ['Código', 'Fecha', 'Tipo', 'Monto', 'Categoría', 'Actividad', 'Campo destino',
  'Medio de pago', 'Comprobante', 'CUIT', 'Comercio', 'Nota', 'Envío', 'Reintegro'];
const ANCHO = CAMPOS.length;

/* ---------- recibir gastos de la app ---------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const pedido = JSON.parse(e.postData.contents);
    if (pedido.clave !== CLAVE) return responder({ ok: false, error: 'clave incorrecta' });
    pasarHojaViejaAMeses();

    // al terminar de sincronizar: dejar prolija la hoja de cada mes tocado
    if (pedido.accion === 'armar') {
      const hojas = [...new Set((pedido.meses || []).map(textoMes).filter(Boolean))]
        .map(m => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaMes(m)))
        .filter(Boolean);
      hojas.forEach(ordenarHoja);
      return responder({ ok: true, hojas: hojas.map(h => h.getName()) });
    }

    const g = pedido.gasto || {};
    if (!/^G-\d+$/.test(g.codigo || '')) return responder({ ok: false, error: 'gasto sin código' });
    const donde = buscarGasto(g.codigo);
    const mesAnterior = donde ? mesDeHoja(donde.hoja) : '';

    if (pedido.accion === 'borrar') {
      if (donde) donde.hoja.deleteRow(donde.fila);
      return responder({ ok: true, meses: [mesAnterior] });
    }

    const mes = textoMes(g.mes) || textoMes(g.fecha);
    if (!mes) return responder({ ok: false, error: 'gasto sin fecha' });
    const hoja = obtenerHojaMes(mes);
    const fila = filaDeGasto(g);
    if (donde && donde.hoja.getName() === hoja.getName()) {
      hoja.getRange(donde.fila, 1, 1, ANCHO).setValues([fila]);
    } else {
      if (donde) donde.hoja.deleteRow(donde.fila); // le cambiaron la fecha a otro mes
      const nueva = ultimaFila(hoja) + 1;
      if (nueva > hoja.getMaxRows()) hoja.insertRowsAfter(hoja.getMaxRows(), 200);
      hoja.getRange(nueva, 1, 1, ANCHO).setValues([fila]);
    }
    return responder({ ok: true, meses: [mes, mesAnterior] });
  } catch (err) {
    return responder({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return responder({ ok: true, mensaje: 'La planilla de gastos está lista para recibir datos.' });
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function filaDeGasto(g) {
  return CAMPOS.map(c => {
    if (c === 'fecha') return comoFecha(g.fecha);
    if (c === 'monto') return Number(g.monto) || 0;
    if (c === 'envio') return g.tipo !== 'Empresa' ? '' : (g.envio === 'Enviado' ? 'Enviado' : 'Sin enviar');
    return (g[c] === undefined || g[c] === null) ? '' : g[c];
  });
}

/* ---------- hojas de mes ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gastos')
    .addItem('Ordenar esta hoja', 'ordenarHojaActual')
    .addToUi();
}

function ordenarHojaActual() {
  pasarHojaViejaAMeses();
  const hoja = SpreadsheetApp.getActiveSheet();
  if (mesDeHoja(hoja)) ordenarHoja(hoja);
}

function nombreHojaMes(ym) {
  const [y, m] = ym.split('-').map(Number);
  return MESES[m - 1] + ' ' + y;
}

// "2026-10" si la hoja es de un mes, si no ''
function mesDeHoja(hoja) {
  const p = hoja.getName().match(/^(\S+) (\d{4})$/);
  const i = p ? MESES.indexOf(p[1]) : -1;
  return i < 0 ? '' : p[2] + '-' + ('0' + (i + 1)).slice(-2);
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

function hojasDeMes() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(h => mesDeHoja(h));
}

// Última fila del cuadro con un gasto (mira la columna Código; FILA_ENC si no hay ninguno)
function ultimaFila(hoja) {
  const v = hoja.getRange(FILA_DATOS, 1, hoja.getMaxRows() - FILA_ENC, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '') return FILA_DATOS + i;
  return FILA_ENC;
}

function buscarGasto(codigo) {
  for (const hoja of hojasDeMes()) {
    const ultima = ultimaFila(hoja);
    if (ultima < FILA_DATOS) continue;
    const codigos = hoja.getRange(FILA_DATOS, 1, ultima - FILA_ENC, 1).getValues();
    for (let i = 0; i < codigos.length; i++) {
      if (codigos[i][0] === codigo) return { hoja: hoja, fila: FILA_DATOS + i };
    }
  }
  return null;
}

function obtenerHojaMes(ym) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const nombre = nombreHojaMes(ym);
  let hoja = libro.getSheetByName(nombre);
  if (hoja) return hoja;
  // los meses quedan en orden, el más nuevo primero
  const despues = hojasDeMes().filter(h => mesDeHoja(h) > ym).length;
  hoja = libro.insertSheet(nombre, despues);
  armarResumen(hoja, nombre);
  return hoja;
}

// Resumen con fórmulas: todas leen el cuadro de abajo, así se ven y se pueden revisar.
function armarResumen(hoja, nombre) {
  const L = { codigo: 'A', tipo: 'C', monto: 'D', categoria: 'E', rubro: 'F', campo: 'G', medioPago: 'H', reintegro: 'N' };
  const R = c => '$' + L[c] + '$' + FILA_DATOS + ':$' + L[c];
  const monto = R('monto'), tipo = R('tipo');

  hoja.getRange('A1').setValue('Gastos de ' + nombre).setFontSize(16).setFontWeight('bold');
  hoja.getRange('A2').setValue('Filtrá el cuadro desde los títulos de la fila ' + FILA_ENC
    + '. "Total de lo filtrado" suma solo lo que queda a la vista.').setFontColor('#666666');

  hoja.getRange('A4:B11').setValues([
    ['Resumen', ''],
    ['Empresa', '=SUMIF(' + tipo + ',"Empresa",' + monto + ')'],
    ['Personal', '=SUMIF(' + tipo + ',"Personal",' + monto + ')'],
    ['Total del mes', '=SUM(' + monto + ')'],
    ['Pagado de tu bolsillo (empresa)', '=SUMIFS(' + monto + ',' + tipo + ',"Empresa",' + R('medioPago') + ',"Transferencia personal")'],
    ['Falta que te reintegren', '=SUMIFS(' + monto + ',' + tipo + ',"Empresa",' + R('medioPago') + ',"Transferencia personal",' + R('reintegro') + ',"<>Reintegrado")'],
    ['Cantidad de gastos', '=COUNTA(' + R('codigo') + ')'],
    ['Total de lo filtrado', '=SUBTOTAL(109,' + monto + ')']
  ]);
  hoja.getRange('B5:B9').setNumberFormat(MONEDA);
  hoja.getRange('B11').setNumberFormat(MONEDA);
  hoja.getRange('A11:B11').setFontWeight('bold');

  // bloques por rubro: la lista sale sola de los gastos y cada total es un SUMIFS
  const bloques = [
    ['D', 'Empresa por categoría', 'Empresa', 'categoria'],
    ['G', 'Empresa por actividad', 'Empresa', 'rubro'],
    ['J', 'Empresa por campo', 'Empresa', 'campo'],
    ['M', 'Empresa por medio de pago', 'Empresa', 'medioPago'],
    ['P', 'Personal por categoría', 'Personal', 'categoria']
  ];
  const ultimaFila = FILA_ENC - 2; // deja una fila libre antes del cuadro
  bloques.forEach(([letra, titulo, cual, campo]) => {
    const c = hoja.getRange(letra + '1').getColumn();
    const rango = hoja.getRange(5, c, ultimaFila - 4, 1).getA1Notation();
    hoja.getRange(4, c, 1, 2).setValues([[titulo, '']]);
    hoja.getRange(5, c).setFormula('=IFERROR(SORT(UNIQUE(FILTER(' + R(campo) + ',' + tipo + '="' + cual + '",' + R(campo) + '<>""))),"")');
    hoja.getRange(5, c + 1).setFormula('=ARRAYFORMULA(IF(' + rango + '="","",SUMIFS(' + monto + ',' + tipo + ',"' + cual + '",' + R(campo) + ',' + rango + ')))');
    hoja.getRange(5, c + 1, ultimaFila - 4, 1).setNumberFormat(MONEDA);
  });
  hoja.getRange(4, 1, 1, 17).setFontWeight('bold');
  ['A4:B4', 'D4:E4', 'G4:H4', 'J4:K4', 'M4:N4', 'P4:Q4'].forEach(a => hoja.getRange(a).setBackground('#e8f0e8'));

  hoja.getRange(FILA_ENC, 1, 1, ANCHO).setValues([TITULOS]).setFontWeight('bold');
  hoja.getRange(FILA_DATOS, 2, hoja.getMaxRows() - FILA_ENC, 1).setNumberFormat('dd/mm/yyyy');
  hoja.getRange(FILA_DATOS, 4, hoja.getMaxRows() - FILA_ENC, 1).setNumberFormat(MONEDA);
}

// Ordena por fecha y deja el cuadro con filtros y franjas sobre todos los gastos.
function ordenarHoja(hoja) {
  const filtro = hoja.getFilter();
  if (filtro) filtro.remove();
  hoja.getBandings().forEach(b => b.remove());
  const n = ultimaFila(hoja) - FILA_ENC;
  if (n) hoja.getRange(FILA_DATOS, 1, n, ANCHO).sort([{ column: 2, ascending: true }, { column: 1, ascending: true }]);
  const cuadro = hoja.getRange(FILA_ENC, 1, n + 1, ANCHO);
  cuadro.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false);
  cuadro.createFilter();
  armarReintegro(hoja);
  hoja.setFrozenRows(0);
  hoja.autoResizeColumns(1, 23);
}

// A la derecha del cuadro: gastos de empresa pagados de tu bolsillo que falta reintegrar.
// Es un FILTER sobre el cuadro, así se actualiza solo y se puede copiar tal cual.
const COL_REINT = 16; // P
function armarReintegro(hoja) {
  const d = c => '$' + c + '$' + FILA_DATOS + ':$' + c;
  hoja.getRange(FILA_ENC - 1, COL_REINT, 2, 8).setValues([
    ['Para reintegrar (pagado de tu bolsillo)', '', '', '', '', '', '', ''],
    ['Fecha', 'Código', 'Categoría', 'Campo destino', 'Comprobante', 'Comercio', 'Monto', 'Reintegro']
  ]);
  hoja.getRange(FILA_ENC - 1, COL_REINT + 6).setFormula('=SUM(' + d('V') + ')');
  hoja.getRange(FILA_DATOS, COL_REINT).setFormula('=IFERROR(FILTER({' + ['B', 'A', 'E', 'G', 'I', 'K', 'D', 'N'].map(d).join(',')
    + '},' + d('C') + '="Empresa",' + d('H') + '="Transferencia personal",' + d('N') + '<>"Reintegrado"),"")');
  hoja.getRange(FILA_ENC - 1, COL_REINT, 1, 8).setFontWeight('bold');
  hoja.getRange(FILA_ENC, COL_REINT, 1, 8).setFontWeight('bold').setBackground('#f4e4c8');
  hoja.getRange(FILA_ENC - 1, COL_REINT + 6).setNumberFormat(MONEDA);
  hoja.getRange(FILA_DATOS, COL_REINT, hoja.getMaxRows() - FILA_ENC, 1).setNumberFormat('dd/mm/yyyy');
  hoja.getRange(FILA_DATOS, COL_REINT + 6, hoja.getMaxRows() - FILA_ENC, 1).setNumberFormat(MONEDA);
}

// Pasa los gastos de la hoja "Gastos" (versión anterior) a la hoja de su mes y la borra.
function pasarHojaViejaAMeses() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const vieja = libro.getSheetByName(HOJA_VIEJA);
  if (!vieja) return;
  const datos = vieja.getDataRange().getValues();
  const enc = datos.shift() || [];
  const idx = t => enc.indexOf(t);
  const viejos = { codigo: 'Código', fecha: 'Fecha', mes: 'Mes', monto: 'Monto', tipo: 'Tipo', categoria: 'Categoría',
    rubro: 'Actividad', campo: 'Campo destino', medioPago: 'Medio de pago', comprobante: 'Comprobante', cuit: 'CUIT',
    comercio: 'Comercio', nota: 'Nota', envio: 'Envío', reintegro: 'Reintegro' };
  const porMes = {};
  datos.forEach(f => {
    const g = {};
    Object.keys(viejos).forEach(k => { const i = idx(viejos[k]); g[k] = i < 0 ? '' : f[i]; });
    if (!/^G-\d+$/.test(g.codigo)) return;
    const mes = textoMes(g.mes) || textoMes(g.fecha);
    if (mes) (porMes[mes] = porMes[mes] || []).push(filaDeGasto(g));
  });
  Object.keys(porMes).forEach(mes => {
    const anterior = libro.getSheetByName(nombreHojaMes(mes));
    if (anterior) libro.deleteSheet(anterior); // se rearma entera con los datos de "Gastos"
    const hoja = obtenerHojaMes(mes);
    hoja.getRange(FILA_DATOS, 1, porMes[mes].length, ANCHO).setValues(porMes[mes]);
    ordenarHoja(hoja);
  });
  libro.deleteSheet(vieja);
}

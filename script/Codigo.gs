/**
 * Plata — puente entre la app y la planilla.
 *
 * Pegá este código en Extensiones > Apps Script de tu planilla,
 * publicalo como aplicación web y pegá la dirección en la app.
 *
 * Las hojas se crean solas la primera vez.
 */

var HOJAS = {
  movs: {
    nombre: 'Plata · Movimientos',
    cols: ['id', 'tipo', 'fecha', 'monto', 'detalle', 'compania', 'categoria', 'medio', 'tarjetaId', 'cuotas', 'recurrente', 'mod']
  },
  tarjetas: {
    nombre: 'Plata · Tarjetas',
    cols: ['id', 'nombre', 'banco', 'cierre', 'vence', 'limite', 'mod']
  },
  inversiones: {
    nombre: 'Plata · Inversiones',
    cols: ['id', 'nombre', 'tipo', 'donde', 'monto', 'tna', 'fecha', 'vence', 'notas', 'mod']
  },
  objetivos: {
    nombre: 'Plata · Metas',
    cols: ['id', 'nombre', 'meta', 'fecha', 'aportes', 'mod']
  },
  borrados: {
    nombre: 'Plata · Borrados',
    cols: ['id', 'tabla', 'mod']
  }
};

/* ---------- utilidades de hoja ---------- */

function hoja_(clave) {
  var def = HOJAS[clave];
  if (def.nombre.indexOf('Plata') !== 0) {
    throw new Error('Plata solo escribe en sus propias hojas.');
  }
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var h = libro.getSheetByName(def.nombre);
  if (!h) {
    h = libro.insertSheet(def.nombre);
    h.getRange(1, 1, 1, def.cols.length).setValues([def.cols]);
    h.getRange(1, 1, 1, def.cols.length).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, def.cols.length).setValues([def.cols]);
    h.setFrozenRows(1);
  }
  return h;
}

function aTexto_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).trim();
}

function aNumero_(v) {
  if (typeof v === 'number') return v;
  var s = aTexto_(v).replace(/\$/g, '').replace(/\s/g, '');
  if (s === '') return 0;

  if (s.indexOf(',') > -1) {
    // formato argentino: 1.250.000,50 — el punto separa miles
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) {
    // 1.250.000 — puntos de miles, sin decimales
    s = s.replace(/\./g, '');
  }
  // cualquier otro caso (1250000 o 1250000.50) ya viene listo
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function leerHoja_(clave) {
  var def = HOJAS[clave];
  var h = hoja_(clave);
  var ultima = h.getLastRow();
  if (ultima < 2) return [];
  var valores = h.getRange(2, 1, ultima - 1, def.cols.length).getValues();
  var filas = [];
  for (var i = 0; i < valores.length; i++) {
    var fila = {};
    var vacia = true;
    for (var c = 0; c < def.cols.length; c++) {
      var val = valores[i][c];
      if (aTexto_(val) !== '') vacia = false;
      fila[def.cols[c]] = val;
    }
    if (!vacia) filas.push(normalizar_(clave, fila, i + 2));
  }
  return filas;
}

/* Convierte una fila de la planilla al formato que usa la app. */
function normalizar_(clave, f, numeroFila) {
  var mod = aNumero_(f.mod);
  var base = { mod: mod > 0 ? mod : 1 };

  if (clave === 'movs') {
    return {
      id: aTexto_(f.id) || ('sheet-' + numeroFila),
      tipo: aTexto_(f.tipo).toLowerCase() === 'ingreso' ? 'ingreso' : 'gasto',
      fecha: aTexto_(f.fecha),
      monto: aNumero_(f.monto),
      desc: aTexto_(f.detalle),
      cia: aTexto_(f.compania),
      cat: aTexto_(f.categoria),
      medio: aTexto_(f.medio) || 'efectivo',
      tarjetaId: aTexto_(f.tarjetaId),
      cuotas: Math.max(1, aNumero_(f.cuotas) || 1),
      recurrente: /^(si|sí|true|verdadero|x|1)$/i.test(aTexto_(f.recurrente)),
      mod: base.mod
    };
  }
  if (clave === 'tarjetas') {
    return {
      id: aTexto_(f.id) || ('sheet-' + numeroFila),
      nombre: aTexto_(f.nombre),
      banco: aTexto_(f.banco),
      cierre: aNumero_(f.cierre) || 25,
      vence: aNumero_(f.vence) || 10,
      limite: aNumero_(f.limite) || '',
      mod: base.mod
    };
  }
  if (clave === 'inversiones') {
    return {
      id: aTexto_(f.id) || ('sheet-' + numeroFila),
      nombre: aTexto_(f.nombre),
      tipo: aTexto_(f.tipo) || 'otro',
      donde: aTexto_(f.donde),
      monto: aNumero_(f.monto),
      tna: aNumero_(f.tna),
      fecha: aTexto_(f.fecha),
      vence: aTexto_(f.vence),
      notas: aTexto_(f.notas),
      mod: base.mod
    };
  }
  if (clave === 'objetivos') {
    var aportes = [];
    var crudo = aTexto_(f.aportes);
    if (crudo) {
      // formato: 2026-08-01:200000 | 2026-09-01:150000
      var partes = crudo.split('|');
      for (var i = 0; i < partes.length; i++) {
        var p = partes[i].trim();
        if (!p) continue;
        var corte = p.lastIndexOf(':');
        if (corte > 0) {
          aportes.push({ fecha: p.slice(0, corte).trim(), monto: aNumero_(p.slice(corte + 1)) });
        }
      }
    }
    return {
      id: aTexto_(f.id) || ('sheet-' + numeroFila),
      nombre: aTexto_(f.nombre),
      meta: aNumero_(f.meta),
      fecha: aTexto_(f.fecha),
      aportes: aportes,
      mod: base.mod
    };
  }
  // borrados
  return { id: aTexto_(f.id), tabla: aTexto_(f.tabla), mod: base.mod };
}

/* Convierte un registro de la app a fila de planilla. */
function aFila_(clave, r) {
  if (clave === 'movs') {
    return [r.id, r.tipo, r.fecha, r.monto, r.desc || '', r.cia || '', r.cat || '',
            r.medio || '', r.tarjetaId || '', r.cuotas || 1, r.recurrente ? 'si' : '', r.mod || 1];
  }
  if (clave === 'tarjetas') {
    return [r.id, r.nombre || '', r.banco || '', r.cierre || 25, r.vence || 10, r.limite || '', r.mod || 1];
  }
  if (clave === 'inversiones') {
    return [r.id, r.nombre || '', r.tipo || '', r.donde || '', r.monto || 0, r.tna || 0,
            r.fecha || '', r.vence || '', r.notas || '', r.mod || 1];
  }
  if (clave === 'objetivos') {
    var ap = (r.aportes || []).map(function (a) { return a.fecha + ':' + a.monto; }).join(' | ');
    return [r.id, r.nombre || '', r.meta || 0, r.fecha || '', ap, r.mod || 1];
  }
  return [r.id, r.tabla || '', r.mod || 1];
}

function escribirHoja_(clave, filas) {
  var def = HOJAS[clave];
  var h = hoja_(clave);
  var ultima = h.getLastRow();
  if (ultima > 1) h.getRange(2, 1, ultima - 1, def.cols.length).clearContent();
  if (!filas.length) return;
  var datos = filas.map(function (r) { return aFila_(clave, r); });
  h.getRange(2, 1, datos.length, def.cols.length).setValues(datos);
}

/* ---------- fusión ---------- */

function fusionar_(locales, remotos, borrados, tabla) {
  var porId = {};
  var i, r;

  for (i = 0; i < remotos.length; i++) {
    r = remotos[i];
    if (!r.id) continue;
    porId[r.id] = r;
  }
  for (i = 0; i < locales.length; i++) {
    r = locales[i];
    if (!r.id) continue;
    var actual = porId[r.id];
    // gana el más nuevo
    if (!actual || (r.mod || 1) > (actual.mod || 1)) porId[r.id] = r;
  }
  // aplicar borrados
  for (i = 0; i < borrados.length; i++) {
    var b = borrados[i];
    if (b.tabla !== tabla) continue;
    var reg = porId[b.id];
    if (reg && (b.mod || 1) >= (reg.mod || 1)) delete porId[b.id];
  }

  var salida = [];
  for (var k in porId) if (porId.hasOwnProperty(k)) salida.push(porId[k]);
  return salida;
}

/* ---------- puntos de entrada ---------- */

function responder_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var datos = {
      ok: true,
      movs: leerHoja_('movs'),
      tarjetas: leerHoja_('tarjetas'),
      inversiones: leerHoja_('inversiones'),
      objetivos: leerHoja_('objetivos'),
      borrados: leerHoja_('borrados'),
      servidor: new Date().getTime()
    };
    return responder_(datos);
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return responder_({ ok: false, error: 'La planilla está ocupada, probá de nuevo.' });
  }
  try {
    var entrada = JSON.parse(e.postData.contents);
    var locales = entrada.datos || {};

    var borradosRemotos = leerHoja_('borrados');
    var borradosLocales = (locales.borrados || []).map(function (b) {
      return { id: b.id, tabla: b.tabla, mod: b.mod || 1 };
    });
    var borrados = fusionarBorrados_(borradosLocales, borradosRemotos);

    var resultado = { ok: true, borrados: borrados, servidor: new Date().getTime() };
    var tablas = ['movs', 'tarjetas', 'inversiones', 'objetivos'];

    /* Lo que anotó a mano en la hoja de carga entra como un movimiento más. */
    var cargas = leerCargasNuevas_();

    for (var i = 0; i < tablas.length; i++) {
      var t = tablas[i];
      var desdeHoja = leerHoja_(t);
      if (t === 'movs' && cargas.movs.length) desdeHoja = desdeHoja.concat(cargas.movs);
      var unidos = fusionar_(locales[t] || [], desdeHoja, borrados, t);
      escribirHoja_(t, unidos);
      resultado[t] = unidos;
    }
    escribirHoja_('borrados', borrados);

    /* Recién acá, con todo guardado, marco las filas como cargadas. */
    marcarCargas_(cargas.filas, '✓ cargado');
    resultado.cargadas = cargas.movs.length;

    /* Y reescribo el cuadro del año para poder mirarlo de un vistazo. */
    try {
      var anio = new Date().getFullYear();
      escribirResumen_(resultado.movs, resultado.tarjetas, anio);
      resultado.resumen = anio;
    } catch (errResumen) {
      resultado.avisoResumen = String(errResumen);   // el resumen no puede romper la sincronización
    }

    return responder_(resultado);
  } catch (err) {
    return responder_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function fusionarBorrados_(a, b) {
  var porId = {};
  var todos = a.concat(b);
  for (var i = 0; i < todos.length; i++) {
    var x = todos[i];
    if (!x.id || !x.tabla) continue;
    var clave = x.tabla + '::' + x.id;
    if (!porId[clave] || (x.mod || 1) > (porId[clave].mod || 1)) porId[clave] = x;
  }
  var salida = [];
  for (var k in porId) if (porId.hasOwnProperty(k)) salida.push(porId[k]);
  return salida;
}

/* Corré esta función una vez desde el editor para crear las hojas. */
function prepararPlanilla() {
  ['movs', 'tarjetas', 'inversiones', 'objetivos', 'borrados'].forEach(function (k) { hoja_(k); });
  hojaCargar_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Hojas listas.', 'Plata', 5);
}

/* Para rehacer el cuadro a mano, sin esperar a que sincronice el celular. */
function rehacerResumen() {
  var anio = new Date().getFullYear();
  var h = escribirResumen_(leerHoja_('movs'), leerHoja_('tarjetas'), anio);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(h);
  SpreadsheetApp.getActiveSpreadsheet().toast('Cuadro de ' + anio + ' actualizado.', 'Plata', 5);
}
/* ============================================================
   Plata · Cargar — para anotar desde la compu
   ============================================================
   Una fila por gasto o ingreso. Escribís y la próxima vez que
   la app sincronice, entra como un movimiento más y la fila
   queda marcada. No hace falta tocar nada más.
*/

var CARGAR = {
  nombre: 'Plata · Cargar',
  cols: ['Fecha', 'Tipo', 'Detalle', 'Categoría', 'Monto', 'Medio', 'Cuotas', 'Es fijo', 'Estado']
};

function hojaCargar_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var h = libro.getSheetByName(CARGAR.nombre);
  if (h) return h;

  h = libro.insertSheet(CARGAR.nombre);   // al final, para no moverle las pestañas
  h.getRange(1, 1, 1, CARGAR.cols.length).setValues([CARGAR.cols])
    .setFontWeight('bold').setBackground('#1F3A2E').setFontColor('#FFFFFF');
  h.setFrozenRows(1);

  var anchos = [95, 85, 230, 150, 120, 110, 75, 75, 120];
  for (var i = 0; i < anchos.length; i++) h.setColumnWidth(i + 1, anchos[i]);

  h.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
  h.getRange('E2:E').setNumberFormat('$#,##0.00');

  /* Listas desplegables, para no tener que acordarse de los nombres. */
  ponerLista_(h, 'B2:B', ['gasto', 'ingreso']);
  ponerLista_(h, 'F2:F', ['efectivo', 'débito', 'transferencia', 'tarjeta']);
  ponerLista_(h, 'H2:H', ['sí', 'no']);

  h.getRange('A2:I2').setValues([[
    new Date(), 'gasto', 'Ejemplo: borrá esta fila', 'Supermercado', 15000, 'efectivo', 1, 'no', ''
  ]]);
  h.getRange('A2:I2').setFontColor('#9AA69F').setFontStyle('italic');

  return h;
}

function ponerLista_(hoja, rango, opciones) {
  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true).setAllowInvalid(true).build();
  hoja.getRange(rango).setDataValidation(regla);
}

/* Lee las filas sin marcar y las devuelve como movimientos nuevos.
   Las marca recién cuando la sincronización terminó bien. */
function leerCargasNuevas_() {
  var h = hojaCargar_();
  var ultima = h.getLastRow();
  if (ultima < 2) return { movs: [], filas: [] };

  var valores = h.getRange(2, 1, ultima - 1, CARGAR.cols.length).getValues();
  var movs = [];
  var filas = [];

  for (var i = 0; i < valores.length; i++) {
    var v = valores[i];
    var estado = aTexto_(v[8]);
    var monto = aNumero_(v[4]);
    var detalle = aTexto_(v[2]);

    if (estado !== '') continue;                       // ya entró antes
    if (monto <= 0) continue;                          // fila vacía o sin plata
    if (detalle.indexOf('Ejemplo:') === 0) continue;   // la fila de muestra

    var fecha = aFechaISO_(v[0]);
    if (!fecha) { marcarFila_(h, i + 2, '✗ falta la fecha'); continue; }

    var medio = normalizarMedio_(aTexto_(v[5]));
    var cuotas = Math.max(1, aNumero_(v[6]) || 1);

    movs.push({
      id: 'hoja-' + fecha.replace(/-/g, '') + '-' + (i + 2) + '-' + Math.floor(monto),
      tipo: aTexto_(v[1]).toLowerCase().indexOf('ingres') === 0 ? 'ingreso' : 'gasto',
      fecha: fecha,
      monto: monto,
      desc: detalle,
      cia: '',
      cat: aTexto_(v[3]),
      medio: medio,
      tarjetaId: '',
      cuotas: medio === 'tarjeta' ? cuotas : 1,
      recurrente: /^(s[ií]|true|x|1)$/i.test(aTexto_(v[7])),
      mod: new Date().getTime()
    });
    filas.push(i + 2);
  }
  return { movs: movs, filas: filas };
}

function marcarCargas_(filas, texto) {
  if (!filas.length) return;
  var h = hojaCargar_();
  for (var i = 0; i < filas.length; i++) marcarFila_(h, filas[i], texto);
}

function marcarFila_(hoja, fila, texto) {
  hoja.getRange(fila, 9).setValue(texto);
  hoja.getRange(fila, 1, 1, 9).setFontColor(texto.indexOf('✗') === 0 ? '#A03232' : '#9AA69F');
}

/* La app guarda las fechas como texto yyyy-mm-dd; la planilla, como fecha. */
function aFechaISO_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = aTexto_(v);
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);   // 05/11/2026
  if (m) {
    var anio = m[3].length === 2 ? '20' + m[3] : m[3];
    return anio + '-' + dosDig_(m[2]) + '-' + dosDig_(m[1]);
  }
  return '';
}

function dosDig_(n) { return ('0' + String(n)).slice(-2); }

function normalizarMedio_(s) {
  var t = s.toLowerCase();
  if (t.indexOf('tarj') === 0 || t.indexOf('créd') === 0 || t.indexOf('cred') === 0) return 'tarjeta';
  if (t.indexOf('transf') === 0) return 'transferencia';
  if (t.indexOf('déb') === 0 || t.indexOf('deb') === 0) return 'debito';
  return 'efectivo';
}
/* ============================================================
   Plata · Resumen — el cuadro del año, llenado solo
   ============================================================
   Mismo formato que venís usando: los meses en columnas y los
   conceptos en filas. Lo reescribe la app en cada sincronización,
   así que no conviene escribir a mano acá: para anotar algo está
   la hoja "Plata · Cargar".
*/

var MESES_GS = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

/* Las mismas categorías que usa la app: lo que vuelve todos los meses
   va por un lado y el día a día por otro. */
var CAT_FIJOS_GS = ['Celular', 'Internet', 'Empleada', 'Almacén', 'Luz', 'Gas',
                    'Tarjeta Naranja', 'Tarjeta BBVA', 'Tarjeta Galicia', 'Nafta',
                    'Crédito casa', 'Seguro de retiro', 'Deportes', 'Juan Pablo',
                    'Impuestos'];
var CAT_VARIOS_GS = ['Salidas', 'Café', 'Otros'];

/* Lo cargado antes traía otras categorías. Las traduzco al vuelo para que
   el cuadro se lea parejo, sin tocar los movimientos ya guardados. */
var EQUIVALENCIAS = {
  'Combustible': 'Nafta',
  'Supermercado': 'Almacén',
  'Comida y salidas': 'Salidas',
  'Ocio': 'Salidas',
  'Servicios': 'Otros',
  'Salud': 'Otros',
  'Auto': 'Otros',
  'Tecnología': 'Otros',
  'Ropa': 'Otros',
  'Educación': 'Otros'
};

function categoriaDe_(mv) {
  var c = mv.cat || '';
  if (EQUIVALENCIAS[c]) return EQUIVALENCIAS[c];
  return c || 'Otros';
}

function esFijaGS_(cat) {
  for (var i = 0; i < CAT_FIJOS_GS.length; i++) {
    if (CAT_FIJOS_GS[i] === cat) return true;
  }
  return false;
}

var VERDE = '#1F3A2E', ARENA = '#F4F1EA', LINEA = '#D8D4C8', GRIS = '#9AA69F';

/* Además de la planilla de datos, el cuadro se copia a la planilla de siempre,
   la que mirás todos los días. Siempre en una hoja nueva: no toca las tuyas.
   Dejalo en '' si alguna vez querés que escriba solo acá. */
var PLANILLA_DE_SIEMPRE = '13pVUsT9CfNvuDGu8bmsL_mg8dwtAQ29nnpfKu5TK0g4';

/* Reparte un gasto en cuotas sobre los meses en que realmente vence. */
function cuotasDelMov_(mov, tarjetas) {
  var n = Math.max(1, Number(mov.cuotas) || 1);
  var partes = mov.fecha.split('-');
  var Y = Number(partes[0]), M = Number(partes[1]), D = Number(partes[2]);
  if (!Y || !M) return [];

  if (mov.medio !== 'tarjeta' || n === 1) {
    return [{ y: Y, m: M - 1, monto: Number(mov.monto) || 0 }];
  }

  var cierre = 25;
  for (var i = 0; i < tarjetas.length; i++) {
    if (tarjetas[i].id === mov.tarjetaId) { cierre = Number(tarjetas[i].cierre) || 25; break; }
  }

  var salto = D <= cierre ? 1 : 2;
  var valor = (Number(mov.monto) || 0) / n;
  var salida = [];
  for (var k = 0; k < n; k++) {
    var total = (M - 1) + salto + k;
    salida.push({ y: Y + Math.floor(total / 12), m: ((total % 12) + 12) % 12, monto: valor });
  }
  return salida;
}

/* Arma el cuadro del año a partir de los movimientos. */
function armarResumen_(movs, tarjetas, anio) {
  var vacio = function () { var a = []; for (var i = 0; i < 12; i++) a.push(0); return a; };

  var ingresos = vacio(), fijos = vacio(), varios = vacio(), tarjeta = vacio();
  var porFijo = {}, porVario = {}, porCia = {};

  for (var i = 0; i < movs.length; i++) {
    var mv = movs[i];
    if (!mv.fecha) continue;

    var trozos = cuotasDelMov_(mv, tarjetas);
    for (var t = 0; t < trozos.length; t++) {
      var tr = trozos[t];
      if (tr.y !== anio) continue;
      var m = tr.m, monto = tr.monto;

      if (mv.tipo === 'ingreso') {
        ingresos[m] += monto;
        var cia = mv.cia || 'Sin compañía';
        if (!porCia[cia]) porCia[cia] = vacio();
        porCia[cia][m] += monto;
        continue;
      }

      /* Siempre por categoría: el detalle de cada compra se mira en la app.
         Cómo se pagó no cambia de qué es el gasto. */
      var cat = categoriaDe_(mv);
      if (mv.medio === 'tarjeta') tarjeta[m] += monto;   // dato aparte, no un balde más

      if (esFijaGS_(cat)) {
        fijos[m] += monto;
        if (!porFijo[cat]) porFijo[cat] = vacio();
        porFijo[cat][m] += monto;
      } else if (mv.recurrente) {
        /* Marcado como fijo pero sin una categoría de la lista. */
        fijos[m] += monto;
        if (!porFijo['Otros fijos']) porFijo['Otros fijos'] = vacio();
        porFijo['Otros fijos'][m] += monto;
      } else {
        varios[m] += monto;
        if (!porVario[cat]) porVario[cat] = vacio();
        porVario[cat][m] += monto;
      }
    }
  }

  return {
    ingresos: ingresos, fijos: fijos, varios: varios, tarjeta: tarjeta,
    porFijo: porFijo, porVario: porVario, porCia: porCia
  };
}

/* Las filas del cuadro: siempre las mismas y en el mismo orden, aunque un mes
   dé cero. Así la estructura del año se lee de un vistazo. Lo que aparezca
   fuera de la lista se agrega al final, para no esconder nada. */
function filasEnOrden_(mapa, orden) {
  var vacio = function () { var a = []; for (var i = 0; i < 12; i++) a.push(0); return a; };
  var filas = [];
  var puestas = {};

  for (var i = 0; i < orden.length; i++) {
    var nombre = orden[i];
    var meses = mapa[nombre] || vacio();
    var total = 0;
    for (var m = 0; m < 12; m++) total += meses[m];
    filas.push({ nombre: nombre, meses: meses, total: total });
    puestas[nombre] = true;
  }

  var sueltas = [];
  for (var k in mapa) {
    if (!mapa.hasOwnProperty(k) || puestas[k]) continue;
    var t = 0;
    for (var j = 0; j < 12; j++) t += mapa[k][j];
    if (t > 0) sueltas.push({ nombre: k, meses: mapa[k], total: t });
  }
  sueltas.sort(function (a, b) { return b.total - a.total; });

  return filas.concat(sueltas);
}

/* Ordena los conceptos por cuánto pesan en el año. */
function ordenarPorPeso_(mapa) {
  var filas = [];
  for (var k in mapa) {
    if (!mapa.hasOwnProperty(k)) continue;
    var total = 0;
    for (var i = 0; i < 12; i++) total += mapa[k][i];
    if (total > 0) filas.push({ nombre: k, meses: mapa[k], total: total });
  }
  filas.sort(function (a, b) { return b.total - a.total; });
  return filas;
}

function escribirResumen_(movs, tarjetas, anio) {
  var h = resumenEnLibro_(SpreadsheetApp.getActiveSpreadsheet(), movs, tarjetas, anio);

  /* La copia en la planilla de siempre no puede voltear la sincronización. */
  if (PLANILLA_DE_SIEMPRE) {
    try {
      resumenEnLibro_(SpreadsheetApp.openById(PLANILLA_DE_SIEMPRE), movs, tarjetas, anio);
    } catch (err) {
      Logger.log('No pude copiar el cuadro a la planilla de siempre: ' + err);
    }
  }
  return h;
}

function resumenEnLibro_(libro, movs, tarjetas, anio) {
  var nombre = 'Plata · Resumen ' + anio;
  var h = libro.getSheetByName(nombre);
  if (!h) h = libro.insertSheet(nombre);
  h.clear();
  h.clearConditionalFormatRules();

  var r = armarResumen_(movs, tarjetas, anio);
  var filas = [];

  /* --- Encabezado --- */
  filas.push(['Plata · ' + anio, '', '', '', '', '', '', '', '', '', '', '', '', '']);
  filas.push(['Se llena solo con lo que cargás en la app o en la hoja "Plata · Cargar".',
              '', '', '', '', '', '', '', '', '', '', '', '', '']);
  filas.push(new Array(14).join(',').split(','));

  /* --- El año mes por mes --- */
  var cab = ['Mes', 'Ingresos', 'Gastos fijos', 'Gastos varios', 'Total gastos',
             'Resultado', 'Quedó acumulado', 'De eso, con tarjeta'];
  while (cab.length < 14) cab.push('');
  filas.push(cab);

  var acumulado = 0;
  var totIng = 0, totFij = 0, totVar = 0, totTar = 0;
  for (var m = 0; m < 12; m++) {
    var gastos = r.fijos[m] + r.varios[m];
    var resultado = r.ingresos[m] - gastos;
    acumulado += resultado;
    totIng += r.ingresos[m]; totFij += r.fijos[m];
    totVar += r.varios[m]; totTar += r.tarjeta[m];

    var fila = [MESES_GS[m], r.ingresos[m], r.fijos[m], r.varios[m],
                gastos, resultado, acumulado, r.tarjeta[m]];
    while (fila.length < 14) fila.push('');
    filas.push(fila);
  }

  var totGas = totFij + totVar;
  var filaTot = ['Total del año', totIng, totFij, totVar, totGas, totIng - totGas, '', totTar];
  while (filaTot.length < 14) filaTot.push('');
  filas.push(filaTot);

  /* --- Los bloques de detalle --- */
  var marcas = { anioDesde: 5, anioHasta: 16, total: 17 };
  var cursor = filas.length + 1;   // 1 fila en blanco entre bloques

  cursor = agregarBloque_(filas, 'Gastos fijos', filasEnOrden_(r.porFijo, CAT_FIJOS_GS), marcas, 'fijos');
  cursor = agregarBloque_(filas, 'Gastos varios', filasEnOrden_(r.porVario, CAT_VARIOS_GS), marcas, 'varios');
  cursor = agregarBloque_(filas, 'Ingresos por compañía', ordenarPorPeso_(r.porCia), marcas, 'cias');

  h.getRange(1, 1, filas.length, 14).setValues(filas);
  pintarResumen_(h, filas.length, marcas);
  return h;
}

/* Un bloque: título con los meses, una fila por concepto, y el sub-total. */
function agregarBloque_(filas, titulo, items, marcas, clave) {
  filas.push(new Array(14).join(',').split(','));

  var cab = [titulo].concat(MESES_GS).concat(['Total']);
  filas.push(cab);
  marcas[clave + 'Cab'] = filas.length;

  if (!items.length) {
    var nada = ['Todavía no hay nada acá'];
    while (nada.length < 14) nada.push('');
    filas.push(nada);
    marcas[clave + 'Desde'] = filas.length;
    marcas[clave + 'Hasta'] = filas.length;
    return filas.length;
  }

  marcas[clave + 'Desde'] = filas.length + 1;
  var sub = [];
  for (var i = 0; i < 12; i++) sub.push(0);

  for (var k = 0; k < items.length; k++) {
    var it = items[k];
    var fila = [it.nombre];
    for (var m = 0; m < 12; m++) { fila.push(it.meses[m]); sub[m] += it.meses[m]; }
    fila.push(it.total);
    filas.push(fila);
  }
  marcas[clave + 'Hasta'] = filas.length;

  var filaSub = ['Sub-total'];
  var totalSub = 0;
  for (var m2 = 0; m2 < 12; m2++) { filaSub.push(sub[m2]); totalSub += sub[m2]; }
  filaSub.push(totalSub);
  filas.push(filaSub);
  marcas[clave + 'Sub'] = filas.length;

  return filas.length;
}

function pintarResumen_(h, largo, marcas) {
  h.setFrozenColumns(1);
  h.setColumnWidth(1, 210);
  for (var c = 2; c <= 14; c++) h.setColumnWidth(c, 112);

  h.getRange(1, 1).setFontSize(15).setFontWeight('bold').setFontColor(VERDE);
  h.getRange(2, 1, 1, 6).setFontSize(10).setFontColor(GRIS);

  /* La tabla del año */
  h.getRange(4, 1, 1, 8).setFontWeight('bold').setBackground(VERDE).setFontColor('#FFFFFF');
  h.getRange(5, 2, 13, 7).setNumberFormat('$#,##0');
  h.getRange(17, 1, 1, 8).setFontWeight('bold').setBackground(ARENA)
    .setBorder(true, null, null, null, null, null, LINEA, SpreadsheetApp.BorderStyle.SOLID);

  /* Resultado en rojo cuando da negativo */
  var rojo = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setFontColor('#A03232')
    .setRanges([h.getRange(5, 7, 13, 2)]).build();
  h.setConditionalFormatRules([rojo]);

  /* Los bloques de detalle */
  var bloques = ['fijos', 'varios', 'cias'];
  for (var i = 0; i < bloques.length; i++) {
    var b = bloques[i];
    var cab = marcas[b + 'Cab'], desde = marcas[b + 'Desde'], sub = marcas[b + 'Sub'];
    if (!cab) continue;
    h.getRange(cab, 1, 1, 14).setFontWeight('bold').setBackground(VERDE).setFontColor('#FFFFFF');
    if (desde) {
      var alto = (marcas[b + 'Hasta'] || desde) - desde + 1;
      h.getRange(desde, 2, alto, 13).setNumberFormat('$#,##0');
    }
    if (sub) {
      h.getRange(sub, 1, 1, 14).setFontWeight('bold').setBackground(ARENA).setNumberFormat('$#,##0');
      h.getRange(sub, 1).setNumberFormat('@');
    }
  }
}

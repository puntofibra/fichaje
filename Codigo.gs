/**
 * FICHAJE NFC — backend Apps Script (API JSONP)
 * Etiqueta NTAG216 pegada en el local. La URL grabada lleva:
 *   https://puntofibra.github.io/fichaje/?p=LOCAL01&k=SECRETO&c=000000
 * donde c es el ASCII mirror del contador NFC (hex, 6 dígitos) que la
 * propia etiqueta sustituye en cada lectura. El contador demuestra que la
 * etiqueta física se leyó y en qué orden: una URL copiada o un tag clonado
 * se caen en cuanto alguien ficha después.
 *
 * Frontend: https://puntofibra.github.io/fichaje/  (GitHub Pages)
 * Hoja:     fichar nfc
 */

var SHEET_ID   = '10NUxnmK3zNddlp_2ztnLxP0E751I07DZxj5ymPw3xbM';
var TZ         = 'Europe/Madrid';
var MAX_SALTO  = 100;   // salto máximo tolerado en el contador NFC
var VERSION    = '1.0';

var H_PUNTOS    = ['punto_id', 'nombre', 'secreto', 'ultimo_contador', 'ultima_lectura', 'activo'];
var H_EMPLEADOS = ['emp_id', 'nombre', 'codigo_alta', 'token', 'alta_fecha', 'dispositivo', 'activo'];
var H_FICHAJES  = ['id', 'fecha_hora', 'fecha', 'hora', 'emp_id', 'nombre', 'tipo', 'punto_id', 'contador', 'geo', 'precision_m', 'dispositivo'];
var H_CONFIG    = ['clave', 'valor', 'descripcion'];

/* ───────────────────────── ENTRADA HTTP ───────────────────────── */

function doGet(e) {
  var p  = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  var out;
  try {
    out = enrutar(p);
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) { return doGet(e); }

function enrutar(p) {
  switch (String(p.a || '')) {
    case 'ping':   return { ok: true, version: VERSION, hora: ahoraStr() };
    case 'alta':   return apiAlta(p);
    case 'fichar': return apiFichar(p);
    case 'estado': return apiEstado(p);
    case 'panel':  return apiPanel(p);
    default:       return { ok: false, error: 'Acción desconocida' };
  }
}

/* ───────────────────────── API ───────────────────────── */

/** Alta del móvil del empleado: canjea el código de la hoja por un token. */
function apiAlta(p) {
  var punto = validarPunto(p);            // exige estar en el local (p + k)
  if (!punto.ok) return punto;

  var codigo = String(p.codigo || '').trim().toUpperCase();
  if (!codigo) return { ok: false, error: 'Falta el código de alta' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh   = hoja('Empleados');
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][2]).trim().toUpperCase() !== codigo) continue;
      if (String(vals[i][6]).trim().toUpperCase() === 'NO') return { ok: false, error: 'Empleado dado de baja' };

      var yaTenia = String(vals[i][3] || '').trim();
      if (yaTenia && conf('permitir_realta', 'SI').toUpperCase() !== 'SI') {
        return { ok: false, error: 'Este código ya está en uso en otro móvil' };
      }
      var empId = String(vals[i][0] || '').trim() || nuevoEmpId(vals);
      var token = Utilities.getUuid();
      sh.getRange(i + 1, 1).setValue(empId);
      sh.getRange(i + 1, 4).setValue(token);
      sh.getRange(i + 1, 5).setValue(new Date());
      sh.getRange(i + 1, 6).setValue(String(p.dev || '').slice(0, 120));
      return { ok: true, token: token, emp_id: empId, nombre: String(vals[i][1]), realta: !!yaTenia };
    }
    return { ok: false, error: 'Código no válido' };
  } finally {
    lock.releaseLock();
  }
}

/** Fichaje: valida punto + secreto + contador NFC y escribe la fila. */
function apiFichar(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var emp = buscarEmpleadoPorToken(p.t);
    if (!emp) return { ok: false, error: 'Móvil no dado de alta', realta: true };

    var punto = validarPunto(p);
    if (!punto.ok) return punto;

    var cnt = leerContador(p);
    var chk = validarContador(punto.fila, cnt);
    if (!chk.ok) return chk;

    var tipo = String(p.tipo || '').toLowerCase();
    var ult  = ultimoFichajeDe(emp.emp_id);
    if (tipo !== 'entrada' && tipo !== 'salida') {
      tipo = (!ult || ult.tipo === 'salida') ? 'entrada' : 'salida';
    }

    var margen = Number(conf('minutos_antirrebote', '1'));
    if (ult && margen > 0 && (new Date() - ult.fecha_hora) / 60000 < margen) {
      return { ok: false, error: 'Acabas de fichar hace un momento (' + ult.tipo + ' a las ' + ult.hora + ')' };
    }

    var ahora = new Date();
    var fila  = [
      Utilities.getUuid().slice(0, 8),
      ahora,
      Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(ahora, TZ, 'HH:mm:ss'),
      emp.emp_id,
      emp.nombre,
      tipo,
      punto.punto_id,
      cnt,
      String(p.geo || ''),
      String(p.acc || ''),
      String(p.dev || '').slice(0, 120)
    ];
    hoja('Fichajes').appendRow(fila);
    guardarContador(punto.fila, cnt, ahora);

    return {
      ok: true,
      tipo: tipo,
      nombre: emp.nombre,
      punto: punto.nombre,
      hora: Utilities.formatDate(ahora, TZ, 'HH:mm'),
      fecha: Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy'),
      contador: cnt,
      horas_hoy: horasDelDia(emp.emp_id, Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd'))
    };
  } finally {
    lock.releaseLock();
  }
}

/** Estado del empleado (sin tocar la etiqueta): dentro/fuera y horas de hoy. */
function apiEstado(p) {
  var emp = buscarEmpleadoPorToken(p.t);
  if (!emp) return { ok: false, error: 'Móvil no dado de alta', realta: true };
  var hoy = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var ult = ultimoFichajeDe(emp.emp_id);
  return {
    ok: true,
    nombre: emp.nombre,
    emp_id: emp.emp_id,
    dentro: !!(ult && ult.tipo === 'entrada'),
    ultimo: ult ? { tipo: ult.tipo, hora: ult.hora, fecha: ult.fecha } : null,
    horas_hoy: horasDelDia(emp.emp_id, hoy),
    hoy: listarDia(emp.emp_id, hoy)
  };
}

/** Panel de control (hash SHA-256 de la contraseña en la query, nunca la contraseña). */
function apiPanel(p) {
  if (String(p.pw || '').toLowerCase() !== conf('admin_hash', '').toLowerCase() || !conf('admin_hash', '')) {
    Utilities.sleep(1200);
    return { ok: false, error: 'Contraseña incorrecta' };
  }
  var desde = String(p.desde || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-01'));
  var hasta = String(p.hasta || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));

  var vals = hoja('Fichajes').getDataRange().getValues();
  var filas = [], resumen = {}, estado = {};
  for (var i = 1; i < vals.length; i++) {
    var f = String(vals[i][2]);
    var empId = String(vals[i][4]), nombre = String(vals[i][5]), tipo = String(vals[i][6]);
    estado[empId] = { nombre: nombre, tipo: tipo, hora: String(vals[i][3]), fecha: f };
    if (f < desde || f > hasta) continue;
    filas.push({
      fecha: f, hora: String(vals[i][3]), emp_id: empId, nombre: nombre,
      tipo: tipo, punto: String(vals[i][7]), contador: vals[i][8], geo: String(vals[i][9])
    });
  }
  // horas por empleado y día
  var abierto = {};
  filas.forEach(function (r) {
    var clave = r.emp_id + '|' + r.fecha;
    if (r.tipo === 'entrada') { abierto[clave] = r.hora; return; }
    if (abierto[clave]) {
      var min = minutos(r.hora) - minutos(abierto[clave]);
      if (min > 0) {
        resumen[r.emp_id] = resumen[r.emp_id] || { nombre: r.nombre, minutos: 0, dias: {} };
        resumen[r.emp_id].minutos += min;
        resumen[r.emp_id].dias[r.fecha] = (resumen[r.emp_id].dias[r.fecha] || 0) + min;
      }
      delete abierto[clave];
    }
  });

  var dentro = [];
  Object.keys(estado).forEach(function (k) {
    if (estado[k].tipo === 'entrada') dentro.push(estado[k]);
  });

  return { ok: true, desde: desde, hasta: hasta, filas: filas, resumen: resumen, dentro: dentro };
}

/* ───────────────────────── VALIDACIÓN ───────────────────────── */

function validarPunto(p) {
  var id = String(p.p || '').trim().toUpperCase();
  var k  = String(p.k || '').trim();
  if (!id || !k) return { ok: false, error: 'Faltan datos de la etiqueta. Acerca el móvil al tag del local.' };

  var vals = hoja('Puntos').getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() !== id) continue;
    if (String(vals[i][5]).trim().toUpperCase() === 'NO') return { ok: false, error: 'Punto de fichaje desactivado' };
    if (String(vals[i][2]).trim() !== k) return { ok: false, error: 'Etiqueta no válida' };
    return { ok: true, fila: i + 1, punto_id: String(vals[i][0]).trim(), nombre: String(vals[i][1]), ultimo: Number(vals[i][3] || 0) };
  }
  return { ok: false, error: 'Punto de fichaje desconocido' };
}

/** El mirror llega como 6 hex ASCII; toleramos que TagWriter lo coloque en otro parámetro. */
function leerContador(p) {
  var candidatos = [p.c, p.cnt, p.x];
  for (var k in p) if (candidatos.indexOf(p[k]) === -1) candidatos.push(p[k]);
  for (var i = 0; i < candidatos.length; i++) {
    var v = String(candidatos[i] == null ? '' : candidatos[i]).trim();
    if (/^[0-9A-Fa-f]{6}$/.test(v)) return parseInt(v, 16);
    var m = v.match(/[0-9A-Fa-f]{6}/);
    if (i < 3 && m) return parseInt(m[0], 16);
  }
  return -1;
}

function validarContador(fila, cnt) {
  if (cnt < 0) return { ok: false, error: 'La etiqueta no envió el contador NFC. Revisa el mirror en TagWriter.' };
  var sh  = hoja('Puntos');
  var ult = Number(sh.getRange(fila, 4).getValue() || 0);
  if (!ult) return { ok: true };
  if (cnt <= ult) return { ok: false, error: 'Lectura repetida o etiqueta clonada (contador ' + cnt + ' ≤ ' + ult + ')' };
  if (cnt - ult > MAX_SALTO) return { ok: false, error: 'Salto de contador anómalo (' + ult + ' → ' + cnt + ')' };
  return { ok: true };
}

function guardarContador(fila, cnt, ahora) {
  var sh = hoja('Puntos');
  sh.getRange(fila, 4).setValue(cnt);
  sh.getRange(fila, 5).setValue(ahora);
}

/* ───────────────────────── DATOS ───────────────────────── */

function libro() { return SpreadsheetApp.openById(SHEET_ID); }

function hoja(nombre) {
  var sh = libro().getSheetByName(nombre);
  if (!sh) throw new Error('Falta la hoja "' + nombre + '". Ejecuta setup() una vez.');
  return sh;
}

function conf(clave, porDefecto) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('cfg_' + clave);
  if (hit !== null) return hit;
  var vals = hoja('Config').getDataRange().getValues();
  var val = porDefecto == null ? '' : porDefecto;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === clave) { val = String(vals[i][1]); break; }
  }
  cache.put('cfg_' + clave, val, 60);
  return val;
}

function buscarEmpleadoPorToken(token) {
  token = String(token || '').trim();
  if (!token) return null;
  var vals = hoja('Empleados').getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][3]).trim() !== token) continue;
    if (String(vals[i][6]).trim().toUpperCase() === 'NO') return null;
    return { fila: i + 1, emp_id: String(vals[i][0]), nombre: String(vals[i][1]) };
  }
  return null;
}

function ultimoFichajeDe(empId) {
  var sh = hoja('Fichajes');
  var n  = sh.getLastRow();
  if (n < 2) return null;
  var desde = Math.max(2, n - 400);
  var vals  = sh.getRange(desde, 1, n - desde + 1, H_FICHAJES.length).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][4]) !== String(empId)) continue;
    return { fecha_hora: vals[i][1], fecha: String(vals[i][2]), hora: String(vals[i][3]), tipo: String(vals[i][6]) };
  }
  return null;
}

function listarDia(empId, fecha) {
  var sh = hoja('Fichajes');
  var n  = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(2, 1, n - 1, H_FICHAJES.length).getValues(), out = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][4]) === String(empId) && String(vals[i][2]) === fecha) {
      out.push({ hora: String(vals[i][3]), tipo: String(vals[i][6]) });
    }
  }
  return out;
}

function horasDelDia(empId, fecha) {
  var lista = listarDia(empId, fecha), min = 0, ini = null;
  lista.forEach(function (r) {
    if (r.tipo === 'entrada') { ini = r.hora; return; }
    if (ini) { min += Math.max(0, minutos(r.hora) - minutos(ini)); ini = null; }
  });
  if (ini) {
    var ahora = Utilities.formatDate(new Date(), TZ, 'HH:mm:ss');
    min += Math.max(0, minutos(ahora) - minutos(ini));
  }
  return Math.round(min);
}

function minutos(hhmmss) {
  var p = String(hhmmss).split(':');
  return (Number(p[0]) || 0) * 60 + (Number(p[1]) || 0) + (Number(p[2]) || 0) / 60;
}

function nuevoEmpId(vals) {
  var max = 0;
  for (var i = 1; i < vals.length; i++) {
    var m = String(vals[i][0]).match(/^E(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return 'E' + ('00' + (max + 1)).slice(-3);
}

function ahoraStr() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'); }

function sha256(txt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, txt, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* ───────────────────────── SETUP Y UTILIDADES ───────────────────────── */

/** Ejecutar UNA VEZ. Crea las hojas y el primer punto de fichaje. */
function setup() {
  var ss = libro();
  crearHoja(ss, 'Puntos', H_PUNTOS);
  crearHoja(ss, 'Empleados', H_EMPLEADOS);
  crearHoja(ss, 'Fichajes', H_FICHAJES);
  crearHoja(ss, 'Config', H_CONFIG);

  var cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 4, 3).setValues([
      ['admin_hash', sha256('fichar2026'), 'SHA-256 de la contraseña del panel. Cámbiala con cambiarPassword("nueva")'],
      ['permitir_realta', 'SI', 'SI = el mismo código puede darse de alta en otro móvil (cambio de teléfono)'],
      ['minutos_antirrebote', '1', 'Minutos mínimos entre dos fichajes del mismo empleado'],
      ['empresa', 'FuturMovil', 'Nombre que aparece en la página']
    ]);
  }

  var pu = ss.getSheetByName('Puntos');
  if (pu.getLastRow() < 2) {
    pu.appendRow(['LOCAL01', 'Tienda principal', secretoNuevo(), '', '', 'SI']);
  }

  var em = ss.getSheetByName('Empleados');
  if (em.getLastRow() < 2) {
    em.appendRow(['E001', 'Manu', 'MANU01', '', '', '', 'SI']);
  }

  var fi = ss.getSheetByName('Fichajes');
  fi.setFrozenRows(1);
  fi.getRange('B:B').setNumberFormat('dd/mm/yyyy HH:mm:ss');

  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Hoja1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);

  Logger.log(urlEtiqueta('LOCAL01'));
  return urlEtiqueta('LOCAL01');
}

function crearHoja(ss, nombre, cabecera) {
  var sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cabecera.length).setValues([cabecera])
      .setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function secretoNuevo() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

/** Muestra la URL exacta que hay que grabar en el NTAG216 con TagWriter. */
function urlEtiqueta(puntoId) {
  var vals = hoja('Puntos').getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === String(puntoId).trim().toUpperCase()) {
      var url = 'https://puntofibra.github.io/fichaje/?p=' + vals[i][0] + '&k=' + vals[i][2] + '&c=000000';
      Logger.log('Graba esta URL en el tag y activa el counter mirror sobre los 6 ceros:\n' + url);
      return url;
    }
  }
  throw new Error('Punto no encontrado');
}

/** Añade un punto de fichaje nuevo y devuelve su URL de grabado. */
function nuevoPunto(puntoId, nombre) {
  hoja('Puntos').appendRow([String(puntoId).toUpperCase(), nombre || puntoId, secretoNuevo(), '', '', 'SI']);
  return urlEtiqueta(puntoId);
}

/** Añade un empleado y devuelve su código de alta. */
function nuevoEmpleado(nombre, codigo) {
  var sh   = hoja('Empleados');
  var vals = sh.getDataRange().getValues();
  var cod  = String(codigo || (String(nombre).replace(/[^A-Za-z]/g, '').toUpperCase().slice(0, 4) + Math.floor(10 + Math.random() * 89)));
  sh.appendRow([nuevoEmpId(vals), nombre, cod, '', '', '', 'SI']);
  Logger.log('Código de alta de ' + nombre + ': ' + cod);
  return cod;
}

/** Cambia la contraseña del panel. */
function cambiarPassword(nueva) {
  var sh = hoja('Config'), vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'admin_hash') { sh.getRange(i + 1, 2).setValue(sha256(nueva)); CacheService.getScriptCache().remove('cfg_admin_hash'); return 'Contraseña actualizada'; }
  }
  sh.appendRow(['admin_hash', sha256(nueva), 'SHA-256 de la contraseña del panel']);
  return 'Contraseña creada';
}

/** Desvincula el móvil de un empleado (borra su token) para que vuelva a darse de alta. */
function resetearMovil(empIdONombre) {
  var sh = hoja('Empleados'), vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === empIdONombre || String(vals[i][1]) === empIdONombre) {
      sh.getRange(i + 1, 4, 1, 3).setValues([['', '', '']]);
      return 'Móvil desvinculado de ' + vals[i][1];
    }
  }
  throw new Error('Empleado no encontrado');
}

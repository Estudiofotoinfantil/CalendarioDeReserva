/**
 * Estudio Fotoinfantil · Servidor de reservas v2 (Google Apps Script)
 *
 * Qué cambia respecto del servidor anterior:
 *  - El formulario público solo ve horarios ocupados (sin nombres, teléfonos ni montos) y solo puede crear reservas.
 *  - El panel entra con un código de acceso que controla el servidor (no queda escrito en ninguna página).
 *  - Cada respuesta dice si salió bien o mal: nada se da por guardado sin confirmación.
 *  - Bloqueo para que dos personas no tomen el mismo horario a la vez.
 *  - Seña: la reserva de la mamá se guarda 24 h para transferir; el panel puede asegurar un turno sin seña.
 *  - Todo lo nuevo se guarda en pestañas nuevas (Seguimiento, Registro, Ajustes).
 *    Las reservas existentes NUNCA se reescriben: solo cambia una celda cuando el panel lo pide,
 *    y cada cambio queda anotado en "Registro" con el valor anterior (se puede volver atrás).
 *
 * Configuración privada (Configuración del proyecto → Propiedades del script):
 *   CODIGO_INICIAL  código de acceso del panel. Al primer uso se guarda cifrado y esta propiedad se borra sola.
 *   EMAILS_ADMIN    correos para recuperar el acceso y recibir avisos, separados por coma.
 *   GEMINI_API_KEY  clave de la IA (opcional).
 *   GEMINI_MODEL    modelo de la IA (opcional; por defecto gemini-3.6-flash).
 *   SHEET_ID        id de la planilla (solo si este proyecto NO está dentro de la planilla).
 *
 * Funciones para ejecutar a mano desde el editor:
 *   revisarInstalacion()   revisa todo, crea las pestañas nuevas y muestra un resumen en el registro.
 *   instalarCopiaDiaria()  deja programada una copia de seguridad todos los días (y hace una ahora).
 *   probarMail()           manda un mail de prueba a los correos de EMAILS_ADMIN.
 */

var VERSION = 'v2.0 · 2026-10-23';
var TZ = 'America/Argentina/Buenos_Aires';
var H_RES = 'Reservas', H_CONF = 'Configuracion', H_SEG = 'Seguimiento', H_REG = 'Registro', H_AJ = 'Ajustes';
var ENC_SEG = ['id', 'origen', 'creado', 'venceSena', 'senaRecibida', 'asegurado'];
var ENC_REG = ['fecha', 'accion', 'id', 'antes', 'despues'];
var ENC_AJ = ['clave', 'valor', 'descripcion'];
var CARPETA_RESPALDOS = 'Respaldos Estudio Infantil';
var MAX_COPIAS = 30;
var ESTADOS = ['PENDIENTE', 'CONFIRMADO', 'CANCELADO'];
var TIPOS_BLOQUEO = ['dia', 'slot', 'slot_extra'];
var AJUSTES_BASE = {
  senaMonto:       { valor: '17000', desc: 'Monto de la seña en pesos' },
  senaHoras:       { valor: '24', desc: 'Horas que se guarda una reserva de la web para transferir la seña' },
  alias:           { valor: '', desc: 'Alias para transferir la seña (se muestra a la clienta)' },
  whatsappEstudio: { valor: '541133408139', desc: 'WhatsApp del estudio (con 54, sin + ni espacios)' },
  grillaSemana:    { valor: '18:00,19:00', desc: 'Horarios de lunes a viernes' },
  grillaSabado:    { valor: '10:00,11:00,13:00,14:00', desc: 'Horarios de los sábados' },
  minDias:         { valor: '2', desc: 'Días mínimos de anticipación para reservar desde la web' },
  tiposSesion:     { valor: 'Mes a mes,Temático,Pre añito', desc: 'Tipos de sesión que se pueden reservar desde la web' },
  avisoReservaMail: { valor: 'si', desc: 'Mandar un mail a los correos del panel con cada reserva web (si/no)' }
};
var MAX_DIAS_WEB = 400;   // hasta cuántos días adelante se puede reservar desde la web

// ════════════════════════════════════════════════════════════════════════════
// ENTRADA
// ════════════════════════════════════════════════════════════════════════════
function doGet(e) {
  var a = e && e.parameter ? String(e.parameter.action || '') : '';
  try {
    if (a === 'horarios') return json(apiHorarios());
    if (a === 'ping') return json({ ok: true, version: VERSION });
    return json({ ok: false, error: 'Acción no disponible' });
  } catch (err) {
    return json({ ok: false, error: mensaje(err) });
  }
}

var PUBLICAS = {
  login: apiLogin, reservar: apiReservar, recuperar: apiRecuperar, restablecer: apiRestablecer
};
var PRIVADAS = {
  verificar: function () { return { ok: true, version: VERSION }; },
  getAll: apiGetAll, getConfig: apiGetConfig, saveConfig: apiSaveConfig,
  updateStatus: apiUpdateStatus, updatePago: apiUpdatePago, senaRecibida: apiSenaRecibida,
  reprogramar: apiReprogramar, nuevoTurno: apiNuevoTurno,
  getAjustes: apiGetAjustes, saveAjustes: apiSaveAjustes,
  copiaAhora: apiCopiaAhora, cerrarSesiones: apiCerrarSesiones, logout: apiLogout, cambiarCodigo: apiCambiarCodigo,
  geminiChat: apiGeminiChat, geminiResumen: apiGeminiResumen
};

function doPost(e) {
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var a = String(d.action || '');
    if (PUBLICAS.hasOwnProperty(a)) return json(PUBLICAS[a](d));
    if (PRIVADAS.hasOwnProperty(a)) {
      if (!sesionValida(d.token)) return json({ ok: false, error: 'Tu sesión venció. Volvé a ingresar el código.', codigo: 'auth' });
      return json(PRIVADAS[a](d));
    }
    return json({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return json({ ok: false, error: mensaje(err) });
  }
}

// ════════════════════════════════════════════════════════════════════════════
// UTILIDADES
// ════════════════════════════════════════════════════════════════════════════
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function error(msg, codigo) { var o = { ok: false, error: msg }; if (codigo) o.codigo = codigo; return o; }
function mensaje(err) { return String(err && err.message ? err.message : err); }
function props() { return PropertiesService.getScriptProperties(); }
function cache() { return CacheService.getScriptCache(); }
function ahoraISO() { return new Date().toISOString(); }
function hoyLocal() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function vacio(v) { return v === '' || v === null || v === undefined; }
function texto(v, max) { var s = String(v === undefined || v === null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim(); return max ? s.slice(0, max) : s; }
function fechaValida(f) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(f || ''))) return false;
  var p = f.split('-').map(Number), d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2];
}
function horaValida(h) { return /^([01]\d|2[0-3]):[0-5]\d$/.test(String(h || '')); }
function diaSemana(f) { var p = f.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay(); }
function sumarDias(f, n) {
  var p = f.split('-').map(Number), d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}
function sha(txt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(txt), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
}
function digitos(n) {   // n dígitos al azar (con la aleatoriedad de getUuid, no Math.random)
  var hex = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''), out = '';
  for (var i = 0; out.length < n && i < hex.length; i++) { var v = parseInt(hex[i], 16); if (v < 10) out += v; }
  while (out.length < n) out += '0';
  return out;
}
// Reintentos: si llega otra vez el mismo pedido (misma "clave"), se devuelve la respuesta anterior sin duplicar nada
function claveRepetida(prefijo, clave) {
  clave = texto(clave, 100);
  if (!clave) return null;
  return 'idem_' + prefijo + '_' + sha(prefijo + '|' + clave).slice(0, 40);
}
function respuestaAnterior(k) {
  if (!k) return null;
  var v = cache().get(k);
  if (!v) return null;
  try { var o = JSON.parse(v); o.repetida = true; return o; } catch (e) { return null; }
}
function recordarRespuesta(k, o) { if (k) cache().put(k, JSON.stringify(o), 21600); return o; }

// Contador que se reinicia cada hora de reloj (para frenar pedidos masivos)
function contadorHora(nombre) {
  var c = cache(), clave = nombre + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHH');
  return { valor: Number(c.get(clave) || 0), sumar: function () { c.put(clave, String(Number(c.get(clave) || 0) + 1), 3700); } };
}
function textoVence(iso) { return Utilities.formatDate(new Date(iso), TZ, "dd/MM 'a las' HH:mm"); }

// Normaliza fecha y hora (vengan como texto o como fecha de la planilla)
function normFecha(raw, visto, tzLibro) {
  if (raw instanceof Date) return Utilities.formatDate(raw, tzLibro, 'yyyy-MM-dd');
  var s = String(vacio(raw) ? '' : raw).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.slice(0, 10);
  return s || String(visto || '').trim();
}
function normHora(raw, visto, tzLibro) {
  var v = String(visto || '').trim();
  var mv = v.match(/^(\d{1,2})[:.](\d{2})/);
  if (mv) return ('0' + mv[1]).slice(-2) + ':' + mv[2];
  if (raw instanceof Date) return Utilities.formatDate(raw, tzLibro, 'HH:mm');
  if (typeof raw === 'number' && raw >= 0 && raw < 1) { var mins = Math.round(raw * 1440); return ('0' + Math.floor(mins / 60)).slice(-2) + ':' + ('0' + (mins % 60)).slice(-2); }
  var s = String(vacio(raw) ? '' : raw).trim(), m = s.match(/^(\d{1,2})(?:[:.](\d{2}))?$/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + (m[2] || '00');
  return s;
}

// ════════════════════════════════════════════════════════════════════════════
// PLANILLA
// ════════════════════════════════════════════════════════════════════════════
var _libro = null;
function libro() {
  if (_libro) return _libro;
  var id = props().getProperty('SHEET_ID');
  _libro = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_libro) throw new Error('No encuentro la planilla (configurá SHEET_ID)');
  return _libro;
}
function hoja(nombre) { return libro().getSheetByName(nombre); }
function asegurarHoja(nombre, encabezados) {
  var sh = hoja(nombre);
  if (!sh) {
    sh = libro().insertSheet(nombre);
    var enc = sh.getRange(1, 1, 1, encabezados.length);
    escribirTexto(enc, [encabezados]);
    enc.setBackground('#9d714c').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
// Escribe SIEMPRE como texto, igual que el servidor anterior: formato "texto sin formato" (@) y el valor como texto.
// Así no hay conversiones automáticas (fechas, números, fórmulas) ni apóstrofes agregados.
function aTexto(v) { return vacio(v) ? '' : String(v); }
function escribirTexto(rango, matriz) {
  rango.setNumberFormat('@');
  rango.setValues(matriz.map(function (fila) { return fila.map(aTexto); }));
}
function proximaFila(sh) { return Math.max(1, sh.getLastRow()) + 1; }
function mapaColumnas(encabezados) {
  var col = {};
  encabezados.forEach(function (h, i) { h = String(h).trim(); if (h && col[h] === undefined) col[h] = i; });  // si un encabezado se repite, vale el primero
  return col;
}

function leerReservas() {
  var sh = hoja(H_RES);
  if (!sh) throw new Error('No encuentro la pestaña "' + H_RES + '"');
  var rango = sh.getDataRange(), valores = rango.getValues(), vistos = rango.getDisplayValues();
  var enc = valores[0].map(function (h) { return String(h).trim(); });
  var col = mapaColumnas(enc);
  if (col.id === undefined || col.fecha === undefined || col.horario === undefined) throw new Error('La pestaña de reservas no tiene las columnas id, fecha y horario');
  var tz = libro().getSpreadsheetTimeZone() || TZ;
  var filas = [];
  for (var r = 1; r < valores.length; r++) {
    var v = valores[r], dv = vistos[r];
    if (v.every(vacio)) continue;
    var o = { _fila: r + 1 };
    Object.keys(col).forEach(function (k) {
      var raw = v[col[k]], vis = dv[col[k]];
      if (k === 'fecha' || k === 'cumpleBebe') o[k] = normFecha(raw, vis, tz);
      else if (k === 'horario') o[k] = normHora(raw, vis, tz);
      else if (k === 'timestamp') o[k] = raw instanceof Date ? raw.toISOString() : texto(raw);
      else o[k] = raw instanceof Date ? Utilities.formatDate(raw, tz, 'yyyy-MM-dd') : (vacio(raw) ? '' : String(raw));
    });
    o.id = String(o.id || '').trim();
    filas.push(o);
  }
  return { hoja: sh, enc: enc, col: col, filas: filas };
}
function buscarFila(res, id) {
  id = String(id || '').trim();
  if (!id) throw new Error('Falta el ID de la reserva');
  var f = res.filas.filter(function (x) { return x.id === id; });
  if (f.length === 0) throw new Error('No encontré la reserva ' + id + ' (¿la borraron o cambiaron en la planilla?)');
  if (f.length > 1) throw new Error('El ID ' + id + ' está repetido en la planilla: revisalo a mano antes de cambiarlo');
  return f[0];
}
function escribirCampos(res, fila, cambios) {
  var antes = {}, despues = {}, aEscribir = [];
  // 1) Se revisa todo antes de escribir: o se escribe todo lo pedido, o nada
  Object.keys(cambios).forEach(function (k) {
    if (res.col[k] === undefined) return;
    var actual = fila[k] === undefined ? '' : String(fila[k]);
    if (actual === String(cambios[k] === undefined || cambios[k] === null ? '' : cambios[k])) return;   // sin cambios: no se toca la celda
    var celda = res.hoja.getRange(fila._fila, res.col[k] + 1);
    if (celda.getFormula()) {
      if (k === 'resta') return;   // "resta" con fórmula: la calcula la planilla, no se pisa
      throw new Error('En la reserva ' + fila.id + ' la columna "' + k + '" tiene una fórmula: cambiala a mano en la planilla');
    }
    aEscribir.push({ k: k, celda: celda, actual: actual });
  });
  // 2) Solo las celdas que cambian
  aEscribir.forEach(function (x) {
    antes[x.k] = x.actual;
    despues[x.k] = cambios[x.k];
    escribirTexto(x.celda, [[cambios[x.k]]]);
    fila[x.k] = cambios[x.k];
  });
  return { antes: antes, despues: despues, hubo: aEscribir.length > 0 };
}
function nuevoId(res) {
  var usados = {};
  res.filas.forEach(function (f) { usados[f.id] = true; });
  for (var i = 0; i < 200; i++) { var id = 'T' + digitos(6); if (!usados[id]) return id; }
  throw new Error('No pude generar un ID libre');
}
function agregarReserva(res, datos) {
  var fila = res.enc.map(function () { return ''; });
  Object.keys(datos).forEach(function (k) { if (res.col[k] !== undefined) fila[res.col[k]] = datos[k]; });
  var nro = proximaFila(res.hoja);
  escribirTexto(res.hoja.getRange(nro, 1, 1, fila.length), [fila]);
  var o = { _fila: nro };
  Object.keys(datos).forEach(function (k) { o[k] = datos[k]; });
  res.filas.push(o);
  return o;
}

function leerSeguimiento() {
  var sh = hoja(H_SEG), mapa = {};
  if (!sh || sh.getLastRow() < 2) return mapa;
  var v = sh.getDataRange().getValues(), col = mapaColumnas(v[0]);
  for (var r = 1; r < v.length; r++) {
    var id = String(v[r][col.id] || '').trim();
    if (!id) continue;
    var o = { _fila: r + 1 };
    ENC_SEG.forEach(function (k) { var x = col[k] === undefined ? '' : v[r][col[k]]; o[k] = x instanceof Date ? x.toISOString() : (vacio(x) ? '' : String(x)); });
    mapa[id] = o;
  }
  return mapa;
}
function guardarSeguimiento(id, cambios, seg) {
  var sh = asegurarHoja(H_SEG, ENC_SEG);
  var col = mapaColumnas(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  var actual = seg && seg[id];
  if (actual) {
    Object.keys(cambios).forEach(function (k) { if (col[k] !== undefined) escribirTexto(sh.getRange(actual._fila, col[k] + 1), [[cambios[k]]]); actual[k] = cambios[k]; });
  } else {
    var fila = ENC_SEG.map(function (k) { return k === 'id' ? id : cambios[k]; });
    var nro = proximaFila(sh);
    escribirTexto(sh.getRange(nro, 1, 1, fila.length), [fila]);
    if (seg) { var o = { _fila: nro, id: id }; Object.keys(cambios).forEach(function (k) { o[k] = cambios[k]; }); seg[id] = o; }
  }
}
// Fila tal como se ve en la planilla (texto), con el ancho pedido
function filaVista(visto, ncol) {
  var f = [];
  for (var j = 0; j < ncol; j++) f.push(visto && visto[j] !== undefined && visto[j] !== null ? String(visto[j]) : '');
  return f;
}
function registrar(accion, id, antes, despues) {
  var sh = asegurarHoja(H_REG, ENC_REG);
  var fila = [ahoraISO(), accion, id || '', antes === undefined ? '' : JSON.stringify(antes), despues === undefined ? '' : JSON.stringify(despues)];
  escribirTexto(sh.getRange(proximaFila(sh), 1, 1, fila.length), [fila]);
}

function leerConfig() {
  var sh = hoja(H_CONF);
  if (!sh) throw new Error('No encuentro la pestaña "' + H_CONF + '"');
  var rango = sh.getDataRange(), v = rango.getValues(), dv = rango.getDisplayValues();
  var col = mapaColumnas(v[0]), tz = libro().getSpreadsheetTimeZone() || TZ, lista = [], otras = [];
  for (var r = 1; r < v.length; r++) {
    if (v[r].every(vacio)) continue;
    var tipo = String(v[r][col.tipo] || '').trim();
    var fecha = normFecha(v[r][col.fecha], dv[r][col.fecha], tz);
    var horario = col.horario === undefined || vacio(v[r][col.horario]) ? '' : normHora(v[r][col.horario], dv[r][col.horario], tz);
    var nota = col.nota === undefined ? '' : texto(v[r][col.nota]);
    if (TIPOS_BLOQUEO.indexOf(tipo) >= 0 && fechaValida(fecha)) lista.push({ tipo: tipo, fecha: fecha, horario: horario, nota: nota, visto: dv[r] });
    else otras.push(dv[r]);   // filas que este panel no maneja: se conservan tal cual
  }
  return { hoja: sh, col: col, enc: v[0], lista: lista, otras: otras, version: versionConfig(lista) };
}
function versionConfig(lista) {
  return sha(lista.map(function (b) { return [b.tipo, b.fecha, b.horario || ''].join('|'); }).sort().join('\n')).slice(0, 16);
}

function leerAjustes() {
  var aj = {};
  Object.keys(AJUSTES_BASE).forEach(function (k) { aj[k] = AJUSTES_BASE[k].valor; });
  var sh = hoja(H_AJ);
  if (sh && sh.getLastRow() >= 2) {
    var v = sh.getDataRange().getDisplayValues();
    for (var r = 1; r < v.length; r++) { var k = String(v[r][0] || '').trim(); if (AJUSTES_BASE[k]) aj[k] = String(v[r][1] || '').trim(); }
  }
  var lista = function (s) { return String(s || '').split(',').map(function (x) { return normHora(x.trim(), x.trim(), TZ); }).filter(horaValida); };
  return {
    senaMonto: Number(String(aj.senaMonto).replace(/[^\d]/g, '')) || 0,
    senaHoras: Math.max(1, Number(aj.senaHoras) || 24),
    alias: aj.alias,
    whatsappEstudio: String(aj.whatsappEstudio).replace(/\D/g, ''),
    grillaSemana: lista(aj.grillaSemana),
    grillaSabado: lista(aj.grillaSabado),
    minDias: Math.max(0, Number(aj.minDias) || 0),
    tiposSesion: String(aj.tiposSesion || '').split(',').map(function (x) { return x.trim(); }).filter(String),
    avisoReservaMail: /^n/i.test(String(aj.avisoReservaMail || '')) ? 'no' : 'si'
  };
}

// ════════════════════════════════════════════════════════════════════════════
// DISPONIBILIDAD
// ════════════════════════════════════════════════════════════════════════════
function estadoSena(r, s, ahora) {
  if (!s) return '';
  if (s.senaRecibida || parseFloat(r.senya) > 0) return 'recibida';
  if (s.asegurado || String(r.estado || '').toUpperCase() === 'CONFIRMADO') return 'asegurada';
  if (String(r.estado || '').toUpperCase() === 'CANCELADO') return '';
  if (s.venceSena && new Date(s.venceSena).getTime() < ahora) return 'vencida';
  return s.venceSena ? 'esperando' : '';
}
// ¿Esta reserva ocupa su horario? (las reservas web con la seña vencida liberan el horario solas)
function ocupaHorario(r, s, ahora) {
  if (String(r.estado || '').toUpperCase() === 'CANCELADO') return false;
  if (!r.fecha || !r.horario) return false;
  if (s && s.origen === 'clienta' && estadoSena(r, s, ahora) === 'vencida') return false;
  return true;
}
function horariosDelDia(fecha, aj, bloqueos) {
  var dow = diaSemana(fecha);
  var base = dow === 0 ? [] : (dow === 6 ? aj.grillaSabado : aj.grillaSemana).slice();
  bloqueos.forEach(function (b) { if (b.tipo === 'slot_extra' && b.fecha === fecha && horaValida(b.horario) && base.indexOf(b.horario) < 0) base.push(b.horario); });
  return base.sort();
}
function motivoNoDisponible(fecha, horario, ctx, publico, ignorarId) {
  var bl = ctx.config.lista;
  if (bl.some(function (b) { return b.tipo === 'dia' && b.fecha === fecha; })) return 'Ese día no hay turnos';
  if (bl.some(function (b) { return b.tipo === 'slot' && b.fecha === fecha && b.horario === horario; })) return 'Ese horario está bloqueado';
  if (publico) {
    if (fecha < sumarDias(hoyLocal(), ctx.ajustes.minDias)) return 'Ese día ya no se puede reservar desde la web';
    if (horariosDelDia(fecha, ctx.ajustes, bl).indexOf(horario) < 0) return 'Ese horario no está disponible';
  }
  var ahora = Date.now();
  var choque = ctx.reservas.filas.filter(function (r) {
    return r.id !== ignorarId && r.fecha === fecha && r.horario === horario && ocupaHorario(r, ctx.seguimiento[r.id], ahora);
  })[0];
  if (choque) return publico ? 'Ese horario ya fue reservado' : 'Ese horario ya está ocupado por ' + (choque.nombreBebe || 'otra reserva') + ' (' + choque.id + ')';
  return null;
}
function contexto() {
  return { reservas: leerReservas(), seguimiento: leerSeguimiento(), config: leerConfig(), ajustes: leerAjustes() };
}

// ════════════════════════════════════════════════════════════════════════════
// API PÚBLICA (formulario de reservas)
// ════════════════════════════════════════════════════════════════════════════
function apiHorarios() {
  var ctx = contexto(), hoy = hoyLocal(), ahora = Date.now(), ocupados = {};
  ctx.reservas.filas.forEach(function (r) {
    if (r.fecha < hoy || !ocupaHorario(r, ctx.seguimiento[r.id], ahora)) return;
    (ocupados[r.fecha] = ocupados[r.fecha] || []);
    if (ocupados[r.fecha].indexOf(r.horario) < 0) ocupados[r.fecha].push(r.horario);
  });
  return {
    ok: true,
    hoy: hoy,
    maxDias: MAX_DIAS_WEB,
    ocupados: ocupados,
    bloqueos: ctx.config.lista.filter(function (b) { return b.fecha >= hoy; }).map(function (b) { return { tipo: b.tipo, fecha: b.fecha, horario: b.horario }; }),
    grilla: { semana: ctx.ajustes.grillaSemana, sabado: ctx.ajustes.grillaSabado },
    minDias: ctx.ajustes.minDias,
    tiposSesion: ctx.ajustes.tiposSesion,
    sena: { monto: ctx.ajustes.senaMonto, horas: ctx.ajustes.senaHoras }
  };
}

function apiReservar(d) {
  var kIdem = claveRepetida('reservar', d.clave), anterior = respuestaAnterior(kIdem);
  if (anterior) return anterior;
  var aj = leerAjustes();
  var datos = {
    fecha: texto(d.fecha, 10), horario: texto(d.horario, 5),
    tipoSesion: texto(d.tipoSesion, 40), tematica: texto(d.tematica, 80),
    nombreBebe: texto(d.nombreBebe, 60), cumpleBebe: texto(d.cumpleBebe, 10),
    nombreMama: texto(d.nombreMama, 80), telefono: texto(d.telefono, 25), whatsapp: texto(d.whatsapp, 25)
  };
  if (!fechaValida(datos.fecha) || !horaValida(datos.horario)) return error('Fecha u horario inválidos');
  if (!datos.nombreBebe || !datos.nombreMama) return error('Faltan el nombre del bebé o de la mamá');
  if (aj.tiposSesion.indexOf(datos.tipoSesion) < 0) return error('Tipo de sesión inválido');
  if (datos.cumpleBebe && !fechaValida(datos.cumpleBebe)) return error('La fecha de cumpleaños no es válida');
  var tel = datos.telefono.replace(/\D/g, ''), wa = (datos.whatsapp || datos.telefono).replace(/\D/g, '');
  if (tel.length < 8 || tel.length > 15 || wa.length < 8 || wa.length > 15) return error('Revisá el teléfono');
  if (!datos.whatsapp) datos.whatsapp = datos.telefono;

  if (datos.fecha > sumarDias(hoyLocal(), MAX_DIAS_WEB)) return error('Esa fecha todavía no está habilitada');

  // Freno a pedidos masivos (alguien que quiera bloquear la agenda con reservas falsas)
  var cont = contadorHora('reservas_web');
  if (cont.valor >= 15) return error('Hay demasiadas reservas en este momento. Probá en un rato o escribinos por WhatsApp.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    anterior = respuestaAnterior(kIdem);
    if (anterior) return anterior;
    var ctx = contexto(), ahora = Date.now();
    var motivo = motivoNoDisponible(datos.fecha, datos.horario, ctx, true);
    if (motivo) return error(motivo, 'ocupado');
    var pendientesMismoTel = ctx.reservas.filas.filter(function (r) {
      var s = ctx.seguimiento[r.id];
      return s && s.origen === 'clienta' && String(r.telefono || '').replace(/\D/g, '') === tel && estadoSena(r, s, ahora) === 'esperando';
    }).length;
    if (pendientesMismoTel >= 2) return error('Ya tenés reservas esperando la seña. Escribinos por WhatsApp y lo vemos.');
    var vence = new Date(ahora + ctx.ajustes.senaHoras * 3600 * 1000).toISOString();
    var id = nuevoId(ctx.reservas);
    datos.id = id;
    datos.estado = 'PENDIENTE';
    datos.timestamp = ahoraISO();
    agregarReserva(ctx.reservas, datos);
    guardarSeguimiento(id, { origen: 'clienta', creado: datos.timestamp, venceSena: vence }, ctx.seguimiento);
    registrar('reservar (web)', id, undefined, datos);
    cont.sumar();
    avisarReservaWeb(datos, vence, ctx.ajustes);
    return recordarRespuesta(kIdem, {
      ok: true, id: id, fecha: datos.fecha, horario: datos.horario,
      sena: { monto: ctx.ajustes.senaMonto, horas: ctx.ajustes.senaHoras, vence: vence, venceTexto: textoVence(vence), alias: ctx.ajustes.alias, whatsapp: ctx.ajustes.whatsappEstudio }
    });
  } finally {
    lock.releaseLock();
  }
}

function avisarReservaWeb(datos, vence, aj) {
  if (String(aj.avisoReservaMail || 'si').toLowerCase().indexOf('n') === 0) return;
  try {
    var wa = String(datos.whatsapp || datos.telefono || '').replace(/\D/g, '');
    avisarAdmins('Nueva reserva web: ' + datos.nombreBebe + ' · ' + datos.fecha.split('-').reverse().join('/') + ' ' + datos.horario,
      'Entró una reserva desde la web:\n\n' +
      '  ID: ' + datos.id + '\n  Fecha: ' + datos.fecha.split('-').reverse().join('/') + ' a las ' + datos.horario + '\n' +
      '  Sesión: ' + datos.tipoSesion + (datos.tematica ? ' (' + datos.tematica + ')' : '') + '\n' +
      '  Bebé: ' + datos.nombreBebe + '\n  Mamá: ' + datos.nombreMama + '\n  WhatsApp: ' + (wa ? 'https://wa.me/' + wa : '-') + '\n\n' +
      'Queda esperando la seña de $' + aj.senaMonto + ' hasta el ' + textoVence(vence) + ' h. Si no la recibe, el horario se libera solo.\n' +
      'Cuando llegue la transferencia, tocá "Seña recibida" en el panel.\n');
  } catch (e) { /* el aviso es opcional: nunca frena la reserva */ }
}

// ════════════════════════════════════════════════════════════════════════════
// ACCESO AL PANEL
// ════════════════════════════════════════════════════════════════════════════
function hashCodigo(codigo, sal) {
  var h = sha(sal + '|' + codigo);
  for (var i = 0; i < 300; i++) h = sha(sal + h);
  return h;
}
function guardarCodigo(codigo) {
  var sal = Utilities.getUuid();
  props().setProperties({ CODIGO_SAL: sal, CODIGO_HASH: hashCodigo(String(codigo), sal) });
}
function asegurarCodigo() {
  var p = props();
  if (p.getProperty('CODIGO_HASH')) return true;
  var ini = p.getProperty('CODIGO_INICIAL');
  if (!ini) return false;
  guardarCodigo(ini);
  p.deleteProperty('CODIGO_INICIAL');
  return true;
}
function codigoCorrecto(codigo) {
  if (!asegurarCodigo()) return false;
  var p = props();
  return hashCodigo(String(codigo || ''), p.getProperty('CODIGO_SAL')) === p.getProperty('CODIGO_HASH');
}
function leerSesiones() { try { return JSON.parse(props().getProperty('SESIONES') || '{}'); } catch (e) { return {}; } }
function guardarSesiones(s) { props().setProperty('SESIONES', JSON.stringify(s)); }
function crearSesion(recordar) {
  var token = Utilities.getUuid() + Utilities.getUuid();
  var s = leerSesiones(), ahora = Date.now();
  Object.keys(s).forEach(function (k) { if (!(s[k].vence > ahora)) delete s[k]; });
  var claves = Object.keys(s).sort(function (a, b) { return s[a].creada - s[b].creada; });
  while (claves.length >= 30) delete s[claves.shift()];
  s[sha('tok|' + token)] = { creada: ahora, vence: ahora + (recordar ? 30 : 1) * 24 * 3600 * 1000 };
  guardarSesiones(s);
  return token;
}
function sesionValida(token) {
  if (!token) return false;
  var x = leerSesiones()[sha('tok|' + String(token))];
  return !!(x && x.vence > Date.now());
}
function emailsAdmin() {
  return String(props().getProperty('EMAILS_ADMIN') || '').split(',').map(function (s) { return s.trim(); })
    .filter(function (s) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s); });
}
function avisarAdmins(asunto, cuerpo) {
  var to = emailsAdmin();
  if (!to.length) return false;
  MailApp.sendEmail({ to: to.join(','), subject: asunto, body: cuerpo, name: 'Panel Estudio Fotoinfantil' });
  return true;
}
function avisarCambioCodigo(como) {
  avisarAdmins('Se cambió el código de acceso del panel',
    'Hola!\n\nSe cambió el código de acceso del panel de reservas el ' + Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy 'a las' HH:mm") + ' (' + como + ').\n\n' +
    'Si fueron ustedes, no hace falta hacer nada.\nSi NO fueron ustedes: entren al panel, toquen "Olvidé el código" y elijan uno nuevo.\n');
}

function apiLogin(d) {
  var c = cache(), fallos = Number(c.get('login_fallos') || 0);
  if (fallos >= 5) return error('Demasiados intentos. Probá de nuevo en 15 minutos.', 'bloqueado');
  if (!asegurarCodigo()) return error('Todavía no hay código de acceso: tocá «Olvidé el código» y elegí uno (llega un mail a los correos del estudio).', 'sin_codigo');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!codigoCorrecto(d.codigo)) {
      c.put('login_fallos', String(fallos + 1), 900);
      return error(fallos + 1 >= 5 ? 'Código incorrecto. Se bloqueó el ingreso por 15 minutos.' : 'Código incorrecto', 'auth');
    }
    c.remove('login_fallos');
    return { ok: true, token: crearSesion(!!d.recordar), version: VERSION };
  } finally {
    lock.releaseLock();
  }
}
function apiLogout(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { var s = leerSesiones(); delete s[sha('tok|' + String(d.token))]; guardarSesiones(s); return { ok: true }; }
  finally { lock.releaseLock(); }
}
function apiCerrarSesiones() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { guardarSesiones({}); return { ok: true }; } finally { lock.releaseLock(); }
}
function apiCambiarCodigo(d) {
  var nuevo = String(d.nuevo || '');
  if (nuevo.trim().length < 8) return error('El código nuevo tiene que tener al menos 8 caracteres');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!codigoCorrecto(d.actual)) return error('El código actual no es correcto');
    guardarCodigo(nuevo);
    guardarSesiones({});
    var token = crearSesion(!!d.recordar);
    avisarCambioCodigo('desde el panel');
    return { ok: true, token: token };
  } finally {
    lock.releaseLock();
  }
}
// "Olvidé el código": manda un código de 6 números a los correos registrados (nunca a otro correo)
function apiRecuperar() {
  var c = cache(), cont = contadorHora('recuperar');
  if (cont.valor >= 3) return { ok: true };        // respuesta igual siempre: no revela nada
  cont.sumar();
  if (!emailsAdmin().length) return { ok: true };
  var codigo = digitos(6);
  c.put('recuperar_codigo', sha('rec|' + codigo), 900);
  c.remove('recuperar_intentos');
  avisarAdmins('Código para recuperar el acceso al panel',
    'Hola!\n\nTu código para recuperar el acceso al panel de reservas es:\n\n    ' + codigo + '\n\n' +
    'Vence en 15 minutos y sirve una sola vez.\nSi no lo pidieron ustedes, ignoren este mail: sin este código nadie puede entrar.\n');
  return { ok: true };
}
function apiRestablecer(d) {
  var c = cache(), guardado = c.get('recuperar_codigo');
  if (!guardado) return error('El código venció o no existe. Pedí uno nuevo.');
  var intentos = Number(c.get('recuperar_intentos') || 0);
  if (intentos >= 5) { c.remove('recuperar_codigo'); return error('Demasiados intentos. Pedí un código nuevo.'); }
  if (sha('rec|' + String(d.codigoMail || '').trim()) !== guardado) {
    c.put('recuperar_intentos', String(intentos + 1), 900);
    return error('El código del mail no es correcto');
  }
  var nuevo = String(d.nuevo || '');
  if (nuevo.trim().length < 8) return error('El código nuevo tiene que tener al menos 8 caracteres');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    guardarCodigo(nuevo);
    props().deleteProperty('CODIGO_INICIAL');
    guardarSesiones({});
    c.remove('recuperar_codigo'); c.remove('recuperar_intentos'); c.remove('login_fallos');
    var token = crearSesion(!!d.recordar);
    avisarCambioCodigo('con el código enviado por mail');
    return { ok: true, token: token };
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════════════════
// API DEL PANEL (necesita sesión)
// ════════════════════════════════════════════════════════════════════════════
function filaParaPanel(f, s, ahora) {
  var o = {};
  Object.keys(f).forEach(function (k) { if (k !== '_fila') o[k] = f[k]; });
  o.sena = s ? { origen: s.origen, creado: s.creado, vence: s.venceSena, recibida: s.senaRecibida, asegurado: s.asegurado, estado: estadoSena(f, s, ahora) } : null;
  return o;
}
function apiGetAll() {
  var res = leerReservas(), seg = leerSeguimiento(), ahora = Date.now();
  return { ok: true, reservas: res.filas.map(function (f) { return filaParaPanel(f, seg[f.id], ahora); }), version: VERSION };
}
function apiGetConfig() {
  var conf = leerConfig();
  return { ok: true, bloqueos: conf.lista.map(function (b) { return { tipo: b.tipo, fecha: b.fecha, horario: b.horario, nota: b.nota }; }), version: conf.version };
}
function apiSaveConfig(d) {
  if (!Array.isArray(d.bloqueos)) return error('Lista de bloqueos inválida');
  var nueva = [], vistos = {};
  for (var i = 0; i < d.bloqueos.length; i++) {
    var b = d.bloqueos[i] || {}, tipo = texto(b.tipo), fecha = texto(b.fecha), horario = texto(b.horario);
    if (TIPOS_BLOQUEO.indexOf(tipo) < 0 || !fechaValida(fecha)) return error('Bloqueo inválido: ' + JSON.stringify(b));
    if (tipo === 'dia') horario = '';
    else if (!horaValida(horario)) return error('Horario inválido en bloqueo: ' + JSON.stringify(b));
    var k = tipo + '|' + fecha + '|' + horario;
    if (vistos[k]) continue;
    vistos[k] = true;
    nueva.push({ tipo: tipo, fecha: fecha, horario: horario, nota: b.nota === undefined ? undefined : texto(b.nota, 200) });
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var conf = leerConfig();
    if (d.version && d.version !== conf.version) return error('La configuración cambió desde otro dispositivo. Volvé a entrar a Config y repetí el cambio.', 'cambio');
    var sh = conf.hoja, ncol = Math.max(4, sh.getLastColumn()), col = conf.col, previos = {};
    conf.lista.forEach(function (b) { previos[b.tipo + '|' + b.fecha + '|' + (b.horario || '')] = b; });
    var filas = nueva.map(function (b) {
      var p = previos[b.tipo + '|' + b.fecha + '|' + (b.horario || '')];
      if (b.nota === undefined) b.nota = p ? p.nota : '';   // la nota se conserva
      // Fila nueva o existente: tipo, fecha, horario y nota normalizados; si ya existía, el resto de sus columnas queda igual
      var f = p ? filaVista(p.visto, ncol) : filaVista(null, ncol);
      f[col.tipo] = b.tipo;
      f[col.fecha] = b.fecha;
      if (col.horario !== undefined) f[col.horario] = b.horario || '';
      if (col.nota !== undefined) f[col.nota] = b.nota || '';
      return f;
    });
    conf.otras.forEach(function (visto) { filas.push(filaVista(visto, ncol)); });
    // Primero queda anotado lo anterior; después se escribe encima y recién al final se limpian las filas sobrantes
    registrar('saveConfig', '', conf.lista.map(function (b) { return { tipo: b.tipo, fecha: b.fecha, horario: b.horario, nota: b.nota }; }), nueva);
    var filasViejas = Math.max(0, sh.getLastRow() - 1);
    if (filas.length) escribirTexto(sh.getRange(2, 1, filas.length, ncol), filas);
    if (filasViejas > filas.length) sh.getRange(2 + filas.length, 1, filasViejas - filas.length, ncol).clearContent();
    var v = versionConfig(nueva);
    return { ok: true, version: v, bloqueos: nueva.map(function (b) { return { tipo: b.tipo, fecha: b.fecha, horario: b.horario, nota: b.nota }; }) };
  } finally {
    lock.releaseLock();
  }
}

function conCambios(id, fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ctx = contexto(), fila = buscarFila(ctx.reservas, id);
    var r = fn(ctx, fila);
    if (r && r.ok === false) return r;
    return { ok: true, reserva: filaParaPanel(fila, ctx.seguimiento[fila.id], Date.now()) };
  } finally {
    lock.releaseLock();
  }
}
function apiUpdateStatus(d) {
  var st = String(d.status || '').toUpperCase();
  if (ESTADOS.indexOf(st) < 0) return error('Estado inválido');
  return conCambios(d.id, function (ctx, fila) {
    if (st !== 'CANCELADO' && String(fila.estado).toUpperCase() === 'CANCELADO') {
      var m = motivoNoDisponible(fila.fecha, fila.horario, ctx, false, fila.id);
      if (m && !d.forzar) return error(m + '. ¿Reactivar igual?', 'ocupado');
    }
    var estabaCancelada = String(fila.estado).toUpperCase() === 'CANCELADO';
    var c = escribirCampos(ctx.reservas, fila, { estado: st });
    if (c.hubo) registrar('updateStatus', fila.id, c.antes, c.despues);
    if (st === 'CONFIRMADO') asegurarSiEsWeb(ctx, fila, 'confirmada');
    else if (st === 'PENDIENTE' && estabaCancelada) asegurarSiEsWeb(ctx, fila, 'reactivada');
  });
}
function numeroOVacio(v) { var s = texto(v, 20); return s === '' || /^\d+(\.\d+)?$/.test(s) ? s : null; }
function apiUpdatePago(d) {
  var notas = String(d.notasInternas === undefined ? '' : d.notasInternas).slice(0, 2000);
  return conCambios(d.id, function (ctx, fila) {
    var cambios = { notasInternas: notas }, mal = [];
    ['senya', 'total', 'resta'].forEach(function (k) {
      var crudo = d[k] === undefined || d[k] === null ? '' : String(d[k]);
      if (crudo === String(fila[k] === undefined ? '' : fila[k])) { cambios[k] = crudo; return; }   // sin cambios: se acepta tal cual estaba
      var v = texto(crudo, 20);
      if (v === '' || (k === 'resta' ? /^-?\d+(\.\d+)?$/ : /^\d+(\.\d+)?$/).test(v)) cambios[k] = v;
      else mal.push(k === 'senya' ? 'seña' : k);
    });
    if (mal.length) return error('Los montos tienen que ser números, sin puntos ni $ (' + mal.join(', ') + ')');
    var senaAntes = parseFloat(fila.senya) > 0;
    var c = escribirCampos(ctx.reservas, fila, cambios);
    if (c.hubo) registrar('updatePago', fila.id, c.antes, c.despues);
    var s = ctx.seguimiento[fila.id];
    if (!senaAntes && parseFloat(fila.senya) > 0 && s && !s.senaRecibida) guardarSeguimiento(fila.id, { senaRecibida: ahoraISO() }, ctx.seguimiento);   // seña anotada en "Pago": ya no se libera sola
  });
}
// Si el estudio toca una reserva de la web que no tiene seña, queda asegurada (ya no se libera sola)
function asegurarSiEsWeb(ctx, fila, motivo) {
  var s = ctx.seguimiento[fila.id];
  if (s && s.origen === 'clienta' && !s.senaRecibida && !s.asegurado && !(parseFloat(fila.senya) > 0)) {
    guardarSeguimiento(fila.id, { asegurado: ahoraISO() }, ctx.seguimiento);
    registrar('asegurar sin seña (' + motivo + ')', fila.id, undefined, { asegurado: true });
  }
}
function apiSenaRecibida(d) {
  return conCambios(d.id, function (ctx, fila) {
    var monto = numeroOVacio(d.monto === undefined ? String(ctx.ajustes.senaMonto) : d.monto);
    if (!monto) return error('Monto de seña inválido');
    var cambios = { senya: monto, estado: 'CONFIRMADO' };
    var total = parseFloat(fila.total);
    if (!isNaN(total) && total > 0) cambios.resta = String(Math.max(0, total - parseFloat(monto)));
    var c = escribirCampos(ctx.reservas, fila, cambios);
    guardarSeguimiento(fila.id, { senaRecibida: ahoraISO() }, ctx.seguimiento);
    registrar('senaRecibida', fila.id, c.antes, Object.assign({ senaRecibida: true }, c.despues));
  });
}
function apiReprogramar(d) {
  var nf = texto(d.newFecha), nh = texto(d.newHorario);
  if (!nf && !nh) return error('Indicá la fecha o el horario nuevo');
  if (nf && !fechaValida(nf)) return error('Fecha inválida');
  if (nh && !horaValida(nh)) return error('Horario inválido');
  return conCambios(d.id, function (ctx, fila) {
    var f = nf || fila.fecha, h = nh || fila.horario;
    var m = motivoNoDisponible(f, h, ctx, false, fila.id);
    if (m && !d.forzar) return error(m, 'ocupado');
    var cambios = {};
    if (f !== fila.fecha) cambios.fecha = f;
    if (h !== fila.horario) cambios.horario = h;
    if (!Object.keys(cambios).length) return error('Es el mismo día y horario');
    var c = escribirCampos(ctx.reservas, fila, cambios);
    registrar('reprogramar', fila.id, c.antes, c.despues);
    asegurarSiEsWeb(ctx, fila, 'movida');
  });
}
function apiNuevoTurno(d) {
  var kIdem = claveRepetida('nuevoTurno', d.clave), anterior = respuestaAnterior(kIdem);
  if (anterior) return anterior;
  var aj = leerAjustes(), t = d.turno || {};
  var datos = {
    fecha: texto(t.fecha, 10), horario: normHora(texto(t.horario, 5), texto(t.horario, 5), TZ),
    tipoSesion: texto(t.tipoSesion, 40), tematica: texto(t.tematica, 80),
    nombreBebe: texto(t.nombreBebe, 60), cumpleBebe: texto(t.cumpleBebe, 10),
    nombreMama: texto(t.nombreMama, 80), telefono: texto(t.telefono, 25), whatsapp: texto(t.whatsapp, 25),
    total: numeroOVacio(t.total), senya: numeroOVacio(t.senya), notasInternas: String(t.notasInternas || '').slice(0, 2000)
  };
  if (!fechaValida(datos.fecha) || !horaValida(datos.horario)) return error('Fecha u horario inválidos');
  if (!datos.nombreBebe) return error('Falta el nombre del bebé');
  if (datos.cumpleBebe && !fechaValida(datos.cumpleBebe)) return error('La fecha de cumpleaños no es válida');
  if (datos.total === null || datos.senya === null) return error('Los montos tienen que ser números');
  var conSena = !!datos.senya && parseFloat(datos.senya) > 0;
  if (!datos.whatsapp) datos.whatsapp = datos.telefono;
  var asegurado = !!d.asegurado;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    anterior = respuestaAnterior(kIdem);
    if (anterior) return anterior;
    var ctx = contexto();
    var m = motivoNoDisponible(datos.fecha, datos.horario, ctx, false);
    if (m && !d.forzar) return error(m + '. ¿Cargarlo igual?', 'ocupado');
    datos.id = nuevoId(ctx.reservas);
    datos.estado = asegurado || conSena ? 'CONFIRMADO' : 'PENDIENTE';
    datos.timestamp = ahoraISO();
    if (datos.total) datos.resta = String(Math.max(0, parseFloat(datos.total) - (conSena ? parseFloat(datos.senya) : 0)));
    if (!conSena) delete datos.senya;
    var fila = agregarReserva(ctx.reservas, datos);
    var seg = { origen: 'panel', creado: datos.timestamp };
    if (conSena) seg.senaRecibida = datos.timestamp;
    else if (asegurado) seg.asegurado = datos.timestamp;
    else seg.venceSena = new Date(Date.now() + aj.senaHoras * 3600 * 1000).toISOString();
    guardarSeguimiento(datos.id, seg, ctx.seguimiento);
    registrar('nuevoTurno (panel)', datos.id, undefined, datos);
    return recordarRespuesta(kIdem, { ok: true, reserva: filaParaPanel(fila, ctx.seguimiento[datos.id], Date.now()) });
  } finally {
    lock.releaseLock();
  }
}
function apiGetAjustes() {
  var aj = leerAjustes();
  return { ok: true, ajustes: {
    senaMonto: aj.senaMonto, senaHoras: aj.senaHoras, alias: aj.alias, whatsappEstudio: aj.whatsappEstudio,
    grillaSemana: aj.grillaSemana.join(','), grillaSabado: aj.grillaSabado.join(','), minDias: aj.minDias, tiposSesion: aj.tiposSesion.join(','),
    avisoReservaMail: aj.avisoReservaMail
  } };
}
function apiSaveAjustes(d) {
  var nuevos = d.ajustes || {}, limpios = {};
  for (var k in nuevos) {
    if (!AJUSTES_BASE[k]) continue;
    var v = texto(nuevos[k], 200);
    if ((k === 'senaMonto' || k === 'senaHoras' || k === 'minDias') && !/^\d+$/.test(v)) return error('"' + k + '" tiene que ser un número');
    if ((k === 'grillaSemana' || k === 'grillaSabado') && v && !v.split(',').every(function (h) { return horaValida(normHora(h.trim(), h.trim(), TZ)); })) return error('Revisá los horarios de "' + k + '" (formato 18:00,19:00)');
    if (k === 'avisoReservaMail') v = /^n/i.test(v) ? 'no' : 'si';
    if (k === 'tiposSesion' && !v) return error('Tiene que haber al menos un tipo de sesión');
    if (k === 'whatsappEstudio' && v.replace(/\D/g, '').length < 10) return error('Revisá el WhatsApp del estudio (con 54, sin + ni espacios)');
    limpios[k] = v;
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = crearAjustesSiFalta(), v = sh.getDataRange().getValues(), antes = {};
    Object.keys(limpios).forEach(function (k) {
      var fila = -1;
      for (var r = 1; r < v.length; r++) if (String(v[r][0]).trim() === k) { fila = r + 1; antes[k] = String(v[r][1]); break; }
      if (fila < 0) { escribirTexto(sh.getRange(proximaFila(sh), 1, 1, 3), [[k, limpios[k], AJUSTES_BASE[k].desc]]); v = sh.getDataRange().getValues(); }
      else escribirTexto(sh.getRange(fila, 2), [[limpios[k]]]);
    });
    registrar('saveAjustes', '', antes, limpios);
    return apiGetAjustes();
  } finally {
    lock.releaseLock();
  }
}
function crearAjustesSiFalta() {
  var sh = hoja(H_AJ);
  if (sh) return sh;
  sh = asegurarHoja(H_AJ, ENC_AJ);
  var filas = Object.keys(AJUSTES_BASE).map(function (k) { return [k, AJUSTES_BASE[k].valor, AJUSTES_BASE[k].desc]; });
  escribirTexto(sh.getRange(2, 1, filas.length, 3), filas);
  return sh;
}
function apiCopiaAhora() { return { ok: true, nombre: copiaDiaria() }; }

// ════════════════════════════════════════════════════════════════════════════
// COPIAS DE SEGURIDAD
// ════════════════════════════════════════════════════════════════════════════
function carpetaRespaldos() {
  var it = DriveApp.getFoldersByName(CARPETA_RESPALDOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_RESPALDOS);
}
function copiaDiaria() {
  var ss = libro(), carpeta = carpetaRespaldos();
  var nombre = 'Respaldo ' + ss.getName() + ' ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH-mm');
  DriveApp.getFileById(ss.getId()).makeCopy(nombre, carpeta);
  var archivos = [], it = carpeta.getFiles();
  while (it.hasNext()) { var f = it.next(); if (String(f.getName()).indexOf('Respaldo ') === 0) archivos.push(f); }
  archivos.sort(function (a, b) { return b.getDateCreated().getTime() - a.getDateCreated().getTime(); });
  archivos.slice(MAX_COPIAS).forEach(function (f) { f.setTrashed(true); });   // conserva las últimas 30
  return nombre;
}
function instalarCopiaDiaria() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'copiaDiaria') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('copiaDiaria').timeBased().everyDays(1).atHour(3).inTimezone(TZ).create();
  var nombre = copiaDiaria();
  Logger.log('Listo: copia diaria programada (3 a. m.). Primera copia: ' + nombre + ' en la carpeta "' + CARPETA_RESPALDOS + '".');
  return nombre;
}

// ════════════════════════════════════════════════════════════════════════════
// IA (Gemini)
// ════════════════════════════════════════════════════════════════════════════
function datosParaIA() {
  var res = leerReservas(), seg = leerSeguimiento(), hoy = hoyLocal(), desde = sumarDias(hoy, -45), ahora = Date.now();
  return res.filas.filter(function (f) { return String(f.estado).toUpperCase() !== 'CANCELADO' && f.fecha >= desde; })
    .map(function (f) {
      var total = parseFloat(f.total), senya = parseFloat(f.senya);
      return { id: f.id, fecha: f.fecha, horario: f.horario, bebe: f.nombreBebe, responsable: f.nombreMama, tipo: f.tipoSesion,
        tematica: f.tematica, estado: f.estado, cumple: f.cumpleBebe, sena: estadoSena(f, seg[f.id], ahora) || undefined,
        saldo: isNaN(total) ? undefined : Math.max(0, total - (isNaN(senya) ? 0 : senya)) };     // sin teléfonos
    });
}
function llamarGemini(sistema, mensajes) {
  var key = props().getProperty('GEMINI_API_KEY') || props().getProperty('GEMINI_KEY');   // GEMINI_KEY: nombre del servidor anterior
  if (!key) throw new Error('Falta la clave de la IA (GEMINI_API_KEY) en las propiedades del script');
  var modelo = props().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelo) + ':generateContent';
  var cuerpo = {
    systemInstruction: { parts: [{ text: sistema }] },
    contents: mensajes.map(function (m) { return { role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: String(m.content || '').slice(0, 4000) }] }; }),
    generationConfig: { temperature: 0.4, maxOutputTokens: 8192 }
  };
  var r = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(cuerpo), muteHttpExceptions: true });
  var codigo = r.getResponseCode(), datos = {};
  try { datos = JSON.parse(r.getContentText() || '{}'); } catch (e) {}
  if (codigo !== 200) throw new Error('La IA respondió ' + codigo + (datos.error && datos.error.message ? ': ' + datos.error.message : '') + ' (modelo ' + modelo + ')');
  var partes = datos.candidates && datos.candidates[0] && datos.candidates[0].content && datos.candidates[0].content.parts;
  var texto_ = (partes || []).filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('').trim();
  if (!texto_) throw new Error('La IA no devolvió texto');
  return texto_;
}
function apiGeminiChat(d) {
  var msgs = (Array.isArray(d.messages) ? d.messages : []).slice(-20);
  if (!msgs.length) return error('No hay mensaje');
  var hoy = hoyLocal();
  var sistema = [
    'Sos el asistente del panel de reservas de Estudio Fotoinfantil (fotos de bebés y niños). Respondé en español rioplatense, corto y claro.',
    'Hoy es ' + hoy + ' (' + Utilities.formatDate(new Date(), TZ, 'EEEE') + '). Horarios base: lunes a viernes ' + leerAjustes().grillaSemana.join(', ') + '; sábados ' + leerAjustes().grillaSabado.join(', ') + '; domingos cerrado.',
    'Reservas (últimos 45 días y próximas, sin canceladas), en JSON: ' + JSON.stringify(datosParaIA()),
    'Si te piden un cambio, proponelo con UN bloque por acción exactamente así:',
    '```action',
    '{"tipo":"updateStatus","id":"T123456","newStatus":"CONFIRMADO","descripcion":"Confirmar la sesión de ..."}',
    '```',
    'Tipos permitidos: updateStatus (id, newStatus: PENDIENTE|CONFIRMADO|CANCELADO), reprogramar (id, newFecha AAAA-MM-DD, newHorario HH:MM),',
    'bloquearDia (fecha), desbloquearDia (fecha), bloquearSlot (fecha, horario), desbloquearSlot (fecha, horario), agregarSlotExtra (fecha, horario).',
    'Usá SOLO ids que estén en la lista. Nunca digas que algo ya se hizo: el panel le pide confirmación a la persona antes de cambiar nada.',
    'No inventes datos. Si falta información, preguntá.'
  ].join('\n');
  return { ok: true, reply: llamarGemini(sistema, msgs) };
}
function apiGeminiResumen() {
  var hoy = hoyLocal(), datos = datosParaIA();
  var deHoy = datos.filter(function (x) { return x.fecha === hoy; });
  var sistema = 'Sos el asistente de Estudio Fotoinfantil. Escribí un resumen de 2 o 3 oraciones, cálido y práctico, en español rioplatense. No uses listas.';
  var pedido = 'Hoy es ' + hoy + '. Sesiones de hoy: ' + JSON.stringify(deHoy) + '. Reservas próximas: ' + JSON.stringify(datos.filter(function (x) { return x.fecha > hoy; }).slice(0, 20)) +
    '. Contá cuántas sesiones hay hoy y a qué hora, y mencioná algo útil de los próximos días (por ejemplo señas pendientes).';
  return { ok: true, reply: llamarGemini(sistema, [{ role: 'user', content: pedido }]) };
}

// ════════════════════════════════════════════════════════════════════════════
// PARA EJECUTAR A MANO
// ════════════════════════════════════════════════════════════════════════════
function revisarInstalacion() {
  var out = [], p = props();
  try {
    var res = leerReservas();
    out.push('✓ Reservas: ' + res.filas.length + ' filas. Columnas: ' + res.enc.filter(String).join(', '));
    var ids = {}, rep = [];
    res.filas.forEach(function (f) { if (ids[f.id]) rep.push(f.id); ids[f.id] = true; });
    out.push(rep.length ? '⚠ IDs repetidos (revisar a mano): ' + rep.join(', ') : '✓ Sin IDs repetidos');
    var formulas = res.hoja.getDataRange().getFormulas(), conFormula = {};
    formulas.forEach(function (f, r) { if (r > 0) f.forEach(function (x, c) { if (x) { var n = res.enc[c] || ('columna ' + (c + 1)); conFormula[n] = (conFormula[n] || 0) + 1; } }); });
    var cf = Object.keys(conFormula);
    out.push(cf.length ? '· Columnas con fórmulas (el panel no las pisa): ' + cf.map(function (k) { return k + ' (' + conFormula[k] + ')'; }).join(', ') : '✓ Sin fórmulas en las reservas');
  } catch (e) { out.push('✗ Reservas: ' + mensaje(e)); }
  try { out.push('✓ Configuración: ' + leerConfig().lista.length + ' bloqueos/extras'); } catch (e) { out.push('✗ Configuración: ' + mensaje(e)); }
  asegurarHoja(H_SEG, ENC_SEG); asegurarHoja(H_REG, ENC_REG); crearAjustesSiFalta();
  out.push('✓ Pestañas nuevas listas: ' + [H_SEG, H_REG, H_AJ].join(', '));
  var aj = leerAjustes();
  out.push('· Seña: $' + aj.senaMonto + ' · plazo web: ' + aj.senaHoras + ' h · alias: ' + (aj.alias || '(falta cargarlo en Ajustes)'));
  out.push(p.getProperty('CODIGO_HASH') ? '✓ Código de acceso configurado' : (p.getProperty('CODIGO_INICIAL') ? '✓ CODIGO_INICIAL cargado (se cifra al primer ingreso)' : '· Sin código todavía: se crea desde el panel con «Olvidé el código» (llega un mail)'));
  out.push(emailsAdmin().length ? '✓ Correos para recuperar: ' + emailsAdmin().length : '✗ Falta EMAILS_ADMIN');
  out.push(p.getProperty('GEMINI_API_KEY') || p.getProperty('GEMINI_KEY') ? '✓ Clave de IA cargada (modelo ' + (p.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash') + ')' : '· Sin clave de IA (el asistente no va a responder)');
  var trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  out.push(trig.indexOf('copiaDiaria') >= 0 ? '✓ Copia diaria programada' : '· Copia diaria sin programar: ejecutá instalarCopiaDiaria()');
  if (trig.length) out.push('· Activadores del proyecto: ' + trig.join(', '));
  Logger.log(out.join('\n'));
  return out;
}
function probarMail() {
  var ok = avisarAdmins('Prueba del panel de reservas', 'Si te llegó este mail, la recuperación del código va a funcionar.');
  Logger.log(ok ? 'Mail de prueba enviado a: ' + emailsAdmin().join(', ') : 'Falta EMAILS_ADMIN en las propiedades del script');
}

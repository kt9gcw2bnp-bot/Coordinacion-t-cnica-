/**
 * PLATAFORMA DE COORDINACIÓN TÉCNICA — Backend (Google Apps Script)
 * Acceso solo para correos @mep.go.cr autorizados en la hoja "Usuarios".
 * Verificación por código de 6 dígitos enviado al correo institucional.
 *
 * Proyecto independiente en script.google.com (no necesita abrirse desde la hoja).
 */

// Hoja de datos en su Drive (Coordinacion_Tecnica_Datos)
const SHEET_ID = '1WZQUxle8VVDghZfEEylDKB7hJXTmyEwiSuT_F6vyvMU';
// Hoja con el inventario actual para cargarlo una sola vez (Inventario_Importacion)
const IMPORT_ID = '1W-S_bwTBGUbX16GkyPJk9eRHlnDF78TMI7TUatABJYE';

const CONFIG = {
  DOMINIO: 'mep.go.cr',
  ADMIN_INICIAL: 'luis.gonzalez.gonzalez@mep.go.cr',
  NOMBRE_ADMIN_INICIAL: 'Gerardo González González',
  NOMBRE_PLATAFORMA: 'Plataforma de Coordinación Técnica',
  MINUTOS_CODIGO: 10,
  MAX_INTENTOS: 5,
  HORAS_SESION: 8,
  SEGUNDOS_ENTRE_CODIGOS: 60
};

const ROLES = ['Administrador', 'Dirección', 'Docente técnico', 'Junta Administrativa', 'Contabilidad'];

const HOJAS = {
  Usuarios: ['Correo', 'Nombre', 'Rol', 'Activo', 'Creado', 'CreadoPor', 'UltimoIngreso'],
  Codigos: ['Correo', 'Hash', 'Expira', 'Intentos', 'Creado'],
  Sesiones: ['Token', 'Correo', 'Expira', 'Creado'],
  Bitacora: ['Fecha', 'Correo', 'Accion', 'Detalle'],
  Inventario: ['Placa', 'Descripcion', 'Marca', 'Modelo', 'Serie', 'Estado', 'Ubicacion', 'UbicacionReal', 'Especialidad',
               'ModoAdquisicion', 'TFA', 'Observaciones', 'Situacion', 'FechaBaja', 'MotivoBaja', 'ActaBaja', 'Actualizado', 'ActualizadoPor']
};

const ESTADOS = ['Bueno', 'Regular', 'Malo'];
const MODOS = ['Ley 7372', 'Donación MEP', 'Donación', 'Fondos Propios', 'Llave en mano DIEE', 'Otro'];
const MOTIVOS = ['Inservible', 'Obsoleto', 'Robo o hurto', 'Extravío', 'Traslado a otra institución', 'Otro'];

/* ============ INSTALACIÓN (ejecutar UNA vez desde el editor) ============ */
function setup() {
  Object.keys(HOJAS).forEach(nombre => hoja(nombre));
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SAL')) props.setProperty('SAL', Utilities.getUuid() + Utilities.getUuid());

  const admin = norm(CONFIG.ADMIN_INICIAL);
  if (!buscarUsuario(admin)) {
    hoja('Usuarios').appendRow([admin, CONFIG.NOMBRE_ADMIN_INICIAL, 'Administrador', 'SI', new Date(), 'setup', '']);
  }
  bitacora(admin, 'setup', 'Instalación inicial');
  Logger.log('Listo. Administrador: ' + admin);
}

/* ============ ENTRADA WEB ============ */
function doGet() {
  return json({ ok: true, servicio: CONFIG.NOMBRE_PLATAFORMA });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json({ ok: false, error: 'Solicitud inválida.' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (body.action) {
      case 'solicitarCodigo': return json(solicitarCodigo(body));
      case 'verificarCodigo': return json(verificarCodigo(body));
      case 'sesion':          return json(conSesion(body, u => ({ ok: true, usuario: publico(u) })));
      case 'cerrarSesion':    return json(cerrarSesion(body));
      case 'listarUsuarios':  return json(conSesion(body, () => listarUsuarios(), true));
      case 'guardarUsuario':  return json(conSesion(body, u => guardarUsuario(u, body.datos || {}), true));
      case 'invListar':       return json(conSesion(body, u => invListar(u)));
      case 'invGuardar':      return json(conSesion(body, u => invGuardar(u, body.datos || {})));
      case 'invBaja':         return json(conSesion(body, u => invBaja(u, body.datos || {}), true));
      case 'invRevertir':     return json(conSesion(body, u => invRevertir(u, body.datos || {}), true));
      default:                return json({ ok: false, error: 'Acción no reconocida.' });
    }
  } catch (err) {
    bitacora('', 'error', String(err && err.stack || err));
    return json({ ok: false, error: 'Error del servidor. Intente de nuevo.' });
  } finally {
    lock.releaseLock();
  }
}

/* ============ ACCESO ============ */
function solicitarCodigo(body) {
  const correo = norm(body.correo);
  const generico = { ok: true, mensaje: 'Si el correo está autorizado, recibirá un código de acceso en unos segundos.' };
  if (!correoValido(correo)) return { ok: false, error: 'Use su correo institucional @' + CONFIG.DOMINIO + '.' };

  const u = buscarUsuario(correo);
  if (!u || u.Activo !== 'SI') { bitacora(correo, 'codigo_denegado', 'No autorizado'); return generico; }

  const h = hoja('Codigos');
  const previos = filas('Codigos').filter(r => r.Correo === correo);
  const ahora = new Date();
  const reciente = previos.find(r => (ahora - new Date(r.Creado)) / 1000 < CONFIG.SEGUNDOS_ENTRE_CODIGOS);
  if (reciente) return { ok: false, error: 'Espere un minuto antes de pedir otro código.' };
  borrarFilas(h, previos);

  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  const expira = new Date(ahora.getTime() + CONFIG.MINUTOS_CODIGO * 60000);
  h.appendRow([correo, hash(correo + '|' + codigo), expira, 0, ahora]);

  MailApp.sendEmail({
    to: correo,
    subject: 'Código de acceso: ' + codigo + ' — ' + CONFIG.NOMBRE_PLATAFORMA,
    htmlBody:
      '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:480px;margin:auto;border:1px solid #D5DCE6">' +
      '<div style="background:#0B2A55;color:#fff;padding:16px 20px;border-bottom:4px solid #C9A227">' +
      '<b>' + CONFIG.NOMBRE_PLATAFORMA + '</b></div>' +
      '<div style="padding:20px;color:#1B2635">Estimado(a) ' + escapar(u.Nombre) + ':<br><br>' +
      'Su código de acceso es:<div style="font-size:32px;letter-spacing:8px;font-weight:bold;color:#0B2A55;margin:14px 0">' + codigo + '</div>' +
      'Vence en ' + CONFIG.MINUTOS_CODIGO + ' minutos. Si usted no lo solicitó, ignore este mensaje.</div></div>'
  });
  bitacora(correo, 'codigo_enviado', '');
  return generico;
}

function verificarCodigo(body) {
  const correo = norm(body.correo);
  const codigo = String(body.codigo || '').replace(/\D/g, '');
  const h = hoja('Codigos');
  const r = filas('Codigos').find(x => x.Correo === correo);
  if (!r || new Date(r.Expira) < new Date()) {
    if (r) borrarFilas(h, [r]);
    return { ok: false, error: 'Código vencido o inválido. Solicite uno nuevo.' };
  }
  if (hash(correo + '|' + codigo) !== r.Hash) {
    const intentos = Number(r.Intentos) + 1;
    if (intentos >= CONFIG.MAX_INTENTOS) {
      borrarFilas(h, [r]);
      bitacora(correo, 'codigo_bloqueado', 'Demasiados intentos');
      return { ok: false, error: 'Demasiados intentos. Solicite un código nuevo.' };
    }
    h.getRange(r._fila, 4).setValue(intentos);
    return { ok: false, error: 'Código incorrecto. Le quedan ' + (CONFIG.MAX_INTENTOS - intentos) + ' intentos.' };
  }
  borrarFilas(h, [r]);

  const u = buscarUsuario(correo);
  if (!u || u.Activo !== 'SI') return { ok: false, error: 'Su acceso no está habilitado.' };

  limpiarSesiones();
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const ahora = new Date();
  hoja('Sesiones').appendRow([hash(token), correo, new Date(ahora.getTime() + CONFIG.HORAS_SESION * 3600000), ahora]);
  hoja('Usuarios').getRange(u._fila, 7).setValue(ahora);
  bitacora(correo, 'ingreso', '');
  return { ok: true, token: token, usuario: publico(u) };
}

function conSesion(body, fn, soloAdmin) {
  const t = hash(String(body.token || ''));
  const s = filas('Sesiones').find(x => x.Token === t);
  if (!s || new Date(s.Expira) < new Date()) return { ok: false, sesion: false, error: 'Su sesión venció. Ingrese de nuevo.' };
  const u = buscarUsuario(s.Correo);
  if (!u || u.Activo !== 'SI') return { ok: false, sesion: false, error: 'Su acceso fue deshabilitado.' };
  if (soloAdmin && u.Rol !== 'Administrador') return { ok: false, error: 'Solo el administrador puede hacer esto.' };
  return fn(u);
}

function cerrarSesion(body) {
  const t = hash(String(body.token || ''));
  borrarFilas(hoja('Sesiones'), filas('Sesiones').filter(x => x.Token === t));
  return { ok: true };
}

/* ============ USUARIOS (solo administrador) ============ */
function listarUsuarios() {
  return { ok: true, roles: ROLES, usuarios: filas('Usuarios').map(publico) };
}

function guardarUsuario(admin, d) {
  const correo = norm(d.correo);
  if (!correoValido(correo)) return { ok: false, error: 'El correo debe ser @' + CONFIG.DOMINIO + '.' };
  const nombre = String(d.nombre || '').trim();
  if (!nombre) return { ok: false, error: 'Indique el nombre completo.' };
  const rol = ROLES.indexOf(d.rol) >= 0 ? d.rol : 'Docente técnico';
  const activo = d.activo === 'NO' ? 'NO' : 'SI';
  if (correo === admin.Correo && (activo === 'NO' || rol !== 'Administrador')) {
    return { ok: false, error: 'No puede quitarse a sí mismo el acceso de administrador.' };
  }
  const h = hoja('Usuarios');
  const u = buscarUsuario(correo);
  if (u) {
    h.getRange(u._fila, 2, 1, 3).setValues([[nombre, rol, activo]]);
  } else {
    h.appendRow([correo, nombre, rol, activo, new Date(), admin.Correo, '']);
  }
  if (activo === 'NO') borrarFilas(hoja('Sesiones'), filas('Sesiones').filter(x => x.Correo === correo));
  bitacora(admin.Correo, u ? 'usuario_editado' : 'usuario_agregado', correo + ' · ' + rol + ' · ' + activo);
  return listarUsuarios();
}

/* ============ INVENTARIO ============ */
function puedeEditar(u) { return u.Rol === 'Administrador' || u.Rol === 'Docente técnico'; }

function invPub(r) {
  const o = {};
  HOJAS.Inventario.forEach(c => {
    const v = r[c];
    o[c] = v instanceof Date ? Utilities.formatDate(v, 'America/Costa_Rica', 'yyyy-MM-dd') : String(v == null ? '' : v).trim();
  });
  return o;
}

function invListar(u) {
  return {
    ok: true, puedeEditar: puedeEditar(u), esAdmin: u.Rol === 'Administrador',
    items: filas('Inventario').map(invPub),
    catalogos: { estados: ESTADOS, modos: MODOS, motivos: MOTIVOS }
  };
}

function invBuscar(placa) { return filas('Inventario').find(r => String(r.Placa).trim() === placa) || null; }

function invEscribir(fila, o, quien) {
  o.Actualizado = Utilities.formatDate(new Date(), 'America/Costa_Rica', 'yyyy-MM-dd HH:mm');
  o.ActualizadoPor = quien;
  const vals = [HOJAS.Inventario.map(c => String(o[c] == null ? '' : o[c]))];
  const h = hoja('Inventario');
  if (fila) h.getRange(fila, 1, 1, vals[0].length).setValues(vals);
  else { capacidad(h, 1); h.getRange(h.getLastRow() + 1, 1, 1, vals[0].length).setValues(vals); }
  return invPub(o);
}

function txt(v, max) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max || 200); }

function invGuardar(u, d) {
  if (!puedeEditar(u)) return { ok: false, error: 'Su rol solo permite consultar el inventario.' };
  const placa = txt(d.Placa, 20);
  if (!/^\d{4}-\d{1,6}$/.test(placa)) return { ok: false, error: 'La placa debe tener el formato 6526-0000.' };
  if (!txt(d.Descripcion)) return { ok: false, error: 'Indique la descripción del bien.' };
  const ex = invBuscar(placa);
  if (d.nuevo && ex) return { ok: false, error: 'Ya existe un bien con la placa ' + placa + (ex.Situacion === 'Baja' ? ' (dado de baja).' : '.') };
  if (!d.nuevo && !ex) return { ok: false, error: 'No se encontró la placa ' + placa + '.' };
  if (ex && ex.Situacion === 'Baja') return { ok: false, error: 'Un bien dado de baja no se puede editar.' };
  const o = ex ? invPub(ex) : { Situacion: 'Activo', FechaBaja: '', MotivoBaja: '', ActaBaja: '' };
  o.Placa = placa;
  ['Descripcion', 'Marca', 'Modelo', 'Serie', 'Ubicacion', 'UbicacionReal', 'Especialidad', 'TFA'].forEach(c => o[c] = txt(d[c]));
  o.Observaciones = txt(d.Observaciones, 500);
  o.Estado = ESTADOS.indexOf(d.Estado) >= 0 ? d.Estado : '';
  o.ModoAdquisicion = MODOS.indexOf(d.ModoAdquisicion) >= 0 ? d.ModoAdquisicion : txt(d.ModoAdquisicion);
  const item = invEscribir(ex ? ex._fila : null, o, u.Correo);
  bitacora(u.Correo, ex ? 'inv_editado' : 'inv_agregado', placa + ' · ' + o.Descripcion);
  return { ok: true, item: item };
}

function invBaja(u, d) {
  const ex = invBuscar(txt(d.Placa, 20));
  if (!ex) return { ok: false, error: 'No se encontró la placa.' };
  if (ex.Situacion === 'Baja') return { ok: false, error: 'Ese bien ya está dado de baja.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.FechaBaja || ''))) return { ok: false, error: 'Indique la fecha de la baja.' };
  if (MOTIVOS.indexOf(d.MotivoBaja) < 0) return { ok: false, error: 'Seleccione el motivo.' };
  if (!txt(d.ActaBaja)) return { ok: false, error: 'Indique el número de acta.' };
  const o = invPub(ex);
  o.Situacion = 'Baja'; o.FechaBaja = d.FechaBaja; o.MotivoBaja = d.MotivoBaja; o.ActaBaja = txt(d.ActaBaja, 30);
  if (txt(d.TFA)) o.TFA = txt(d.TFA, 30);
  const item = invEscribir(ex._fila, o, u.Correo);
  bitacora(u.Correo, 'inv_baja', o.Placa + ' · acta ' + o.ActaBaja + ' · ' + o.MotivoBaja);
  return { ok: true, item: item };
}

function invRevertir(u, d) {
  const ex = invBuscar(txt(d.Placa, 20));
  if (!ex || ex.Situacion !== 'Baja') return { ok: false, error: 'Ese bien no está dado de baja.' };
  const o = invPub(ex);
  o.Situacion = 'Activo'; o.FechaBaja = ''; o.MotivoBaja = ''; o.ActaBaja = '';
  const item = invEscribir(ex._fila, o, u.Correo);
  bitacora(u.Correo, 'inv_baja_revertida', o.Placa);
  return { ok: true, item: item };
}

/** Ejecutar UNA vez desde el editor: carga el inventario actual (activos y bajas). No duplica placas. */
function importarInventarioInicial() {
  const origen = SpreadsheetApp.openById(IMPORT_ID).getSheets()[0].getDataRange().getDisplayValues();
  const enc = origen.shift();
  const yaEstan = {};
  filas('Inventario').forEach(r => yaEstan[String(r.Placa).trim()] = true);
  const ahora = Utilities.formatDate(new Date(), 'America/Costa_Rica', 'yyyy-MM-dd HH:mm');
  const nuevas = [];
  origen.forEach(f => {
    const o = {}; enc.forEach((c, j) => o[c] = String(f[j] || '').trim());
    if (!o.Placa || yaEstan[o.Placa]) return;
    yaEstan[o.Placa] = true;
    o.Actualizado = ahora; o.ActualizadoPor = 'importación';
    nuevas.push(HOJAS.Inventario.map(c => o[c] || ''));
  });
  const h = hoja('Inventario');
  if (nuevas.length) capacidad(h, nuevas.length);
  if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, HOJAS.Inventario.length).setValues(nuevas);
  bitacora('', 'inv_importacion', nuevas.length + ' bienes');
  Logger.log('Importados: ' + nuevas.length + ' bienes.');
}

/* ============ UTILIDADES ============ */
// Asegura filas libres (en formato texto) al final de la hoja de inventario
function capacidad(h, n) {
  const faltan = h.getLastRow() + n - h.getMaxRows();
  if (faltan > 0) {
    h.insertRowsAfter(h.getMaxRows(), faltan + 200);
    h.getRange(1, 1, h.getMaxRows(), HOJAS.Inventario.length).setNumberFormat('@');
  }
}

let _ss = null;
function hoja(n) {
  if (!_ss) _ss = SpreadsheetApp.openById(SHEET_ID);
  let h = _ss.getSheetByName(n);
  if (!h && HOJAS[n]) {
    h = _ss.insertSheet(n);
    if (n === 'Inventario') h.getRange(1, 1, h.getMaxRows(), HOJAS[n].length).setNumberFormat('@');
  }
  if (h && h.getLastRow() === 0 && HOJAS[n]) {
    h.appendRow(HOJAS[n]);
    h.getRange(1, 1, 1, HOJAS[n].length).setFontWeight('bold').setBackground('#0B2A55').setFontColor('#FFFFFF');
    h.setFrozenRows(1);
  }
  return h;
}

function filas(n) {
  const v = hoja(n).getDataRange().getValues();
  const enc = v.shift() || [];
  return v.map((f, i) => {
    const o = { _fila: i + 2 };
    enc.forEach((c, j) => o[c] = f[j]);
    return o;
  });
}

function borrarFilas(h, lista) {
  lista.map(r => r._fila).sort((a, b) => b - a).forEach(f => h.deleteRow(f));
}

function buscarUsuario(correo) {
  return filas('Usuarios').find(r => norm(r.Correo) === correo) || null;
}

function publico(u) {
  return {
    correo: norm(u.Correo), nombre: u.Nombre, rol: u.Rol, activo: u.Activo,
    ultimoIngreso: u.UltimoIngreso ? Utilities.formatDate(new Date(u.UltimoIngreso), 'America/Costa_Rica', 'dd/MM/yyyy HH:mm') : ''
  };
}

function limpiarSesiones() {
  const ahora = new Date();
  borrarFilas(hoja('Sesiones'), filas('Sesiones').filter(x => new Date(x.Expira) < ahora));
}

function norm(c) { return String(c || '').trim().toLowerCase(); }

function correoValido(c) {
  return new RegExp('^[a-z0-9._%+-]+@' + CONFIG.DOMINIO.replace(/\./g, '\\.') + '$').test(c);
}

function hash(s) {
  const sal = PropertiesService.getScriptProperties().getProperty('SAL') || '';
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function escapar(s) {
  return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function bitacora(correo, accion, detalle) {
  try { hoja('Bitacora').appendRow([new Date(), correo, accion, detalle]); } catch (e) {}
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

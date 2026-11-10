/**
 * PLATAFORMA DE COORDINACIÓN TÉCNICA — Backend (Google Apps Script)
 * Acceso solo para correos @mep.go.cr autorizados en la hoja "Usuarios".
 * Verificación por código de 6 dígitos enviado al correo institucional.
 *
 * Proyecto independiente en script.google.com (no necesita abrirse desde la hoja).
 */

// Hoja de datos en su Drive (Coordinacion_Tecnica_Datos)
const SHEET_ID = '1WZQUxle8VVDghZfEEylDKB7hJXTmyEwiSuT_F6vyvMU';

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
  Bitacora: ['Fecha', 'Correo', 'Accion', 'Detalle']
};

/* ============ INSTALACIÓN (ejecutar UNA vez desde el editor) ============ */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(HOJAS).forEach(nombre => {
    let h = ss.getSheetByName(nombre);
    if (!h) h = ss.insertSheet(nombre);
    if (h.getLastRow() === 0) {
      h.appendRow(HOJAS[nombre]);
      h.getRange(1, 1, 1, HOJAS[nombre].length).setFontWeight('bold').setBackground('#0B2A55').setFontColor('#FFFFFF');
      h.setFrozenRows(1);
    }
  });
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

/* ============ UTILIDADES ============ */
function hoja(n) { return SpreadsheetApp.openById(SHEET_ID).getSheetByName(n); }

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

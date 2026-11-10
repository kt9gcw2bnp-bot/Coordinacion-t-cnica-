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
               'ModoAdquisicion', 'TFA', 'Observaciones', 'Situacion', 'FechaBaja', 'MotivoBaja', 'ActaBaja', 'Actualizado', 'ActualizadoPor',
               'Responsable', 'FechaIngreso', 'ConsecutivoBaja'],
  Movimientos: ['Consecutivo', 'Fecha', 'Tipo', 'Placa', 'Descripcion', 'Estado', 'DeUbicacion', 'AUbicacion', 'DeResponsable',
                'AResponsable', 'Causal', 'Detalle', 'RegistradoPor', 'Registrado'],
  Config: ['Clave', 'Valor']
};

const ESTADOS = ['Bueno', 'Regular', 'Malo'];
const MODOS = ['Ley 7372', 'Donación MEP', 'Donación', 'Fondos Propios', 'Llave en mano DIEE', 'Otro'];

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
      case 'invMovimientos':  return json(conSesion(body, u => invMovimientos(u, body.datos || {})));
      case 'invReporte':      return json(conSesion(body, u => invReporte(u, body.datos || {})));
      case 'invGuardar':      return json(conSesion(body, u => invGuardar(u, body.datos || {}), true));
      case 'invTraslado':     return json(conSesion(body, u => invTraslado(u, body.datos || {}), true));
      case 'invBaja':         return json(conSesion(body, u => invBaja(u, body.datos || {}), true));
      case 'invAnularBaja':   return json(conSesion(body, u => invAnularBaja(u, body.datos || {}), true));
      case 'invEliminar':     return json(conSesion(body, u => invEliminar(u, body.datos || {}), true));
      case 'invCorregirPlaca':return json(conSesion(body, u => invCorregirPlaca(u, body.datos || {}), true));
      case 'invImportar':     return json(conSesion(body, u => invImportar(u, body.datos || {}), true));
      case 'cfgGuardar':      return json(conSesion(body, u => cfgGuardar(u, body.datos || {}), true));
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

/* ============ INVENTARIO ============
 * Base normativa aplicada:
 *  - Ley General de Control Interno N.° 8292, art. 8 a): proteger y conservar el patrimonio público.
 *  - Normas de Control Interno para el Sector Público (N-2-2009-CO-DFOE), norma 4.3 Protección y conservación del patrimonio.
 *  - Instructivo MEP (2021) "Temas adicionales a considerar en la administración de los inventarios de los centros educativos":
 *    Libro de Inventarios (T.F.A.), acuerdo de Junta, actas en dos originales, placas no reutilizables.
 *  - Reglamento 40797-H (referencia): art. 7 deberes del funcionario a cargo, art. 42 traslados, arts. 22, 25, 27, 30 bajas.
 */

// Requisitos por causal de baja (campos obligatorios además de los comunes)
const CAUSALES = {
  'Inservibilidad (destrucción)':      { junta: true,  campos: ['Empresa', 'PermisoSalud'] },
  'Donación':                          { junta: true,  campos: ['Donatario', 'CedulaJuridica'] },
  'Venta':                             { junta: true,  campos: ['Comprador', 'CedulaComprador', 'Monto', 'Comprobante'] },
  'Caso fortuito o fuerza mayor':      { junta: true,  campos: ['DescripcionHecho'] },
  'Hurto o robo':                      { junta: true,  campos: ['DenunciaOIJ', 'FechaDenuncia'] },
  'Pérdida o desaparición':            { junta: true,  campos: ['OficioJunta'] },
  'Exclusión de bien no patrimoniable':{ junta: false, campos: [] },
  'Retiro por la FOD (PRONIE)':        { junta: false, campos: [] }
};
const ROLES_LECTURA_TOTAL = ['Administrador', 'Dirección', 'Junta Administrativa', 'Contabilidad'];
const CFG_CLAVES = ['Institucion', 'CodigoPresupuestario', 'DRE', 'Circuito', 'Canton', 'Director', 'PresidenteJunta', 'CoordinadorTecnico'];

function esAdminU(u) { return u.Rol === 'Administrador'; }
function veTodo(u) { return ROLES_LECTURA_TOTAL.indexOf(u.Rol) >= 0; }
function hoyCR() { return Utilities.formatDate(new Date(), 'America/Costa_Rica', 'yyyy-MM-dd'); }
function ahoraCR() { return Utilities.formatDate(new Date(), 'America/Costa_Rica', 'yyyy-MM-dd HH:mm'); }
function txt(v, max) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max || 200); }
function fechaOk(f) { return /^\d{4}-\d{2}-\d{2}$/.test(String(f || '')); }

function invPub(r) {
  const o = {};
  HOJAS.Inventario.forEach(c => {
    const v = r[c];
    o[c] = v instanceof Date ? Utilities.formatDate(v, 'America/Costa_Rica', 'yyyy-MM-dd') : String(v == null ? '' : v).trim();
  });
  o.Responsable = o.Responsable.toLowerCase();
  return o;
}

function leerConfig() {
  const c = {};
  filas('Config').forEach(r => c[r.Clave] = String(r.Valor == null ? '' : r.Valor));
  return c;
}

function nombresUsuarios() {
  const m = {};
  filas('Usuarios').forEach(r => m[norm(r.Correo)] = r.Nombre);
  return m;
}

function invListar(u) {
  const todos = filas('Inventario').map(invPub);
  const items = veTodo(u) ? todos : todos.filter(i => i.Responsable === u.Correo);
  const res = {
    ok: true, version: 5, esAdmin: esAdminU(u), veTodo: veTodo(u), yo: u.Correo,
    items: items, config: leerConfig(), nombres: nombresUsuarios(),
    catalogos: { estados: ESTADOS, modos: MODOS, causales: Object.keys(CAUSALES), requisitos: CAUSALES }
  };
  if (esAdminU(u)) res.usuarios = filas('Usuarios').filter(r => r.Activo === 'SI').map(r => ({ correo: norm(r.Correo), nombre: r.Nombre, rol: r.Rol }));
  return res;
}

function invBuscarFila(placa, todas) { return (todas || filas('Inventario')).find(r => String(r.Placa).trim() === placa) || null; }

function invEscribir(fila, o, quien) {
  o.Actualizado = ahoraCR();
  o.ActualizadoPor = quien;
  const h = hoja('Inventario');
  const vals = [encabezado(h).map(c => String(o[c] == null ? '' : o[c]))];
  if (fila) h.getRange(fila, 1, 1, vals[0].length).setValues(vals);
  else { capacidad(h, 1); h.getRange(h.getLastRow() + 1, 1, 1, vals[0].length).setValues(vals); }
  return invPub(o);
}

function registrarMov(m, quien) {
  const h = hoja('Movimientos');
  h.appendRow(encabezado(h).map(c => c === 'RegistradoPor' ? quien : c === 'Registrado' ? ahoraCR() : String(m[c] == null ? '' : m[c])));
}

function siguienteConsecutivo(prefijo, fecha) {
  const anio = String(fecha || hoyCR()).slice(0, 4);
  const base = prefijo + '-' + anio + '-';
  let max = 0;
  filas('Movimientos').forEach(r => {
    const c = String(r.Consecutivo || '');
    if (c.indexOf(base) === 0) max = Math.max(max, Number(c.slice(base.length)) || 0);
  });
  return base + ('000' + (max + 1)).slice(-4);
}

/* Alta y modificación de datos descriptivos (solo Administrador / Coordinación Técnica).
   La ubicación y el responsable solo cambian mediante traslado, para dejar trazabilidad. */
function invGuardar(u, d) {
  const placa = txt(d.Placa, 20);
  if (!/^\d{4}-\d{1,6}$/.test(placa)) return { ok: false, error: 'La placa debe tener el formato 6526-0000.' };
  if (!txt(d.Descripcion)) return { ok: false, error: 'Indique la descripción del bien.' };
  const ex = invBuscarFila(placa);
  if (d.nuevo && ex) return { ok: false, error: 'La placa ' + placa + ' ya existe' + (ex.Situacion === 'Baja' ? ' y fue dada de baja. Las placas dadas de baja no se reutilizan.' : '.') };
  if (!d.nuevo && !ex) return { ok: false, error: 'No se encontró la placa ' + placa + '.' };
  if (ex && ex.Situacion === 'Baja') return { ok: false, error: 'Un bien dado de baja no se puede modificar.' };
  const antes = ex ? invPub(ex) : null;
  const o = antes ? Object.assign({}, antes) : { Situacion: 'Activo', FechaBaja: '', MotivoBaja: '', ActaBaja: '', ConsecutivoBaja: '' };
  o.Placa = placa;
  ['Descripcion', 'Marca', 'Modelo', 'Serie', 'Especialidad', 'TFA'].forEach(c => o[c] = txt(d[c]));
  o.Observaciones = txt(d.Observaciones, 500);
  o.Estado = ESTADOS.indexOf(d.Estado) >= 0 ? d.Estado : '';
  o.ModoAdquisicion = MODOS.indexOf(d.ModoAdquisicion) >= 0 ? d.ModoAdquisicion : txt(d.ModoAdquisicion);
  if (!ex) {
    o.Ubicacion = txt(d.Ubicacion); o.UbicacionReal = txt(d.UbicacionReal);
    o.Responsable = norm(d.Responsable);
    if (o.Responsable && !usuarioActivo(o.Responsable)) return { ok: false, error: 'El responsable debe ser un usuario activo.' };
    o.FechaIngreso = fechaOk(d.FechaIngreso) ? d.FechaIngreso : hoyCR();
  }
  const item = invEscribir(ex ? ex._fila : null, o, u.Correo);
  if (!ex) {
    registrarMov({ Fecha: o.FechaIngreso, Tipo: 'Alta', Placa: placa, Descripcion: o.Descripcion, Estado: o.Estado,
      AUbicacion: o.Ubicacion, AResponsable: o.Responsable, Causal: o.ModoAdquisicion, Detalle: 'T.F.A. ' + (o.TFA || '—') }, u.Correo);
  } else {
    const cambios = ['Descripcion', 'Marca', 'Modelo', 'Serie', 'Estado', 'Especialidad', 'ModoAdquisicion', 'TFA', 'Observaciones']
      .filter(c => String(antes[c]) !== String(o[c])).map(c => c + ': "' + antes[c] + '" → "' + o[c] + '"');
    if (cambios.length) registrarMov({ Fecha: hoyCR(), Tipo: 'Modificación', Placa: placa, Descripcion: o.Descripcion, Estado: o.Estado,
      DeUbicacion: o.Ubicacion, AUbicacion: o.Ubicacion, DeResponsable: o.Responsable, AResponsable: o.Responsable, Detalle: cambios.join('; ') }, u.Correo);
  }
  bitacora(u.Correo, ex ? 'inv_editado' : 'inv_agregado', placa);
  return { ok: true, item: item };
}

function usuarioActivo(correo) { const x = buscarUsuario(norm(correo)); return x && x.Activo === 'SI' ? x : null; }

/* Traslado / asignación de custodia (Reglamento 40797-H art. 42; art. 7: responsabilidad del funcionario a cargo) */
function invTraslado(u, d) {
  const placas = (d.Placas || []).map(p => txt(p, 20)).filter(Boolean);
  if (!placas.length) return { ok: false, error: 'Seleccione al menos un bien.' };
  if (!fechaOk(d.Fecha)) return { ok: false, error: 'Indique la fecha del traslado.' };
  if (!txt(d.Motivo)) return { ok: false, error: 'Indique la justificación del traslado.' };
  const cambiaResp = Object.prototype.hasOwnProperty.call(d, 'AResponsable');
  const nuevoResp = norm(d.AResponsable);
  if (cambiaResp && nuevoResp && !usuarioActivo(nuevoResp)) return { ok: false, error: 'El nuevo responsable debe ser un usuario activo.' };
  if (!txt(d.AUbicacion) && !cambiaResp) return { ok: false, error: 'Indique la nueva ubicación o el nuevo responsable.' };
  const todas = filas('Inventario');
  const sel = placas.map(p => invBuscarFila(p, todas));
  const falta = placas.filter((p, k) => !sel[k] || sel[k].Situacion === 'Baja');
  if (falta.length) return { ok: false, error: 'No se pueden trasladar (no existen o están de baja): ' + falta.join(', ') };
  const cons = siguienteConsecutivo('TR', d.Fecha);
  const items = sel.map(r => {
    const o = invPub(r), de = { u: o.Ubicacion, r: o.Responsable };
    if (txt(d.AUbicacion)) { o.Ubicacion = txt(d.AUbicacion); o.UbicacionReal = txt(d.AUbicacionReal) || o.UbicacionReal; }
    if (cambiaResp) o.Responsable = nuevoResp;
    const it = invEscribir(r._fila, o, u.Correo);
    registrarMov({ Consecutivo: cons, Fecha: d.Fecha, Tipo: 'Traslado', Placa: o.Placa, Descripcion: o.Descripcion, Estado: o.Estado,
      DeUbicacion: de.u, AUbicacion: o.Ubicacion, DeResponsable: de.r, AResponsable: o.Responsable, Detalle: txt(d.Motivo, 500) }, u.Correo);
    return it;
  });
  bitacora(u.Correo, 'inv_traslado', cons + ' · ' + placas.length + ' bienes');
  return { ok: true, consecutivo: cons, items: items };
}

/* Baja por causal (Instructivo MEP 2021; Reglamento 40797-H arts. 20-30 como referencia) */
function invBaja(u, d) {
  const placas = (d.Placas || []).map(p => txt(p, 20)).filter(Boolean);
  const causal = d.Causal;
  const req = CAUSALES[causal];
  if (!placas.length) return { ok: false, error: 'Seleccione al menos un bien.' };
  if (!req) return { ok: false, error: 'Seleccione la causal de la baja.' };
  if (!fechaOk(d.Fecha)) return { ok: false, error: 'Indique la fecha de la baja.' };
  if (!txt(d.Acta)) return { ok: false, error: 'Indique el número de acta.' };
  if (!txt(d.AnotacionTFA)) return { ok: false, error: 'Indique el tomo/folio/asiento donde se anota la baja en el Libro de Inventarios.' };
  if (req.junta && (!txt(d.SesionJunta) || !txt(d.AcuerdoJunta))) return { ok: false, error: 'Esta causal requiere el número de sesión y de acuerdo de la Junta Administrativa.' };
  const faltan = req.campos.filter(c => !txt(d[c]));
  if (faltan.length) return { ok: false, error: 'Faltan datos obligatorios para esta causal: ' + faltan.join(', ') + '.' };
  const todas = filas('Inventario');
  const sel = placas.map(p => invBuscarFila(p, todas));
  const mal = placas.filter((p, k) => !sel[k] || sel[k].Situacion === 'Baja');
  if (mal.length) return { ok: false, error: 'No se pueden dar de baja (no existen o ya están de baja): ' + mal.join(', ') };
  const detalle = {};
  ['SesionJunta', 'AcuerdoJunta', 'FechaAcuerdo', 'AnotacionTFA', 'Observaciones'].concat(req.campos).forEach(c => detalle[c] = txt(d[c], 500));
  const cons = siguienteConsecutivo('BJ', d.Fecha);
  const items = sel.map(r => {
    const o = invPub(r);
    o.Situacion = 'Baja'; o.FechaBaja = d.Fecha; o.MotivoBaja = causal; o.ActaBaja = txt(d.Acta, 30); o.ConsecutivoBaja = cons;
    const it = invEscribir(r._fila, o, u.Correo);
    registrarMov({ Consecutivo: cons, Fecha: d.Fecha, Tipo: 'Baja', Placa: o.Placa, Descripcion: o.Descripcion, Estado: o.Estado,
      DeUbicacion: o.Ubicacion, DeResponsable: o.Responsable, Causal: causal,
      Detalle: JSON.stringify(Object.assign({ Acta: o.ActaBaja, TFA: o.TFA, Marca: o.Marca, Modelo: o.Modelo, Serie: o.Serie }, detalle)) }, u.Correo);
    return it;
  });
  bitacora(u.Correo, 'inv_baja', cons + ' · ' + causal + ' · acta ' + d.Acta + ' · ' + placas.length + ' bienes');
  return { ok: true, consecutivo: cons, items: items };
}

/* Anulación de una baja registrada por error: queda como contra-asiento con justificación */
function invAnularBaja(u, d) {
  const r = invBuscarFila(txt(d.Placa, 20));
  if (!r || r.Situacion !== 'Baja') return { ok: false, error: 'Ese bien no está dado de baja.' };
  if (txt(d.Justificacion).length < 10) return { ok: false, error: 'Escriba la justificación de la anulación (mínimo 10 caracteres).' };
  const o = invPub(r), previa = o.ConsecutivoBaja;
  o.Situacion = 'Activo'; o.FechaBaja = ''; o.MotivoBaja = ''; o.ActaBaja = ''; o.ConsecutivoBaja = '';
  const it = invEscribir(r._fila, o, u.Correo);
  registrarMov({ Consecutivo: siguienteConsecutivo('AN'), Fecha: hoyCR(), Tipo: 'Anulación de baja', Placa: o.Placa, Descripcion: o.Descripcion,
    Estado: o.Estado, AUbicacion: o.Ubicacion, AResponsable: o.Responsable, Detalle: 'Anula ' + previa + '. ' + txt(d.Justificacion, 500) }, u.Correo);
  bitacora(u.Correo, 'inv_baja_anulada', o.Placa + ' · ' + previa);
  return { ok: true, item: it };
}

/* El funcionario a cargo informa cambios en el estado o uso del bien (Reglamento 40797-H art. 7 k) */
function invReporte(u, d) {
  const r = invBuscarFila(txt(d.Placa, 20));
  if (!r || r.Situacion === 'Baja') return { ok: false, error: 'Bien no encontrado.' };
  const o = invPub(r);
  if (o.Responsable !== u.Correo && !esAdminU(u)) return { ok: false, error: 'Solo la persona responsable del bien puede reportarlo.' };
  if (ESTADOS.indexOf(d.EstadoReportado) < 0) return { ok: false, error: 'Indique el estado actual del bien.' };
  if (!txt(d.Detalle)) return { ok: false, error: 'Describa la situación.' };
  registrarMov({ Fecha: hoyCR(), Tipo: 'Reporte del responsable', Placa: o.Placa, Descripcion: o.Descripcion, Estado: d.EstadoReportado,
    DeUbicacion: o.Ubicacion, DeResponsable: o.Responsable, Detalle: txt(d.Detalle, 500) }, u.Correo);
  bitacora(u.Correo, 'inv_reporte', o.Placa);
  return { ok: true };
}

function invMovimientos(u, d) {
  let movs = filas('Movimientos').map(r => { const o = {}; HOJAS.Movimientos.forEach(c => o[c] = r[c] instanceof Date ? Utilities.formatDate(r[c], 'America/Costa_Rica', 'yyyy-MM-dd') : String(r[c] == null ? '' : r[c])); return o; });
  if (d.Placa) movs = movs.filter(m => m.Placa === txt(d.Placa, 20));
  if (d.Consecutivo) movs = movs.filter(m => m.Consecutivo === txt(d.Consecutivo, 20));
  if (!d.Placa && !d.Consecutivo && !veTodo(u)) return { ok: false, error: 'Sin permiso.' };
  if (!veTodo(u)) {
    const mias = {}; filas('Inventario').forEach(r => { if (norm(r.Responsable) === u.Correo) mias[String(r.Placa).trim()] = true; });
    movs = movs.filter(m => mias[m.Placa]);
  }
  return { ok: true, movimientos: movs };
}

function cfgGuardar(u, d) {
  const h = hoja('Config'); const actual = filas('Config');
  CFG_CLAVES.forEach(k => {
    const v = txt(d[k], 150); const r = actual.find(x => x.Clave === k);
    if (r) h.getRange(r._fila, 2).setValue(v); else h.appendRow([k, v]);
  });
  bitacora(u.Correo, 'config', 'Datos institucionales');
  return { ok: true, config: leerConfig() };
}

/* Eliminación de registros erróneos (bienes que nunca debieron registrarse).
   Deja constancia en Movimientos con la justificación y una copia de los datos eliminados. */
function invEliminar(u, d) {
  const placas = (d.Placas || (d.Placa ? [d.Placa] : [])).map(p => txt(p, 20)).filter(Boolean);
  const just = txt(d.Justificacion, 500);
  if (!placas.length) return { ok: false, error: 'Seleccione al menos un bien.' };
  if (just.length < 10) return { ok: false, error: 'Escriba la justificación (mínimo 10 caracteres).' };
  const h = hoja('Inventario'), todas = filas('Inventario');
  const sel = placas.map(p => invBuscarFila(p, todas));
  const noEsta = placas.filter((p, k) => !sel[k]);
  if (noEsta.length) return { ok: false, error: 'No se encontraron: ' + noEsta.join(', ') };
  const deBaja = placas.filter((p, k) => sel[k].Situacion === 'Baja');
  if (deBaja.length) return { ok: false, error: 'Estos bienes están dados de baja; primero anule la baja: ' + deBaja.join(', ') };
  const cons = siguienteConsecutivo('EL', hoyCR());
  sel.forEach(r => {
    const o = invPub(r);
    registrarMov({ Consecutivo: cons, Fecha: hoyCR(), Tipo: 'Eliminación de registro erróneo', Placa: o.Placa, Descripcion: o.Descripcion, Estado: o.Estado,
      DeUbicacion: o.Ubicacion, DeResponsable: o.Responsable,
      Detalle: just + ' · Datos eliminados: ' + JSON.stringify({ Marca: o.Marca, Modelo: o.Modelo, Serie: o.Serie, Modo: o.ModoAdquisicion, TFA: o.TFA, Ingreso: o.FechaIngreso }) }, u.Correo);
  });
  sel.map(r => r._fila).sort((a, b) => b - a).forEach(f => h.deleteRow(f));
  bitacora(u.Correo, 'inv_eliminado', cons + ' · ' + placas.join(', '));
  return { ok: true, consecutivo: cons, eliminadas: placas };
}

/* Corrección de una placa mal digitada: conserva el historial del bien */
function invCorregirPlaca(u, d) {
  const placa = txt(d.Placa, 20), nueva = txt(d.PlacaNueva, 20).replace(/\s+/g, '');
  const just = txt(d.Justificacion, 500);
  if (!/^\d{4}-\d{1,6}$/.test(nueva)) return { ok: false, error: 'La placa correcta debe tener el formato 6526-0000.' };
  if (nueva === placa) return { ok: false, error: 'La placa nueva es igual a la actual.' };
  if (just.length < 10) return { ok: false, error: 'Escriba la justificación (mínimo 10 caracteres).' };
  const todas = filas('Inventario'), r = invBuscarFila(placa, todas);
  if (!r) return { ok: false, error: 'No se encontró la placa ' + placa + '.' };
  if (r.Situacion === 'Baja') return { ok: false, error: 'No se puede corregir la placa de un bien dado de baja.' };
  if (invBuscarFila(nueva, todas)) return { ok: false, error: 'La placa ' + nueva + ' ya está registrada.' };
  const o = invPub(r); o.Placa = nueva;
  const item = invEscribir(r._fila, o, u.Correo);
  const hm = hoja('Movimientos'), em = encabezado(hm), cp = em.indexOf('Placa') + 1;
  filas('Movimientos').forEach(m => { if (String(m.Placa) === placa) hm.getRange(m._fila, cp).setValue(nueva); });
  registrarMov({ Fecha: hoyCR(), Tipo: 'Corrección de placa', Placa: nueva, Descripcion: o.Descripcion, Estado: o.Estado,
    AUbicacion: o.Ubicacion, AResponsable: o.Responsable, Detalle: 'Placa anterior: ' + placa + '. ' + just }, u.Correo);
  bitacora(u.Correo, 'inv_placa_corregida', placa + ' → ' + nueva);
  return { ok: true, item: item, anterior: placa };
}

/* Importación desde Excel (solo Administrador). No duplica placas ni reactiva placas dadas de baja. */
function invImportar(u, d) {
  const entrada = d.Filas || [];
  const op = d.Opciones || {};
  const tipo = d.Tipo === 'Baja' ? 'Baja' : 'Activo';
  if (!entrada.length) return { ok: false, error: 'El archivo no tiene filas para importar.' };
  if (entrada.length > 400) return { ok: false, error: 'Envíe como máximo 400 filas por lote.' };
  const resp = norm(op.Responsable);
  if (tipo === 'Activo' && resp && !usuarioActivo(resp)) return { ok: false, error: 'El responsable debe ser un usuario activo.' };
  if (tipo === 'Baja' && (!fechaOk(op.FechaBaja) || !CAUSALES[op.Causal])) return { ok: false, error: 'Para bajas históricas indique la fecha y la causal.' };
  const h = hoja('Inventario'), enc = encabezado(h);
  const existe = {};
  filas('Inventario').forEach(r => existe[String(r.Placa).trim()] = true);
  const fechaIng = fechaOk(op.FechaIngreso) ? op.FechaIngreso : hoyCR();
  const cons = d.Consecutivo && /^IM-\d{4}-\d{4}$/.test(d.Consecutivo) ? d.Consecutivo : siguienteConsecutivo('IM', hoyCR());
  const ahora = ahoraCR(), nuevas = [], movs = [], omitidas = [], errores = [];
  entrada.forEach(f => {
    const placa = txt(f.Placa, 20).replace(/\s+/g, '');
    if (!/^\d{4}-\d{1,6}$/.test(placa)) { errores.push((placa || '(sin placa)') + ': placa inválida'); return; }
    if (!txt(f.Descripcion)) { errores.push(placa + ': sin descripción'); return; }
    if (existe[placa]) { omitidas.push(placa); return; }
    existe[placa] = true;
    const o = {
      Placa: placa, Descripcion: txt(f.Descripcion), Marca: txt(f.Marca), Modelo: txt(f.Modelo), Serie: txt(f.Serie),
      Estado: ESTADOS.indexOf(f.Estado) >= 0 ? f.Estado : '', Ubicacion: txt(f.Ubicacion) || txt(op.Ubicacion),
      UbicacionReal: txt(f.UbicacionReal), Especialidad: txt(f.Especialidad) || txt(op.Especialidad),
      ModoAdquisicion: txt(f.ModoAdquisicion) || txt(op.ModoAdquisicion), TFA: txt(f.TFA, 30), Observaciones: txt(f.Observaciones, 500),
      Situacion: tipo, FechaIngreso: fechaIng, Responsable: tipo === 'Activo' ? resp : '',
      FechaBaja: tipo === 'Baja' ? op.FechaBaja : '', MotivoBaja: tipo === 'Baja' ? op.Causal : '',
      ActaBaja: tipo === 'Baja' ? (txt(f.Acta, 30) || txt(op.Acta, 30)) : '', ConsecutivoBaja: '',
      Actualizado: ahora, ActualizadoPor: u.Correo
    };
    nuevas.push(enc.map(c => String(o[c] == null ? '' : o[c])));
    movs.push({ Consecutivo: cons, Fecha: tipo === 'Baja' ? op.FechaBaja : fechaIng, Tipo: tipo === 'Baja' ? 'Baja histórica (importación)' : 'Alta por importación',
      Placa: placa, Descripcion: o.Descripcion, Estado: o.Estado, AUbicacion: o.Ubicacion, AResponsable: o.Responsable,
      Causal: tipo === 'Baja' ? op.Causal : o.ModoAdquisicion, Detalle: 'Archivo: ' + txt(d.Archivo, 120) + (o.ActaBaja ? ' · acta ' + o.ActaBaja : '') });
  });
  if (nuevas.length) {
    capacidad(h, nuevas.length);
    h.getRange(h.getLastRow() + 1, 1, nuevas.length, enc.length).setValues(nuevas);
    const hm = hoja('Movimientos'), em = encabezado(hm);
    capacidad(hm, movs.length);
    hm.getRange(hm.getLastRow() + 1, 1, movs.length, em.length)
      .setValues(movs.map(m => em.map(c => c === 'RegistradoPor' ? u.Correo : c === 'Registrado' ? ahora : String(m[c] == null ? '' : m[c]))));
  }
  bitacora(u.Correo, 'inv_importacion', cons + ' · ' + nuevas.length + ' nuevos · ' + omitidas.length + ' existentes · ' + errores.length + ' con error');
  return { ok: true, consecutivo: cons, importados: nuevas.length, omitidas: omitidas, errores: errores };
}

/** Ejecutar UNA vez desde el editor: carga el inventario actual (activos y bajas). No duplica placas. */
function importarInventarioInicial() {
  const origen = SpreadsheetApp.openById(IMPORT_ID).getSheets()[0].getDataRange().getDisplayValues();
  const enc = origen.shift();
  const yaEstan = {};
  filas('Inventario').forEach(r => yaEstan[String(r.Placa).trim()] = true);
  const ahora = ahoraCR();
  const nuevas = [];
  origen.forEach(f => {
    const o = {}; enc.forEach((c, j) => o[c] = String(f[j] || '').trim());
    if (!o.Placa || yaEstan[o.Placa]) return;
    yaEstan[o.Placa] = true;
    o.Actualizado = ahora; o.ActualizadoPor = 'importación';
    nuevas.push(o);
  });
  const h = hoja('Inventario');
  const enc2 = encabezado(h);
  const matriz = nuevas.map(o => enc2.map(c => o[c] || ''));
  if (matriz.length) { capacidad(h, matriz.length); h.getRange(h.getLastRow() + 1, 1, matriz.length, enc2.length).setValues(matriz); }
  bitacora('', 'inv_importacion', nuevas.length + ' bienes');
  Logger.log('Importados: ' + nuevas.length + ' bienes.');
}

/* ============ UTILIDADES ============ */
function encabezado(h) { return h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(String); }

// Asegura filas libres (en formato texto) al final de la hoja de inventario
function capacidad(h, n) {
  const faltan = h.getLastRow() + n - h.getMaxRows();
  if (faltan > 0) {
    h.insertRowsAfter(h.getMaxRows(), faltan + 200);
    h.getRange(1, 1, h.getMaxRows(), h.getLastColumn()).setNumberFormat('@');
  }
}

let _ss = null;
const _migradas = {};
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
  } else if (h && HOJAS[n] && !_migradas[n]) {
    // Agrega al final las columnas nuevas que falten (actualización de versiones anteriores)
    const enc = h.getRange(1, 1, 1, Math.max(h.getLastColumn(), 1)).getValues()[0].map(String);
    const faltan = HOJAS[n].filter(c => enc.indexOf(c) < 0);
    if (faltan.length) {
      if (h.getMaxColumns() < enc.length + faltan.length) h.insertColumnsAfter(h.getMaxColumns(), faltan.length);
      h.getRange(1, enc.length + 1, 1, faltan.length).setValues([faltan]).setFontWeight('bold').setBackground('#0B2A55').setFontColor('#FFFFFF');
      if (n === 'Inventario') h.getRange(1, enc.length + 1, h.getMaxRows(), faltan.length).setNumberFormat('@');
    }
    _migradas[n] = true;
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

/**
 * Ejecutar desde el editor SOLO si se quiere empezar de cero.
 * Guarda primero una copia de respaldo de toda la hoja en su Drive y luego
 * borra los bienes y su historial (Inventario y Movimientos). No toca Usuarios ni Datos institucionales.
 */
function borrarInventario() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const copia = ss.copy('Respaldo inventario ' + ahoraCR());
  ['Inventario', 'Movimientos'].forEach(n => {
    const h = ss.getSheetByName(n);
    if (h && h.getLastRow() > 1) h.deleteRows(2, h.getLastRow() - 1);
  });
  bitacora('', 'inv_borrado_total', 'Respaldo: ' + copia.getName());
  Logger.log('Inventario borrado. Respaldo guardado en Drive: "' + copia.getName() + '"');
}

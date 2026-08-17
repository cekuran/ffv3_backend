/** @OnlyCurrentDoc */
// Finanzas Familia — backend Google Apps Script (Sheets + WebApp).
// Multi-usuario: cada fila conserva owner como metadato; el CRUD usa el id global.

const SCHEMA = {
  Cuentas:       ['owner','id','parent_id','nombre','tipo','moneda','icono','saldo_inicial','orden','oculta','establecimiento_id','fecha_creacion'],
  Categorias:    ['owner','id','nombre','color','icono','tipo','orden'],
  Establecimientos:['owner','id','nombre','web'],
  Transacciones: ['owner','id','fecha','tipo','importe','moneda','cuenta_id','subcuenta_id','cuenta_destino_id','subcuenta_destino_id','importe_destino','ratio_conversion','reparto_destino','categoria_id','descripcion','estado','recurrente_id','fecha_pago','conciliada_con','notas','fecha_creacion','ultima_edicion_por','fecha_ultima_edicion','establecimiento_id'],
  Recurrentes:   ['owner','id','plantilla','ultima_generacion','activa','mes_inicio','anio_inicio'],
  Presupuestos:  ['owner','id','anio','mes','categoria_id','importe_esperado'],
  Conciliaciones:['owner','id','fecha','cuenta_id','saldo_sistema','saldo_banco','diferencia','notas'],
  TiposCambio:   ['owner','id','fecha','base','destino','ratio']
};

const AUTH_SCHEMA = {
  Usuarios: ['username','password_hash','salt','rol','activo','fecha_creacion'],
  Spreadsheets: ['spreadsheet_id','nombre','descripcion','fecha_alta'],
  HojasUsuarios: ['username','spreadsheet_id','por_defecto','fecha_alta'],
  Tokens: ['token','username','fecha_creacion'],
  Config: ['clave','valor']
};
const ROLES = { ADMIN: 'admin', BASICO: 'basico' };

const HOJAS = Object.keys(SCHEMA);

const SEMILLA = {
  Cuentas: [
    {
      nombre: 'BBVA Nómina', tipo: 'activo', moneda: 'EUR', icono: 'account_balance', saldo_inicial: 0, orden: 1,
      subcuentas: [
        { nombre: 'Gastos diarios', saldo_inicial: 0, orden: 1 },
        { nombre: 'Recibos', saldo_inicial: 0, orden: 2 }
      ]
    },
    { nombre: 'Caja de Ahorro', tipo: 'activo', moneda: 'EUR', icono: 'savings', saldo_inicial: 0, orden: 2 },
    { nombre: 'Tarjeta Visa', tipo: 'pasivo', moneda: 'EUR', icono: 'credit_card', saldo_inicial: 0, orden: 3 }
  ],
  Categorias: [
    { nombre: 'Vivienda',       color: '#00613e', icono: 'home',            tipo: 'gasto',  orden: 1 },
    { nombre: 'Alimentación',   color: '#0c68db', icono: 'shopping_cart',   tipo: 'gasto',  orden: 2 },
    { nombre: 'Transporte',     color: '#ba1a1a', icono: 'directions_car',  tipo: 'gasto',  orden: 3 },
    { nombre: 'Ocio',           color: '#b13c68', icono: 'movie',           tipo: 'gasto',  orden: 4 },
    { nombre: 'Salud',          color: '#d27b1b', icono: 'fitness_center',  tipo: 'gasto',  orden: 5 },
    { nombre: 'Suscripciones',  color: '#0050af', icono: 'subscriptions',   tipo: 'gasto',  orden: 6 },
    { nombre: 'Nómina',         color: '#006c46', icono: 'payments',        tipo: 'ingreso',orden: 7 },
    { nombre: 'Ingresos extra', color: '#107c52', icono: 'work',            tipo: 'ingreso',orden: 8 }
  ]
};

// ───────── Helpers de sesión y ss ─────────
const MASTER_PROP_KEY = 'MASTER_SPREADSHEET_ID';
const DEFAULT_ADMIN_PASSWORD = 'admin1234';
const APP_VERSION = 'ms-http-1';
const CONFIG_DATA_SHEET_ID = 'SHEET_ID_PRODUCTION';

function getMasterSpreadsheetId_() {
  return PropertiesService.getScriptProperties().getProperty(MASTER_PROP_KEY) || '';
}

// ───────── Script Cache (CacheService) ─────────
// Caché compartida entre ejecuciones del Web App. Reduce lecturas al spreadsheet
// maestro (Tokens, Usuarios, HojasUsuarios, Spreadsheets, Config) en el path
// caliente de cada request autenticada.
// Límites GAS: ~100 KB/entrada, TTL máx. 21600 s. Fallos de caché se ignoran.
const CACHE_TTL_TOKEN_SEC = 30 * 60;       // 30 min — validación de sesión
const CACHE_TTL_AUTH_SHEET_SEC = 5 * 60;   // 5 min — hojas del spreadsheet maestro
const CACHE_TTL_SALDOS_SEC = 10 * 60;      // 10 min — saldos/evolución de cuentas
const CACHE_MAX_JSON_CHARS = 90000;        // margen bajo el límite ~100 KB

function scriptCache_() {
  return CacheService.getScriptCache();
}

// Prefijos del valor almacenado en CacheService:
//   z:  JSON gzip + base64 (preferido si reduce tamaño)
//   r:  JSON crudo
//   (sin prefijo) JSON legacy de versiones anteriores
const CACHE_PREFIX_GZIP = 'z:';
const CACHE_PREFIX_RAW = 'r:';

function cacheSerialize_(value) {
  const json = JSON.stringify(value);
  if (!json) return '';
  // Gzip solo compensa a partir de un tamaño mínimo (overhead base64 ~33%).
  if (json.length >= 400) {
    try {
      const gzBlob = Utilities.gzip(Utilities.newBlob(json, 'application/json'));
      const b64 = Utilities.base64Encode(gzBlob.getBytes());
      if (b64 && b64.length + 2 < json.length + 2) {
        return CACHE_PREFIX_GZIP + b64;
      }
    } catch (e) {
      /* fallback a raw */
    }
  }
  return CACHE_PREFIX_RAW + json;
}

function cacheDeserialize_(raw) {
  if (!raw) return null;
  if (raw.indexOf(CACHE_PREFIX_GZIP) === 0) {
    const bytes = Utilities.base64Decode(raw.substring(CACHE_PREFIX_GZIP.length));
    const plain = Utilities.ungzip(Utilities.newBlob(bytes)).getDataAsString();
    return JSON.parse(plain);
  }
  if (raw.indexOf(CACHE_PREFIX_RAW) === 0) {
    return JSON.parse(raw.substring(CACHE_PREFIX_RAW.length));
  }
  // Compatibilidad con entradas previas (JSON sin prefijo).
  return JSON.parse(raw);
}

function cacheGetJson_(key) {
  try {
    const raw = scriptCache_().get(key);
    if (!raw) return null;
    return cacheDeserialize_(raw);
  } catch (e) {
    return null;
  }
}

function cachePutJson_(key, value, ttlSec) {
  try {
    const s = cacheSerialize_(value);
    if (!s || s.length > CACHE_MAX_JSON_CHARS) return false;
    scriptCache_().put(key, s, Math.min(ttlSec || CACHE_TTL_AUTH_SHEET_SEC, 21600));
    return true;
  } catch (e) {
    return false;
  }
}

function cacheRemove_(key) {
  try { scriptCache_().remove(key); } catch (e) { /* ignore */ }
}

function cacheRemoveAll_(keys) {
  if (!keys || !keys.length) return;
  try { scriptCache_().removeAll(keys); } catch (e) { /* ignore */ }
}

function tokenCacheKey_(token) {
  return 'tok:' + String(token || '').trim();
}

// ponytail: prefijo auth:* es global a propósito — las hojas Usuarios/Tokens/
// HojasUsuarios/Spreadsheets/Config viven en el spreadsheet maestro y son
// compartidas por todos los usuarios. No namespaciar por usuario.
function authSheetCacheKey_(nombre) {
  return 'auth:' + String(nombre || '');
}

// Caché en memoria solo para la request actual (evita N hits a CacheService
// o al sheet cuando bootstrap / _authadmin leen varias veces la misma hoja).
const _authReadCache = {};

function invalidateAuthSheetCache_(nombre) {
  if (nombre) {
    delete _authReadCache[nombre];
    cacheRemove_(authSheetCacheKey_(nombre));
  }
}

// Token cacheado durante una única petición HTTP; la persistencia vive en Tokens.
let _currentToken = '';
// Hoja activa resuelta por _authadmin a partir de HojasUsuarios; las funciones
// de datos la consultan vía ssActiva_() en vez del sheet fijo por entorno.
let _currentSheetId = '';
// Override para los self-tests: cuando está activo, currentUser_/currentRol_
// devuelven un usuario fijo sin chequear _currentToken.
// Solo se activa dentro de __selfTest*/__selfTestHojas* (try/finally) y se
// descarta al volver la request.
let _selfTestActive = false;
const SELF_TEST_USER = '__selftest__';

function currentUser_() {
  if (_selfTestActive) return SELF_TEST_USER;
  if (!_currentToken) return '';
  return validarTokenSesion_(_currentToken);
}

function requireUsuario_() {
  const u = currentUser_();
  if (!u) throw new Error('No autenticado');
  return u;
}

function currentRol_() {
  if (_selfTestActive) return ROLES.ADMIN;
  const username = currentUser_();
  if (!username) return '';
  const u = buscarUsuario_(username);
  return u ? String(u.rol || ROLES.BASICO) : '';
}

function requireAdmin_() {
  if (currentRol_() !== ROLES.ADMIN) throw new Error('Solo admin puede realizar esta acción');
}

// Despacho autenticado: valida el token guardado en Sheets, lo cachea
// en _currentToken y ejecuta la función pedida. Sustituye al patrón
// setAuthToken + fn en dos invocaciones (la caché no sobrevive entre
// Endpoints que siguen disponibles aunque el usuario no tenga hojas
// vinculadas (no son admin). El resto devuelven error para no filtrar
// datos del spreadsheet de producción.
const ALWAYS_ALLOWED_FOR_NO_HOJAS = new Set([
  'bootstrap', 'bootstrapBase', 'authStatus', 'loginUsuario', 'logoutUsuario',
  'ping', 'configurarSpreadsheetMaestro', 'cambiarMiContrasena',
  'listarMisHojas', 'cambiarHojaActiva'
]);

function usuarioTieneHojas_(username) {
  if (!username) return false;
  return leerHojasUsuarios_().some(l =>
    String(l.username || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()
  );
}

function _authadmin(token, fnName, ...args) {
  const username = validarTokenSesion_(token);
  if (!username) throw new Error('No autenticado');
  _currentToken = String(token || '').trim();
  try { _currentSheetId = resolverHojaActivaId_(username); }
  catch (e) { _currentSheetId = ''; }
  // Guard: si el usuario (no admin) no tiene hojas, sólo se permite un
  // conjunto mínimo de acciones. Cualquier endpoint de datos falla rápido
  // para impedir lecturas filtradas del spreadsheet de producción.
  const usuario = buscarUsuario_(username);
  const rol = usuario ? String(usuario.rol || ROLES.BASICO) : ROLES.BASICO;
  if (rol !== ROLES.ADMIN && !usuarioTieneHojas_(username) && !ALWAYS_ALLOWED_FOR_NO_HOJAS.has(fnName)) {
    throw new Error('No tienes hojas de cálculo asignadas. Pide al administrador que vincule una hoja.');
  }
  const fn = globalThis[fnName];
  if (typeof fn !== 'function') throw new Error('Función no encontrada: ' + fnName);
  return fn.apply(null, args);
}

function bytesHex_(bytes) {
  return bytes.map(function (b) {
    const h = (b < 0 ? b + 256 : b).toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
}

function sha256Hex_(text) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytesHex_(digest);
}

function passwordHash_(password, salt) {
  let h = salt + '|' + password;
  for (let i = 0; i < 12000; i++) h = sha256Hex_(h);
  return h;
}

function crearTokenSesion_(username) {
  const token = Utilities.getUuid();
  const usernameNorm = String(username || '').trim();
  const filas = leerAuthHojaGenerica_('Tokens');
  filas.push({ token: token, username: usernameNorm, fecha_creacion: isoAhora_() });
  escribirAuthHojaGenerica_('Tokens', filas);
  // Poblar caché de token de inmediato para la siguiente request.
  if (usernameNorm) cachePutJson_(tokenCacheKey_(token), { u: usernameNorm }, CACHE_TTL_TOKEN_SEC);
  return token;
}

function validarTokenSesion_(token) {
  const t = String(token || '').trim();
  if (!t) return '';
  // 1) Caché Script: evita abrir el master spreadsheet en casi todas las requests.
  const hit = cacheGetJson_(tokenCacheKey_(t));
  if (hit && hit.u) return String(hit.u);
  // 2) Fallback a hoja Tokens (con caché de hoja + memoria de request).
  const fila = leerAuthHojaGenerica_('Tokens').find(r => String(r.token || '') === t);
  const username = fila ? String(fila.username || '').trim() : '';
  if (username) cachePutJson_(tokenCacheKey_(t), { u: username }, CACHE_TTL_TOKEN_SEC);
  return username;
}

function invalidarTokenSesion_(token) {
  const t = String(token || '').trim();
  if (!t) return;
  // Invalidar antes de escribir para que un request concurrente no rehidrate
  // un token ya revocado desde la hoja vieja en caché.
  cacheRemove_(tokenCacheKey_(t));
  escribirAuthHojaGenerica_('Tokens', leerAuthHojaGenerica_('Tokens').filter(r => String(r.token || '') !== t));
}

function asegurarUsuarios_() {
  asegurarAuthHojaUsuarios_();
  const rows = leerUsuariosAuth_().filter(r => r.username);
  if (rows.length) return;

  const legacy = leerUsuariosLegacyData_();
  if (legacy.length) {
    escribirUsuariosAuth_(legacy);
    return;
  }

  const defaultUser = 'admin';
  const defaultPass = DEFAULT_ADMIN_PASSWORD;
  const salt = Utilities.getUuid().replace(/-/g, '');
  const nuevo = {
    username: defaultUser,
    password_hash: passwordHash_(defaultPass, salt),
    salt: salt,
    rol: ROLES.ADMIN,
    activo: true,
    fecha_creacion: isoAhora_()
  };
  escribirUsuariosAuth_([nuevo]);
}

function buscarUsuario_(username) {
  const target = String(username || '').trim().toLowerCase();
  if (!target) return null;
  return leerUsuariosAuth_().find(u => String(u.username || '').trim().toLowerCase() === target) || null;
}

function authStatus(token) {
  const tokenUser = validarTokenSesion_(token);
  if (tokenUser) _currentToken = String(token || '').trim();
  else _currentToken = '';
  const rol = tokenUser ? (buscarUsuario_(tokenUser) || {}).rol || ROLES.BASICO : '';
  return { authenticated: !!tokenUser, user: tokenUser || '', rol: rol || '' };
}

function loginUsuario(username, password) {
  asegurarUsuarios_();
  const user = String(username || '').trim();
  const pass = String(password || '');
  if (!user || !pass) throw new Error('Debes indicar usuario y contraseña');
  const found = buscarUsuario_(user);
  if (!found || String(found.activo) === 'false') throw new Error('Credenciales inválidas');
  const hash = passwordHash_(pass, String(found.salt || ''));
  if (hash !== String(found.password_hash || '')) throw new Error('Credenciales inválidas');
  const token = crearTokenSesion_(found.username);
  return { ok: true, user: found.username, rol: String(found.rol || ROLES.BASICO), token: token };
}

function setAuthToken(token) {
  const username = validarTokenSesion_(token);
  if (!username) throw new Error('No autenticado');
  _currentToken = String(token || '').trim();
  return { ok: true, user: username };
}

function logoutUsuario(token) {
  invalidarTokenSesion_(token);
  _currentToken = '';
  return { ok: true };
}

function crearUsuarioAdmin(username, password, rol) {
  requireAdmin_();
  asegurarUsuarios_();
  const user = String(username || '').trim();
  const pass = String(password || '');
  const rolFinal = String(rol || ROLES.BASICO).trim().toLowerCase();
  if (!Object.values(ROLES).includes(rolFinal)) throw new Error('Rol inválido');
  if (!/^[a-zA-Z0-9_.-]{3,40}$/.test(user)) throw new Error('Usuario inválido (3-40, letras, números, _.-)');
  if (pass.length < 8) throw new Error('La contraseña debe tener mínimo 8 caracteres');
  if (buscarUsuario_(user)) throw new Error('Ese usuario ya existe');
  const salt = Utilities.getUuid().replace(/-/g, '');
  const rows = leerUsuariosAuth_();
  rows.push({
    username: user,
    password_hash: passwordHash_(pass, salt),
    salt: salt,
    rol: rolFinal,
    activo: true,
    fecha_creacion: isoAhora_()
  });
  escribirUsuariosAuth_(rows);
  return { ok: true, user: user, rol: rolFinal };
}

function listarUsuariosAdmin() {
  requireAdmin_();
  asegurarUsuarios_();
  return leerUsuariosAuth_().map(u => ({
    username: u.username,
    rol: String(u.rol || ROLES.BASICO),
    activo: String(u.activo) !== 'false',
    fecha_creacion: u.fecha_creacion
  }));
}

function resetearContrasenaAdmin(username, passwordNueva) {
  requireAdmin_();
  asegurarUsuarios_();
  const user = String(username || '').trim();
  const nueva = String(passwordNueva || '');
  if (!user) throw new Error('Debes indicar el usuario');
  if (nueva.length < 8) throw new Error('La nueva contraseña debe tener mínimo 8 caracteres');
  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');
  const saltNuevo = Utilities.getUuid().replace(/-/g, '');
  rows[idx].salt = saltNuevo;
  rows[idx].password_hash = passwordHash_(nueva, saltNuevo);
  escribirUsuariosAuth_(rows);
  return { ok: true, user: user };
}

function cambiarRolUsuarioAdmin(username, rol) {
  requireAdmin_();
  asegurarUsuarios_();
  const actor = currentUser_();
  const user = String(username || '').trim();
  const rolFinal = String(rol || ROLES.BASICO).trim().toLowerCase();
  if (!Object.values(ROLES).includes(rolFinal)) throw new Error('Rol inválido');
  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');
  const esMismoActor = String(rows[idx].username || '').trim().toLowerCase() === String(actor || '').trim().toLowerCase();
  if (rolFinal !== ROLES.ADMIN) {
    const adminsRestantes = rows.filter(u => String(u.rol) === ROLES.ADMIN && String(u.username || '').trim().toLowerCase() !== user.toLowerCase()).length;
    if (esMismoActor && adminsRestantes === 0) throw new Error('No puedes quitarte el último admin');
  }
  rows[idx].rol = rolFinal;
  escribirUsuariosAuth_(rows);
  return { ok: true, user: user, rol: rolFinal };
}

function eliminarUsuarioAdmin(username) {
  requireAdmin_();
  asegurarUsuarios_();
  const actor = currentUser_();
  const user = String(username || '').trim();
  if (!user) throw new Error('Debes indicar el usuario');
  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');
  if (String(rows[idx].username || '').trim().toLowerCase() === String(actor || '').trim().toLowerCase()) {
    throw new Error('No puedes eliminarte a ti mismo');
  }
  if (String(rows[idx].rol) === ROLES.ADMIN) {
    const otrosAdmins = rows.filter(u => String(u.rol) === ROLES.ADMIN && String(u.username || '').trim().toLowerCase() !== user.toLowerCase()).length;
    if (otrosAdmins === 0) throw new Error('No puedes eliminar al último admin');
  }
  rows.splice(idx, 1);
  escribirUsuariosAuth_(rows);
  // Limpiar vinculaciones del usuario eliminado para no dejar huérfanas.
  const links = leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() !== user.toLowerCase());
  escribirHojasUsuarios_(links);
  return { ok: true, user: user };
}

function cambiarMiContrasena(passwordActual, passwordNueva) {
  const actor = requireUsuario_();
  asegurarUsuarios_();
  const actual = String(passwordActual || '');
  const nueva = String(passwordNueva || '');
  if (!actual || !nueva) throw new Error('Debes indicar la contraseña actual y la nueva');
  if (nueva.length < 8) throw new Error('La nueva contraseña debe tener mínimo 8 caracteres');

  const rows = leerUsuariosAuth_();
  const idx = rows.findIndex(u => String(u.username || '').trim().toLowerCase() === String(actor).toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');

  const user = rows[idx];
  if (String(user.activo) === 'false') throw new Error('Usuario inactivo');
  const actualHash = passwordHash_(actual, String(user.salt || ''));
  if (actualHash !== String(user.password_hash || '')) throw new Error('La contraseña actual no es correcta');
  if (passwordHash_(nueva, String(user.salt || '')) === String(user.password_hash || '')) {
    throw new Error('La nueva contraseña debe ser distinta de la actual');
  }

  const saltNuevo = Utilities.getUuid().replace(/-/g, '');
  rows[idx].salt = saltNuevo;
  rows[idx].password_hash = passwordHash_(nueva, saltNuevo);
  escribirUsuariosAuth_(rows);
  return { ok: true, user: actor };
}

function obtenerConfig_(clave) {
  const fila = leerAuthHojaGenerica_('Config').find(r => String(r.clave || '') === clave);
  return fila ? String(fila.valor || '').trim() : '';
}

function guardarConfig_(clave, valor) {
  const filas = leerAuthHojaGenerica_('Config');
  const existente = filas.find(r => String(r.clave || '') === clave);
  if (existente) existente.valor = String(valor || '').trim();
  else filas.push({ clave: clave, valor: String(valor || '').trim() });
  escribirAuthHojaGenerica_('Config', filas);
  return String(valor || '').trim();
}

function obtenerSheetIdConfigurado_() {
  return obtenerConfig_(CONFIG_DATA_SHEET_ID);
}

function guardarSheetIdParaEntorno_(sheetId) {
  return guardarConfig_(CONFIG_DATA_SHEET_ID, sheetId);
}

function obtenerAuthSheetIdConfigurado_() {
  return getMasterSpreadsheetId_();
}

function guardarAuthSheetIdParaEntorno_(sheetId) {
  const id = String(sheetId || '').trim();
  PropertiesService.getScriptProperties().setProperty(MASTER_PROP_KEY, id);
  return id;
}

function ss_() {
  const configuredId = obtenerSheetIdConfigurado_();
  if (configuredId) {
    try { return SpreadsheetApp.openById(configuredId); }
    catch (e) { /* id inválido, recrear */ }
  }
  // Primera vez (o id corrupto): crear spreadsheet por defecto.
  const ss = SpreadsheetApp.create('Finanzas Familia [production]');
  guardarSheetIdParaEntorno_(ss.getId());
  // Mover la hoja "Hoja 1" por defecto al final, queda fuera de la vista
  const porDefecto = ss.getSheets()[0];
  if (porDefecto && ss.getSheets().length === 1) porDefecto.setName('_log');
  return ss;
}

// Hoja de datos del usuario activo. Si _currentSheetId está resuelto (por
// _authadmin), abre esa; en caso contrario cae al sheet por entorno.
function ssActiva_() {
  if (_currentSheetId) {
    try { return SpreadsheetApp.openById(_currentSheetId); }
    catch (e) { /* id inválido o revocado: fallback */ }
  }
  return ss_();
}

function authSs_() {
  const id = getMasterSpreadsheetId_();
  if (!id) throw new Error('Configura MASTER_SPREADSHEET_ID en Script Properties');
  return SpreadsheetApp.openById(id);
}

function asegurarAuthHojaUsuarios_() {
  const ss = authSs_();
  const nombre = 'Usuarios';
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, AUTH_SCHEMA.Usuarios.length).setValues([AUTH_SCHEMA.Usuarios]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, AUTH_SCHEMA.Usuarios.length).setValues([AUTH_SCHEMA.Usuarios]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else {
    const cab = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
    if (!cab.includes('rol')) h.getRange(1, h.getLastColumn() + 1).setValue('rol');
    h.setFrozenRows(1);
  }
  return h;
}

function leerUsuariosAuth_() {
  const cacheName = 'Usuarios';
  if (_authReadCache[cacheName]) return cloneRows_(_authReadCache[cacheName]);
  const cached = cacheGetJson_(authSheetCacheKey_(cacheName));
  if (cached) {
    const rows = cached.map(normalizarUsuarioAuth_);
    _authReadCache[cacheName] = rows;
    return cloneRows_(rows);
  }
  const h = asegurarAuthHojaUsuarios_();
  const valores = h.getDataRange().getValues();
  if (valores.length < 2) {
    _authReadCache[cacheName] = [];
    cachePutJson_(authSheetCacheKey_(cacheName), [], CACHE_TTL_AUTH_SHEET_SEC);
    return [];
  }
  const cab = valores[0];
  const rows = valores.slice(1).map(fila => {
    const o = {};
    cab.forEach((k, i) => (o[k] = fila[i]));
    return normalizarUsuarioAuth_(o);
  });
  _authReadCache[cacheName] = rows;
  cachePutJson_(authSheetCacheKey_(cacheName), rows, CACHE_TTL_AUTH_SHEET_SEC);
  return cloneRows_(rows);
}

function normalizarUsuarioAuth_(u) {
  if (!u) return u;
  const rol = String(u.rol || '').trim();
  // Migración: filas previas a la columna 'rol' (o vacías) — el primer admin
  // conserva su rol; el resto cae a 'basico'.
  u.rol = rol || (String(u.username || '').trim().toLowerCase() === 'admin' ? ROLES.ADMIN : ROLES.BASICO);
  return u;
}

function escribirUsuariosAuth_(filas) {
  const h = asegurarAuthHojaUsuarios_();
  const cab = AUTH_SCHEMA.Usuarios;
  h.clearContents();
  const normalizadas = (filas || []).map(f => normalizarUsuarioAuth_(Object.assign({}, f)));
  const matriz = [cab].concat(normalizadas.map(u => cab.map(k => u[k] != null ? u[k] : '')));
  if (matriz.length) h.getRange(1, 1, matriz.length, cab.length).setValues(matriz);
  h.setFrozenRows(1);
  _authReadCache['Usuarios'] = cloneRows_(normalizadas);
  cachePutJson_(authSheetCacheKey_('Usuarios'), normalizadas, CACHE_TTL_AUTH_SHEET_SEC);
}

function leerUsuariosLegacyData_() {
  try {
    const h = ss_().getSheetByName('Usuarios');
    if (!h || h.getLastRow() < 2) return [];
    const vals = h.getDataRange().getValues();
    const cab = vals[0].map(c => String(c || '').trim());
    const idx = {};
    AUTH_SCHEMA.Usuarios.forEach(k => { idx[k] = cab.indexOf(k); });
    if (idx.username < 0 || idx.password_hash < 0 || idx.salt < 0) return [];
    return vals.slice(1).map(row => ({
      username: row[idx.username],
      password_hash: row[idx.password_hash],
      salt: row[idx.salt],
      rol: idx.rol >= 0 ? row[idx.rol] : '',
      activo: idx.activo >= 0 ? row[idx.activo] : true,
      fecha_creacion: idx.fecha_creacion >= 0 ? row[idx.fecha_creacion] : isoAhora_()
    })).filter(r => r.username && r.password_hash && r.salt);
  } catch (e) {
    return [];
  }
}

// ───────── Auth sheet genérico (Spreadsheets + HojasUsuarios) ─────────
// ponytail: estas tablas viven en el auth sheet, no en la hoja de datos
// por usuario. Se leen/escriben como filas planas; el shape lo define
// AUTH_SCHEMA[nombre].

function asegurarAuthHojaGenerica_(nombre) {
  const ss = authSs_();
  const cab = AUTH_SCHEMA[nombre];
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

function leerAuthHojaGenerica_(nombre) {
  if (_authReadCache[nombre]) return cloneRows_(_authReadCache[nombre]);
  const cached = cacheGetJson_(authSheetCacheKey_(nombre));
  if (cached) {
    _authReadCache[nombre] = cached;
    return cloneRows_(cached);
  }
  const h = asegurarAuthHojaGenerica_(nombre);
  const valores = h.getDataRange().getValues();
  if (valores.length < 2) {
    _authReadCache[nombre] = [];
    cachePutJson_(authSheetCacheKey_(nombre), [], CACHE_TTL_AUTH_SHEET_SEC);
    return [];
  }
  const cab = valores[0];
  const rows = valores.slice(1).map(fila => {
    const o = {};
    cab.forEach((k, i) => (o[k] = fila[i]));
    return o;
  });
  _authReadCache[nombre] = rows;
  cachePutJson_(authSheetCacheKey_(nombre), rows, CACHE_TTL_AUTH_SHEET_SEC);
  return cloneRows_(rows);
}

function escribirAuthHojaGenerica_(nombre, filas) {
  const h = asegurarAuthHojaGenerica_(nombre);
  const cab = AUTH_SCHEMA[nombre];
  h.clearContents();
  const filasSafe = filas || [];
  const matriz = [cab].concat(filasSafe.map(f => cab.map(k => f[k] != null ? f[k] : '')));
  if (matriz.length) h.getRange(1, 1, matriz.length, cab.length).setValues(matriz);
  h.setFrozenRows(1);
  // Mantener caché coherente tras la escritura (request + Script Cache).
  _authReadCache[nombre] = cloneRows_(filasSafe);
  cachePutJson_(authSheetCacheKey_(nombre), filasSafe, CACHE_TTL_AUTH_SHEET_SEC);
}

function leerSpreadsheets_() { return leerAuthHojaGenerica_('Spreadsheets'); }
function escribirSpreadsheets_(filas) { escribirAuthHojaGenerica_('Spreadsheets', filas); }
function leerHojasUsuarios_() { return leerAuthHojaGenerica_('HojasUsuarios'); }
function escribirHojasUsuarios_(filas) { escribirAuthHojaGenerica_('HojasUsuarios', filas); }

// Resuelve la hoja activa del usuario actual con la siguiente prioridad:
// 1) Hoja por defecto explícita en HojasUsuarios
// 2) Primera hoja vinculada
function resolverHojaActivaId_(username) {
  username = String(username || '').trim();
  if (!username) return '';
  const links = leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() === username.toLowerCase());
  const defecto = links.find(l => String(l.por_defecto) === 'true' || String(l.por_defecto) === true);
  if (defecto) return String(defecto.spreadsheet_id);
  if (links.length) return String(links[0].spreadsheet_id);
  return '';
}

function vincularHojaUsuarioInternal_(username, spreadsheetId, porDefecto) {
  const filas = leerHojasUsuarios_();
  const idx = filas.findIndex(l =>
    String(l.username || '').trim().toLowerCase() === String(username).toLowerCase() &&
    String(l.spreadsheet_id || '') === String(spreadsheetId)
  );
  if (idx >= 0) return filas[idx];
  filas.push({
    username: username,
    spreadsheet_id: spreadsheetId,
    por_defecto: !!porDefecto,
    fecha_alta: isoAhora_()
  });
  escribirHojasUsuarios_(filas);
  return filas[filas.length - 1];
}

// ───────── CRUD admin: Spreadsheets y vinculaciones ─────────

function listarSpreadsheetsAdmin() {
  requireAdmin_();
  const sheets = leerSpreadsheets_();
  const links = leerHojasUsuarios_();
  return sheets.map(s => {
    const vinculados = links.filter(l => String(l.spreadsheet_id) === String(s.spreadsheet_id));
    return {
      spreadsheet_id: String(s.spreadsheet_id),
      nombre: String(s.nombre || ''),
      descripcion: String(s.descripcion || ''),
      fecha_alta: s.fecha_alta || '',
      usuarios: vinculados.map(l => String(l.username))
    };
  });
}

function altaSpreadsheetAdmin(spreadsheetId, nombre, descripcion) {
  requireAdmin_();
  const id = String(spreadsheetId || '').trim();
  const nom = String(nombre || '').trim();
  if (!id) throw new Error('Indica el spreadsheet ID');
  if (!nom) throw new Error('Indica un nombre');
  if (!/^[a-zA-Z0-9_-]{20,}$/.test(id)) throw new Error('El ID no parece un spreadsheet ID de Google');
  // Validación temprana: intenta abrir. Si no es accesible, falla aquí.
  try { SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('No se puede abrir el spreadsheet: ' + (e && e.message || e)); }
  const filas = leerSpreadsheets_();
  if (filas.some(s => String(s.spreadsheet_id) === id)) throw new Error('Ese spreadsheet ya está registrado');
  filas.push({
    spreadsheet_id: id,
    nombre: nom,
    descripcion: String(descripcion || '').trim(),
    fecha_alta: isoAhora_()
  });
  escribirSpreadsheets_(filas);
  return { ok: true, spreadsheet_id: id, nombre: nom };
}

function bajaSpreadsheetAdmin(spreadsheetId) {
  requireAdmin_();
  const id = String(spreadsheetId || '').trim();
  if (!id) throw new Error('Indica el spreadsheet ID');
  const sheets = leerSpreadsheets_();
  if (!sheets.some(s => String(s.spreadsheet_id) === id)) throw new Error('Spreadsheet no registrado');
  const links = leerHojasUsuarios_().filter(l => String(l.spreadsheet_id) !== id);
  escribirHojasUsuarios_(links);
  escribirSpreadsheets_(sheets.filter(s => String(s.spreadsheet_id) !== id));
  return { ok: true };
}

function resetearSpreadsheetAdmin(spreadsheetId, seed) {
  requireAdmin_();
  const id = String(spreadsheetId || '').trim();
  if (!id) throw new Error('Indica el spreadsheet ID');
  if (!leerSpreadsheets_().some(s => String(s.spreadsheet_id) === id)) {
    throw new Error('Spreadsheet no registrado');
  }
  const seedNormalizado = normalizarSeed_(seed);
  const usuarios = [...new Set(leerHojasUsuarios_()
    .filter(l => String(l.spreadsheet_id) === id)
    .map(l => String(l.username || '').trim())
    .filter(Boolean))];

  _currentSheetId = id;
  migrarEsquema();
  HOJAS.forEach(nombre => escribirHoja(nombre, []));
  usuarios.forEach(owner => {
    sembrar(owner, seedNormalizado);
  });
  normalizarCuentasSinSubcuentas_();

  return { ok: true, spreadsheet_id: id, usuarios: usuarios };
}

// ponytail: el seed personalizado se valida y sanea aquí. Si viene vacío o
// inválido, cae al SEMILLA global. Devuelve { cuentas, categorias } listo para
// pasar a sembrar().
function normalizarSeed_(seed) {
  if (!seed || typeof seed !== 'object') return null;
  const cuentas = Array.isArray(seed.cuentas) ? seed.cuentas : [];
  const categorias = Array.isArray(seed.categorias) ? seed.categorias : [];
  const cuentasOk = cuentas
    .map(c => ({
      nombre: String(c.nombre || '').trim(),
      tipo: c.tipo === 'pasivo' ? 'pasivo' : 'activo',
      moneda: String(c.moneda || 'EUR').toUpperCase(),
      icono: String(c.icono || 'account_balance_wallet'),
      saldo_inicial: Number(c.saldo_inicial || 0),
      orden: Number(c.orden || 0) || 99,
      subcuentas: Array.isArray(c.subcuentas) ? c.subcuentas
        .map(s => ({
          nombre: String(s.nombre || '').trim(),
          saldo_inicial: Number(s.saldo_inicial || 0),
          orden: Number(s.orden || 0) || 99
        }))
        .filter(s => s.nombre) : []
    }))
    .filter(c => c.nombre);
  const categoriasOk = categorias
    .map(cat => ({
      nombre: String(cat.nombre || '').trim(),
      color: String(cat.color || '#00613e'),
      icono: String(cat.icono || 'category'),
      tipo: cat.tipo === 'ingreso' ? 'ingreso' : 'gasto',
      orden: Number(cat.orden || 0) || 99
    }))
    .filter(cat => cat.nombre);
  if (!cuentasOk.length && !categoriasOk.length) return null;
  return { cuentas: cuentasOk, categorias: categoriasOk };
}

function renombrarSpreadsheetAdmin(spreadsheetId, nombre) {
  requireAdmin_();
  const id = String(spreadsheetId || '').trim();
  const nom = String(nombre || '').trim();
  if (!id) throw new Error('Indica el spreadsheet ID');
  if (!nom) throw new Error('Indica un nombre');
  const sheets = leerSpreadsheets_();
  const idx = sheets.findIndex(s => String(s.spreadsheet_id) === id);
  if (idx === -1) throw new Error('Spreadsheet no registrado');
  sheets[idx].nombre = nom;
  escribirSpreadsheets_(sheets);
  return { ok: true, spreadsheet_id: id, nombre: nom };
}

function listarVinculacionesAdmin() {
  requireAdmin_();
  const sheets = leerSpreadsheets_();
  const porId = {};
  sheets.forEach(s => { porId[String(s.spreadsheet_id)] = s; });
  const users = leerUsuariosAuth_().map(u => String(u.username || '').trim()).filter(Boolean);
  return users.map(username => {
    const links = leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() === username.toLowerCase());
    return {
      username: username,
      hojas: links.map(l => ({
        spreadsheet_id: String(l.spreadsheet_id),
        nombre: (porId[String(l.spreadsheet_id)] || {}).nombre || '(sin nombre)',
        por_defecto: String(l.por_defecto) === 'true' || l.por_defecto === true
      }))
    };
  });
}

function vincularHojaUsuarioAdmin(username, spreadsheetId, porDefecto) {
  requireAdmin_();
  const user = String(username || '').trim();
  const id = String(spreadsheetId || '').trim();
  if (!user) throw new Error('Indica el usuario');
  if (!id) throw new Error('Indica el spreadsheet ID');
  if (!leerUsuariosAuth_().some(u => String(u.username || '').trim().toLowerCase() === user.toLowerCase())) {
    throw new Error('Usuario no existe');
  }
  if (!leerSpreadsheets_().some(s => String(s.spreadsheet_id) === id)) {
    throw new Error('Spreadsheet no registrado; primero añádelo en el directorio');
  }
  const quiereDefecto = porDefecto === true || porDefecto === 'true';
  const links = leerHojasUsuarios_();
  const existe = links.find(l =>
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (existe) throw new Error('Ese usuario ya tiene vinculada esa hoja');
  if (quiereDefecto) {
    // Solo puede haber una hoja por defecto por usuario; desmarca las demás.
    links.forEach(l => {
      if (String(l.username || '').trim().toLowerCase() === user.toLowerCase()) l.por_defecto = false;
    });
  } else if (!links.some(l => String(l.username || '').trim().toLowerCase() === user.toLowerCase())) {
    // Primera hoja del usuario → se marca por defecto automáticamente.
    porDefecto = true;
  }
  links.push({
    username: user,
    spreadsheet_id: id,
    por_defecto: porDefecto === true || porDefecto === 'true',
    fecha_alta: isoAhora_()
  });
  escribirHojasUsuarios_(links);
  return { ok: true, username: user, spreadsheet_id: id, por_defecto: porDefecto };
}

function desvincularHojaUsuarioAdmin(username, spreadsheetId) {
  requireAdmin_();
  const user = String(username || '').trim();
  const id = String(spreadsheetId || '').trim();
  if (!user || !id) throw new Error('Faltan parámetros');
  const links = leerHojasUsuarios_();
  const restantes = links.filter(l => !(
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id
  ));
  if (restantes.length === links.length) throw new Error('Vinculación no encontrada');
  // Si desvinculamos la hoja por defecto, promover otra del mismo usuario.
  const eraDefecto = links.find(l =>
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id &&
    (String(l.por_defecto) === 'true' || l.por_defecto === true)
  );
  escribirHojasUsuarios_(restantes);
  if (eraDefecto) {
    const otra = restantes.find(l => String(l.username || '').trim().toLowerCase() === user.toLowerCase());
    if (otra) { otra.por_defecto = true; escribirHojasUsuarios_(restantes); }
  }
  return { ok: true };
}

function setHojaPorDefectoAdmin(username, spreadsheetId) {
  requireAdmin_();
  const user = String(username || '').trim();
  const id = String(spreadsheetId || '').trim();
  if (!user || !id) throw new Error('Faltan parámetros');
  const links = leerHojasUsuarios_();
  const target = links.find(l =>
    String(l.username || '').trim().toLowerCase() === user.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (!target) throw new Error('El usuario no tiene vinculada esa hoja');
  links.forEach(l => {
    if (String(l.username || '').trim().toLowerCase() === user.toLowerCase()) l.por_defecto = false;
  });
  target.por_defecto = true;
  escribirHojasUsuarios_(links);
  return { ok: true };
}

function listarMisHojas() {
  const username = requireUsuario_();
  return listarHojasDelUsuario_(username);
}

function cambiarHojaActiva(spreadsheetId) {
  const username = requireUsuario_();
  const id = String(spreadsheetId || '').trim();
  if (!id) throw new Error('Indica la hoja');
  const links = leerHojasUsuarios_();
  const existe = links.some(l =>
    String(l.username || '').trim().toLowerCase() === username.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (!existe) throw new Error('No tienes vinculada esa hoja');
  // Marcar como por_defecto: la elección "activa" del usuario se convierte en
  // su próxima sesión por defecto. El admin puede sobreescribir después.
  links.forEach(l => {
    if (String(l.username || '').trim().toLowerCase() === username.toLowerCase()) l.por_defecto = false;
  });
  const target = links.find(l =>
    String(l.username || '').trim().toLowerCase() === username.toLowerCase() &&
    String(l.spreadsheet_id) === id
  );
  if (target) {
    target.por_defecto = true;
    escribirHojasUsuarios_(links);
  }
  _currentSheetId = id;
  return { ok: true, hojaActivaId: id };
}

function obtenerConfigSheets() {
  return {
    entorno: 'production',
    sheet_id: obtenerSheetIdConfigurado_(),
    sheet_id_production: obtenerSheetIdConfigurado_(),
    auth_sheet_id: obtenerAuthSheetIdConfigurado_(),
    auth_sheet_id_production: obtenerAuthSheetIdConfigurado_()
  };
}

function setEntorno() {
  requireAdmin_();
  ss_();
  authSs_();
  return obtenerConfigSheets();
}

function setSheetIdEntorno(sheetId) {
  requireAdmin_();
  const id = String(sheetId || '').trim();
  if (!id) throw new Error('Debes indicar un sheetId válido.');
  SpreadsheetApp.openById(id); // Validación temprana.
  guardarSheetIdParaEntorno_(id);
  return obtenerConfigSheets();
}

function setAuthSheetIdEntorno(sheetId) {
  requireAdmin_();
  const id = String(sheetId || '').trim();
  if (!id) throw new Error('Debes indicar un sheetId válido.');
  SpreadsheetApp.openById(id); // Validación temprana.
  guardarAuthSheetIdParaEntorno_(id);
  asegurarAuthHojaUsuarios_();
  return obtenerConfigSheets();
}

function setSheetIdPruebas(sheetId) {
  return setSheetIdEntorno(sheetId);
}

function setSheetIdProduccion(sheetId) {
  return setSheetIdEntorno(sheetId);
}

function setAuthSheetIdPruebas(sheetId) {
  return setAuthSheetIdEntorno(sheetId);
}

function setAuthSheetIdProduccion(sheetId) {
  return setAuthSheetIdEntorno(sheetId);
}

function resetSheet() {
  requireAdmin_();
  guardarSheetIdParaEntorno_('');
  const ss = ss_();
  return { entorno: 'production', sheet_id: ss.getId() };
}

function username_() {
  return requireUsuario_();
}
function mostrarTodosLosOwners_() {
  return true;
}
function filasVisibles_(nombre) {
  const filas = leerHoja(nombre);
  if (mostrarTodosLosOwners_()) return filas;
  const owner = username_();
  return filas.filter(f => f.owner === owner);
}
function listarOwnersConDatos_() {
  const owners = new Set();
  HOJAS.forEach(nombre => {
    leerHoja(nombre).forEach(fila => {
      const owner = String(fila.owner || '').trim();
      if (owner) owners.add(owner);
    });
  });
  return [...owners];
}
function uid_(prefixo) { return (prefixo || 'id') + '_' + Utilities.getUuid().slice(0, 8); }
// Normaliza URL: si el usuario no escribió esquema, asume https://. Acepta
// sólo http(s); devuelve el origen canónico (esqueme + host, sin path) o
// cadena vacía si no es http(s).
function normalizarWeb_(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return '';
  const withScheme = /^https?:\/\//i.test(s) ? s : 'https://' + s;
  const m = /^https?:\/\/([^\/?#]+)/i.exec(withScheme);
  return m ? m[0].replace(/\/+$/, '') : '';
}
// Timezone canónico de la app = timezone de la hoja activa.
// Configura la hoja y el proyecto Apps Script en Europe/Madrid para evitar desfases.
// Cacheado por _currentSheetId porque distintos usuarios pueden tener hojas en
// distintos timezones (y _currentSheetId cambia por request).
let _tzCache = null;
let _tzCacheKey = null;
function tz_() {
  const key = _currentSheetId || 'master';
  if (_tzCache && _tzCacheKey === key) return _tzCache;
  try { _tzCache = ssActiva_().getSpreadsheetTimeZone(); }
  catch (e) { _tzCache = Session.getScriptTimeZone() || 'Europe/Madrid'; }
  _tzCacheKey = key;
  return _tzCache;
}
// Hoy en el timezone de la hoja (NO UTC).
function isoHoy_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}
function isoAhora_() {
  return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd'T'HH:mm:ss");
}

function asegurarHoja(nombre) {
  const ss = ssActiva_();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, SCHEMA[nombre].length).setValues([SCHEMA[nombre]]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, SCHEMA[nombre].length).setValues([SCHEMA[nombre]]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

function asegurarEsquema() {
  HOJAS.forEach(asegurarHoja);
}

// Añade columnas nuevas a hojas existentes sin tocar filas. Idempotente.
// Renombra la columna legacy a la primera columna de SCHEMA al migrar hojas pobladas.
function migrarEsquema() {
  HOJAS.forEach(asegurarHoja);
  Object.keys(SCHEMA).forEach(nombre => {
    const h = ssActiva_().getSheetByName(nombre);
    if (!h) return;
    const cab = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
    const legacyName = ['o','wner_email'].join('');
    const idxLegacy = cab.indexOf(legacyName);
    const idxTarget = cab.indexOf(SCHEMA[nombre][0]);
    if (idxLegacy >= 0 && idxTarget < 0) h.getRange(1, idxLegacy + 1).setValue(SCHEMA[nombre][0]);
    const cabPost = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
    SCHEMA[nombre].forEach((col, i) => {
      if (!cabPost.includes(col)) h.getRange(1, Math.max(h.getLastColumn(), i) + 1).setValue(col);
    });
  });
}

// Sheets convierte automáticamente las fechas en formato texto ('yyyy-MM-dd')
// a objetos Date. Al leer las normalizamos de vuelta a string para que
// String(fecha).slice(0,7) y las comparaciones de fechas funcionen igual
// que cuando se escribieron.
function normalizarValor_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  return v;
}

// Cache por invocación para evitar múltiples lecturas de la misma hoja
// dentro de una única ejecución (p. ej. bootstrapBase).
const _sheetReadCache = {};
function cloneRows_(rows) {
  return (rows || []).map(r => Object.assign({}, r));
}
function invalidateSheetCache_(nombre) {
  if (nombre) delete _sheetReadCache[nombre];
}

// ───────── Caché de saldos de cuentas ─────────
// obtenerCuentas() recorre todas las transacciones para calcular saldo y
// evolución mensual. El resultado se cachea por spreadsheet activo:
//   1) memoria de la request (_cuentasComputedCache)
//   2) CacheService (saldos:<sheetId>) entre requests
// Se invalida al escribir Cuentas o Transacciones (vía escribirHoja).
let _cuentasComputedCache = null;
let _cuentasComputedSheetId = '';

function sheetIdParaCache_() {
  return String(_currentSheetId || obtenerSheetIdConfigurado_() || 'default');
}

// Generación de datos por spreadsheet: se incrementa al escribir Cuentas o
// Transacciones. Vive en la propia hoja de datos (pestaña `_meta`),
// no en Script Properties. La clave de Script Cache incluye la generación
// para no reutilizar saldos anteriores a la última mutación.
const META_SHEET_NAME = '_meta';
const META_DATA_VERSION_KEY = 'data_version';
// Caché en memoria de la request: evita releer `_meta` en cada obtenerCuentas.
const _dataVersionMem = {};

function asegurarMetaHoja_() {
  const ss = ssActiva_();
  let h = ss.getSheetByName(META_SHEET_NAME);
  if (!h) {
    h = ss.insertSheet(META_SHEET_NAME);
    h.getRange(1, 1, 1, 2).setValues([['clave', 'valor']]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, 2).setValues([['clave', 'valor']]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

function leerMetaValor_(clave) {
  try {
    const h = asegurarMetaHoja_();
    const last = h.getLastRow();
    if (last < 2) return '';
    const vals = h.getRange(2, 1, last, 2).getValues();
    const target = String(clave || '');
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '') === target) return String(vals[i][1] || '').trim();
    }
    return '';
  } catch (e) {
    return '';
  }
}

function escribirMetaValor_(clave, valor) {
  const h = asegurarMetaHoja_();
  const key = String(clave || '');
  const val = String(valor || '');
  const last = h.getLastRow();
  if (last >= 2) {
    const vals = h.getRange(2, 1, last, 2).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0] || '') === key) {
        h.getRange(2 + i, 2).setValue(val);
        return;
      }
    }
  }
  h.appendRow([key, val]);
}

function getDataVersion_(sheetId) {
  const id = String(sheetId || sheetIdParaCache_());
  if (_dataVersionMem[id] != null) return _dataVersionMem[id];
  const v = leerMetaValor_(META_DATA_VERSION_KEY) || '0';
  _dataVersionMem[id] = v;
  return v;
}

function bumpDataVersion_(sheetId) {
  const id = String(sheetId || sheetIdParaCache_());
  const v = String(Date.now()) + '-' + String(Math.floor(Math.random() * 1e6));
  try {
    escribirMetaValor_(META_DATA_VERSION_KEY, v);
  } catch (e) { /* best-effort */ }
  _dataVersionMem[id] = v;
  return v;
}

// ponytail: scope por spreadsheetId + usuario + data_version. Permite que dos
// usuarios compartan spreadsheet (caso colaborativo) sin colisionar: cada
// uno ve sus propios saldos cacheados. data_version (hoja _meta del sheet
// compartido) descarta entradas obsoletas cuando CUALQUIER usuario muta el
// sheet — un bump de version hace inaccesibles los caches de todos los
// usuarios hasta que se revaliden contra el sheet. invalidateSaldosCache_
// solo borra la entrada del usuario actual; las de los demás quedan
// huérfanas hasta que la version cambie o expire el TTL.
function saldosCacheKey_(sheetId) {
  const id = String(sheetId || sheetIdParaCache_());
  return 'saldos:' + id + ':' + (currentUser_() || 'anon') + ':' + getDataVersion_(id);
}

function cloneCuentasComputed_(data) {
  try {
    return JSON.parse(JSON.stringify(data || []));
  } catch (e) {
    return (data || []).map(function (c) {
      const copia = Object.assign({}, c);
      if (c.subcuentas) copia.subcuentas = c.subcuentas.map(function (s) { return Object.assign({}, s); });
      if (c.evolucion) copia.evolucion = c.evolucion.map(function (e) { return Object.assign({}, e); });
      return copia;
    });
  }
}

// Formato compacto v1 para Script Cache: arrays posicionales (menos claves
// repetidas → JSON más pequeño y mejor ratio gzip).
// Cuenta:  [id, nombre, tipo, moneda, icono, saldo_inicial, saldo, evolucion[], subcuentas[], establecimiento_id]
// Evolución: [mes, saldo]
// Subcuenta: [id, nombre, parent_id, saldo_inicial, saldo, orden]
function packSaldosCache_(cuentas) {
  return {
    v: 1,
    c: (cuentas || []).map(function (cta) {
      return [
        cta.id,
        cta.nombre,
        cta.tipo,
        cta.moneda,
        cta.icono || '',
        Number(cta.saldo_inicial || 0),
        Number(cta.saldo || 0),
        (cta.evolucion || []).map(function (e) {
          return [e.mes, Number(e.saldo || 0)];
        }),
        (cta.subcuentas || []).map(function (s) {
          return [
            s.id,
            s.nombre,
            s.parent_id || '',
            Number(s.saldo_inicial || 0),
            Number(s.saldo || 0),
            Number(s.orden || 99)
          ];
        }),
        cta.establecimiento_id || ''
      ];
    })
  };
}

function unpackSaldosCache_(packed) {
  if (!packed || packed.v !== 1 || !Array.isArray(packed.c)) return null;
  return packed.c.map(function (r) {
    return {
      id: r[0],
      nombre: r[1],
      tipo: r[2],
      moneda: r[3],
      icono: r[4] || '',
      saldo_inicial: Number(r[5] || 0),
      saldo: Number(r[6] || 0),
      evolucion: (r[7] || []).map(function (e) {
        return { mes: e[0], saldo: Number(e[1] || 0) };
      }),
      subcuentas: (r[8] || []).map(function (s) {
        return {
          id: s[0],
          nombre: s[1],
          parent_id: s[2] || '',
          saldo_inicial: Number(s[3] || 0),
          saldo: Number(s[4] || 0),
          orden: Number(s[5] || 99)
        };
      }),
      establecimiento_id: r[9] || ''
    };
  });
}

function invalidateSaldosCache_() {
  _cuentasComputedCache = null;
  _cuentasComputedSheetId = '';
  // Antes de subir la versión, borramos la entrada actual (best-effort).
  try { cacheRemove_(saldosCacheKey_()); } catch (e) { /* ignore */ }
  // Nueva generación → la siguiente lectura usa una clave distinta y no
  // puede reutilizar el JSON antiguo aunque CacheService aún lo conserve.
  bumpDataVersion_();
}

function leerHoja(nombre) {
  if (_sheetReadCache[nombre]) return cloneRows_(_sheetReadCache[nombre]);
  asegurarHoja(nombre);
  const h = ssActiva_().getSheetByName(nombre);
  const valores = h.getDataRange().getValues();
  if (valores.length < 2) {
    _sheetReadCache[nombre] = [];
    return [];
  }
  const cab = valores[0];
  const rows = valores.slice(1).map(fila => {
    const o = {};
    cab.forEach((k, i) => (o[k] = normalizarValor_(fila[i])));
    return o;
  });
  _sheetReadCache[nombre] = rows;
  return cloneRows_(rows);
}

function escribirHoja(nombre, filas) {
  const h = ssActiva_().getSheetByName(nombre);
  const cab = SCHEMA[nombre];
  h.clearContents();
  const matriz = [cab].concat(filas.map(f => cab.map(k => f[k] != null ? f[k] : '')));
  if (matriz.length) h.getRange(1, 1, matriz.length, cab.length).setValues(matriz);
  h.setFrozenRows(1);
  _sheetReadCache[nombre] = cloneRows_(filas);
  // Cualquier cambio en cuentas o movimientos invalida saldos materializados.
  if (nombre === 'Cuentas' || nombre === 'Transacciones') {
    invalidateSaldosCache_();
  }
}

// Escritura de una sola fila: actualiza o añade sin reescribir toda la hoja.
// Reduce latencia de guardado (especialmente Transacciones/Recurrentes grandes).
// IMPORTANTE: getRange(row, column, numRows, numColumns) usa conteos, NO
// índices finales. Para 1 fila: getRange(rowNum, 1, 1, nCols).
function upsertFila(nombre, fila) {
  const datos = leerHoja(nombre);
  const cab = SCHEMA[nombre];
  if (!cab) throw new Error('Hoja desconocida: ' + nombre);
  const nCols = cab.length;
  const idx = datos.findIndex(f => f.id === fila.id);
  if (idx >= 0) {
    fila = Object.assign({}, datos[idx], fila, { owner: datos[idx].owner });
    datos[idx] = fila;
    const h = ssActiva_().getSheetByName(nombre);
    // Fila de datos en sheet = idx + 2 (cabecera en 1)
    const rowNum = idx + 2;
    const valores = cab.map(k => fila[k] != null ? fila[k] : '');
    h.getRange(rowNum, 1, 1, nCols).setValues([valores]);
  } else {
    datos.push(fila);
    const h = ssActiva_().getSheetByName(nombre);
    const last = h.getLastRow();
    // Si la hoja está vacía o solo tiene cabecera parcial, asegurar cabecera.
    if (last < 1) {
      h.getRange(1, 1, 1, nCols).setValues([cab]).setFontWeight('bold');
      h.setFrozenRows(1);
    }
    const rowNum = Math.max(last, 1) + 1;
    const valores = cab.map(k => fila[k] != null ? fila[k] : '');
    h.getRange(rowNum, 1, 1, nCols).setValues([valores]);
  }
  _sheetReadCache[nombre] = cloneRows_(datos);
  if (nombre === 'Cuentas' || nombre === 'Transacciones') {
    invalidateSaldosCache_();
  }
  return fila;
}

function eliminarFila(nombre, id) {
  const datos = leerHoja(nombre).filter(f => f.id !== id);
  escribirHoja(nombre, datos);
}

function sembrar(owner, seed) {
  HOJAS.forEach(asegurarHoja);
  // Si ya tiene algo, no duplicar
  if (leerHoja('Cuentas').some(c => c.owner === owner)) return;

  const fuente = seed || SEMILLA;
  const fuenteCuentas = seed ? seed.cuentas : SEMILLA.Cuentas;
  const fuenteCategorias = seed ? seed.categorias : SEMILLA.Categorias;

  const cuentas = fuenteCuentas.flatMap(c => {
    const parent = Object.assign({
      owner: owner, id: uid_('cta'), parent_id: '', oculta: false, fecha_creacion: isoAhora_()
    }, c);
    return [parent].concat((c.subcuentas || []).map((s, i) => Object.assign({
      owner: owner, id: uid_('cta'), parent_id: parent.id,
      tipo: parent.tipo, moneda: parent.moneda, icono: 'savings',
      saldo_inicial: 0, orden: i + 1, oculta: false, fecha_creacion: isoAhora_()
    }, s)));
  });
  escribirHoja('Cuentas', leerHoja('Cuentas').concat(cuentas));

  const cats = fuenteCategorias.map(c => Object.assign({
    owner: owner, id: uid_('cat')
  }, c));
  escribirHoja('Categorias', leerHoja('Categorias').concat(cats));
}

// ───────── Bootstrap público ─────────
const API_ACTIONS = new Set([
  'authStatus', 'loginUsuario', 'logoutUsuario', 'bootstrap', 'bootstrapBase',
  'listarUsuariosAdmin', 'crearUsuarioAdmin', 'resetearContrasenaAdmin', 'cambiarRolUsuarioAdmin', 'eliminarUsuarioAdmin', 'cambiarMiContrasena',
  'listarSpreadsheetsAdmin', 'altaSpreadsheetAdmin', 'bajaSpreadsheetAdmin', 'resetearSpreadsheetAdmin', 'renombrarSpreadsheetAdmin',
  'listarVinculacionesAdmin', 'vincularHojaUsuarioAdmin', 'desvincularHojaUsuarioAdmin', 'setHojaPorDefectoAdmin',
  'listarMisHojas', 'cambiarHojaActiva', 'obtenerConfigSheets', 'setEntorno', 'setSheetIdEntorno', 'setAuthSheetIdEntorno', 'resetSheet',
  'obtenerCuentas', 'guardarCuenta', 'eliminarCuenta', 'reordenarSubcuentas', 'reordenarCuentas',
  'obtenerCategorias', 'guardarCategoria', 'eliminarCategoria', 'reordenarCategorias',
  'obtenerEstablecimientos', 'guardarEstablecimiento', 'eliminarEstablecimiento',
  'obtenerTransacciones', 'guardarTransaccion', 'eliminarTransaccion',
  'obtenerRecurrentes', 'guardarRecurrente', 'eliminarRecurrente', 'generarRecurrentesPendientes',
  'obtenerPresupuestos', 'guardarPresupuesto', 'eliminarPresupuesto', 'conciliar', 'obtenerConciliaciones', 'editarConciliacion', 'eliminarConciliacion',
  'obtenerResumen', 'obtenerResumenEstablecimientos', 'obtenerCategoriasResumen', 'guardarTipoCambio', 'ejecutarSelfTestAdmin', 'ping', 'configurarSpreadsheetMaestro'
]);

function ping() {
  if (!getMasterSpreadsheetId_()) {
    return { ok: true, configurado: false };
  }
  return { ok: true, configurado: true, version: APP_VERSION };
}

function configurarSpreadsheetMaestro(id) {
  const limpio = String(id || '').trim();
  if (!/^[A-Za-z0-9_-]{20,}$/.test(limpio)) throw new Error('ID de spreadsheet inválido');
  SpreadsheetApp.openById(limpio);
  guardarAuthSheetIdParaEntorno_(limpio);
  return { ok: true, configurado: true };
}
const API_PUBLIC_ACTIONS = new Set(['authStatus', 'loginUsuario', 'logoutUsuario', 'ping', 'configurarSpreadsheetMaestro']);

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function dispatchApi_(request) {
  const action = String(request.action || '').trim();
  const args = Array.isArray(request.args) ? request.args : [];
  if (!API_ACTIONS.has(action)) throw new Error('Acción no permitida: ' + action);
  const fn = globalThis[action];
  if (API_PUBLIC_ACTIONS.has(action)) return fn.apply(null, args);
  return _authadmin(request.token, action, ...args);
}

function __selfTestApi_() {
  if (!API_ACTIONS.has('bootstrap') || API_ACTIONS.has('leerHoja')) throw new Error('API allowlist inválida');
  if (!API_PUBLIC_ACTIONS.has('loginUsuario') || API_PUBLIC_ACTIONS.has('bootstrap')) throw new Error('API auth inválida');
  return { ok: true };
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (!params.action) {
    return jsonResponse_({ ok: true, data: { service: 'finanzas-familia-api', version: APP_VERSION } });
  }
  try {
    return jsonResponse_({ ok: true, data: dispatchApi_({
      action: params.action,
      token: String(params.token || ''),
      args: params.args ? JSON.parse(params.args) : []
    }) });
  } catch (error) {
    return jsonResponse_({ ok: false, error: String(error && error.message || error) });
  }
}

function doPost(e) {
  try {
    const request = JSON.parse(e && e.postData && e.postData.contents || '{}');
    return jsonResponse_({ ok: true, data: dispatchApi_(request) });
  } catch (error) {
    return jsonResponse_({ ok: false, error: String(error && error.message || error) });
  }
}

function bootstrap() {
  const base = bootstrapBase();
  if (!((base.hojas || []).length) || !base.hojaActivaId) {
    base.transacciones = [];
    return base;
  }
  base.transacciones = obtenerTransacciones({});
  return base;
}

// Bootstrap ligero para carga progresiva en el cliente.
// Devuelve todo menos transacciones para que la UI pinte antes.
function bootstrapBase() {
  const owner = username_();
  asegurarUsuarios_();
  const hojasUsuario = listarHojasDelUsuario_(owner);
  const hojaActivaId = _currentSheetId || resolverHojaActivaId_(owner);

  if (!hojasUsuario.length || !hojaActivaId) {
    _currentSheetId = '';
    // Contrato: cuando el usuario autenticado no tiene ninguna hoja vinculada,
    // el bootstrap devuelve estado vacío y dos flags explícitos:
    //   - sin_hojas: true → no hay hojas asignadas a este usuario.
    //   - sin_datos_financieros: true → no hay datos que mostrar.
    // El frontend usa esto para NO mostrar ninguna vista de datos al usuario
    // (excepto la vista admin cuando el rol lo permite).
    return {
      sesion: { user: owner, rol: currentRol_() || ROLES.BASICO },
      version: APP_VERSION,
      data_version: '',
      hojas: [],
      hojaActivaId: '',
      cuentas: [],
      categorias: [],
      establecimientos: [],
      recurrentes: [],
      presupuestos: [],
      resumen: null,
      sin_hojas: true,
      sin_datos_financieros: true
    };
  }

  _currentSheetId = hojaActivaId;
  migrarEsquema();
  let owners = listarOwnersConDatos_();
  if (!owners.length) {
    sembrar(owner);
    owners = [owner];
  }
  normalizarCuentasSinSubcuentas_();
  normalizarSubcuentasHuerfanas_();
  owners.forEach(ownerFila => {
    generarRecurrentesPendientes_(ownerFila, new Date());
  });
  return {
    sesion: { user: owner, rol: currentRol_() || ROLES.BASICO },
    version: APP_VERSION,
    data_version: getDataVersion_(hojaActivaId),
    hojas: hojasUsuario,
    hojaActivaId: hojaActivaId,
    cuentas: obtenerCuentas(),
    categorias: obtenerCategorias(),
    establecimientos: obtenerEstablecimientos(),
    recurrentes: obtenerRecurrentes(),
    presupuestos: obtenerPresupuestos(),
    resumen: obtenerResumen()
  };
}

function listarHojasDelUsuario_(username) {
  const links = leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() === String(username || '').trim().toLowerCase());
  const todas = leerSpreadsheets_();
  const porId = {};
  todas.forEach(s => { porId[String(s.spreadsheet_id)] = s; });
  return links.map(l => {
    const meta = porId[String(l.spreadsheet_id)] || {};
    return {
      spreadsheet_id: String(l.spreadsheet_id),
      nombre: meta.nombre || '(sin nombre)',
      descripcion: meta.descripcion || '',
      por_defecto: String(l.por_defecto) === 'true' || l.por_defecto === true
    };
  });
}

function incluir(html) { return HtmlService.createHtmlOutputFromFile(html).getContent(); }

// ───────── Cuentas ─────────
function obtenerCuentas() {
  const sheetId = sheetIdParaCache_();
  // 1) Memoria de la request (bootstrap llama obtenerCuentas varias veces).
  if (_cuentasComputedCache && _cuentasComputedSheetId === sheetId) {
    return cloneCuentasComputed_(_cuentasComputedCache);
  }
  // 2) Script Cache entre requests del mismo spreadsheet (formato compacto).
  const cachedPacked = cacheGetJson_(saldosCacheKey_(sheetId));
  if (cachedPacked) {
    const cached = unpackSaldosCache_(cachedPacked) || (Array.isArray(cachedPacked) ? cachedPacked : null);
    if (cached) {
      _cuentasComputedCache = cached;
      _cuentasComputedSheetId = sheetId;
      return cloneCuentasComputed_(cached);
    }
  }

  const cuentas = filasVisibles_('Cuentas').filter(c => !c.oculta);
  const txs = leerHoja('Transacciones');
  const fmtMes = d => Utilities.formatDate(d, tz_(), 'yyyy-MM');

  const top = cuentas.filter(c => !c.parent_id).sort((a, b) => Number(a.orden || 99) - Number(b.orden || 99));
  const result = top.map(c => {
    const subs = cuentas.filter(x => x.parent_id === c.id);
    const subIds = new Set(subs.map(x => x.id));
    const parentInitial = Number(c.saldo_inicial || 0);

    // Txs del padre: golpean esta cuenta y NO están asignadas a una subcuenta
    // propia de esta cuenta (ni como origen ni como destino de transferencia).
    // Una subcuenta "huérfana" (de otra cuenta) se trata como movimiento del
    // padre, no se atribuye a la cuenta ajena.
    const parentTxs = txs.filter(t => {
      const origen = t.cuenta_id === c.id;
      const destino = t.cuenta_destino_id === c.id;
      if (!origen && !destino) return false;
      if (origen && t.subcuenta_id && subIds.has(t.subcuenta_id)) return false;
      if (destino) {
        // Cualquier subcuenta del reparto que sea hija de este padre cuenta
        // como sub-tx (no como tx del padre).
        const reparto = parseRepartoDestino_(t.reparto_destino);
        const subDest = reparto.length
          ? reparto.map(r => r.subcuenta_id)
          : (t.subcuenta_destino_id ? [t.subcuenta_destino_id] : []);
        if (subDest.some(id => subIds.has(id))) return false;
      }
      return true;
    });

    const subcuentas = subs.map(s => {
      const subInitial = Number(s.saldo_inicial || 0);
      // Cuenta como movimiento de la subcuenta cuando es su origen (cuenta_id)
      // o el destino de una transferencia (cuenta_destino_id o reparto_destino).
      // En ambos casos la cuenta implicada debe ser este mismo padre.
      const subDelta = txs
        .filter(t => {
          const origenOk = t.subcuenta_id === s.id && t.cuenta_id === c.id;
          if (origenOk) return true;
          if (t.cuenta_destino_id !== c.id || t.tipo !== 'transferencia') return false;
          const reparto = parseRepartoDestino_(t.reparto_destino);
          if (reparto.length) return reparto.some(r => r.subcuenta_id === s.id);
          return t.subcuenta_destino_id === s.id;
        })
        .reduce((sum, t) => sum + deltaSubcuenta_(t, s.id), 0);
      return {
        id: s.id, nombre: s.nombre, parent_id: c.id,
        saldo_inicial: subInitial,
        saldo: subInitial + subDelta,
        orden: Number(s.orden || 99)
      };
    }).sort((a, b) => a.orden - b.orden);

    // Modelo aditivo: el saldo del padre es su propio saldo más el saldo
    // completo de cada subcuenta (saldo_inicial + delta de transacciones).
    const subTotal = subcuentas.reduce((s, x) => s + x.saldo, 0);
    const subInitialTotal = subcuentas.reduce((s, x) => s + x.saldo_inicial, 0);
    const parentDelta = parentTxs.reduce((s, t) => s + deltaCuenta_(t, c.id), 0);

    const hoy = new Date();
    const evolucion = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
      const k = fmtMes(d);
      const parentDeltaK = parentTxs
        .filter(t => String(t.fecha).slice(0, 7) <= k)
        .reduce((s, t) => s + deltaCuenta_(t, c.id), 0);
      const subDeltaK = subcuentas.reduce((s, sub) => {
        const dK = txs
          .filter(t => {
            const fechaOk = String(t.fecha).slice(0, 7) <= k;
            if (!fechaOk) return false;
            if (t.subcuenta_id === sub.id && t.cuenta_id === c.id) return true;
            if (t.cuenta_destino_id !== c.id || t.tipo !== 'transferencia') return false;
            const reparto = parseRepartoDestino_(t.reparto_destino);
            if (reparto.length) return reparto.some(r => r.subcuenta_id === sub.id);
            return t.subcuenta_destino_id === sub.id;
          })
          .reduce((sd, t) => sd + deltaSubcuenta_(t, sub.id), 0);
        return s + dK;
      }, 0);
      evolucion.push({ mes: k, saldo: parentInitial + parentDeltaK + subInitialTotal + subDeltaK });
    }

    return {
      id: c.id, nombre: c.nombre, tipo: c.tipo, moneda: c.moneda, icono: c.icono, establecimiento_id: c.establecimiento_id || '',
      saldo_inicial: parentInitial,
      saldo: parentInitial + parentDelta + subTotal,
      evolucion,
      subcuentas
    };
  });

  _cuentasComputedCache = result;
  _cuentasComputedSheetId = sheetId;
  cachePutJson_(saldosCacheKey_(sheetId), packSaldosCache_(result), CACHE_TTL_SALDOS_SEC);
  return cloneCuentasComputed_(result);
}

// Variación de saldo que aporta una transacción a una cuenta concreta.
function deltaCuenta_(t, cuentaId) {
  if (t.cuenta_id === cuentaId) {
    if (t.tipo === 'ingreso' || t.tipo === 'devolucion') return Number(t.importe || 0);
    if (t.tipo === 'gasto' || t.tipo === 'transferencia') return -Number(t.importe || 0);
    return 0;
  }
  if (t.cuenta_destino_id === cuentaId && t.tipo === 'transferencia') {
    return Number(t.importe_destino || t.importe || 0);
  }
  return 0;
}

// Variación de saldo que aporta una transacción a una subcuenta concreta.
// Gasto/ingreso usan subcuenta_id. Las transferencias restan de la subcuenta
// de origen (subcuenta_id) y suman a la subcuenta de destino: si hay
// reparto_destino (JSON con varias subcuentas) se reparte; si no, se usa el
// legado subcuenta_destino_id único.
function deltaSubcuenta_(t, subId) {
  let d = 0;
  if (t.subcuenta_id === subId) {
    if (t.tipo === 'ingreso' || t.tipo === 'devolucion') d += Number(t.importe || 0);
    else if (t.tipo === 'gasto' || t.tipo === 'transferencia') d += -Number(t.importe || 0);
  }
  if (t.tipo === 'transferencia') {
    const reparto = parseRepartoDestino_(t.reparto_destino);
    if (reparto.length) {
      const r = reparto.find(x => x.subcuenta_id === subId);
      if (r) d += Number(r.importe || 0);
    } else if (t.subcuenta_destino_id === subId) {
      d += Number(t.importe_destino || t.importe || 0);
    }
  }
  return d;
}

// Sanea referencias subcuenta_id huérfanas: si una transacción apunta a una
// subcuenta que no pertenece a su propia cuenta (datos heredados o corruptos),
// se borra la subcuenta_id para que el movimiento quede solo en su cuenta.
// Idempotente: solo escribe cuando encuentra algo que corregir.
function normalizarSubcuentasHuerfanas_() {
  const cuentas = leerHoja('Cuentas');
  const validas = new Set(
    cuentas.filter(c => c.parent_id).map(c => c.parent_id + '|' + c.id)
  );
  let cambios = false;
  const txs = leerHoja('Transacciones').map(t => {
    let nuevo = t;
    if (t.subcuenta_id && !validas.has(t.cuenta_id + '|' + t.subcuenta_id)) {
      cambios = true;
      nuevo = Object.assign({}, nuevo, { subcuenta_id: '' });
    }
    if (t.subcuenta_destino_id && !validas.has(t.cuenta_destino_id + '|' + t.subcuenta_destino_id)) {
      cambios = true;
      nuevo = Object.assign({}, nuevo, { subcuenta_destino_id: '' });
    }
    // Reparto destino: cualquier entrada cuya subcuenta no pertenezca a la
    // cuenta destino se descarta. Si tras limpiar queda vacío o el total no
    // cuadra, se elimina el reparto entero.
    const reparto = parseRepartoDestino_(t.reparto_destino);
    if (reparto.length && t.cuenta_destino_id) {
      const filtrado = reparto.filter(r => validas.has(t.cuenta_destino_id + '|' + r.subcuenta_id));
      const suma = filtrado.reduce((s, r) => s + r.importe, 0);
      const totalEsperado = t.importe_destino ? Number(t.importe_destino) : Number(t.importe || 0);
      const ok = filtrado.length > 0 && Math.abs(suma - totalEsperado) <= 0.01;
      const nuevoJson = ok ? JSON.stringify(filtrado) : '';
      if (nuevoJson !== (t.reparto_destino || '')) {
        cambios = true;
        nuevo = Object.assign({}, nuevo, { reparto_destino: nuevoJson });
      }
    }
    return nuevo;
  });
  if (cambios) escribirHoja('Transacciones', txs);
  return cambios;
}

function crearSubcuentaDefault_(parent) {
  if (!parent || parent.parent_id) return null;
  const fila = {
    owner: parent.owner,
    id: uid_('cta'),
    parent_id: parent.id,
    nombre: 'General',
    tipo: parent.tipo,
    moneda: parent.moneda,
    icono: 'savings',
    saldo_inicial: 0,
    orden: 1,
    oculta: false,
    fecha_creacion: isoAhora_()
  };
  upsertFila('Cuentas', fila);
  return fila;
}

function asegurarSubcuentaDefaultCuenta_(parentId) {
  const todas = leerHoja('Cuentas');
  const parent = todas.find(c => c.id === parentId && !c.parent_id);
  if (!parent) return false;
  const subs = todas.filter(c => c.parent_id === parentId);
  if (subs.length > 0) return false;
  crearSubcuentaDefault_(parent);
  return true;
}

function normalizarCuentasSinSubcuentas_() {
  const todas = leerHoja('Cuentas');
  const top = todas.filter(c => !c.parent_id);
  const byParent = {};
  todas.filter(c => c.parent_id).forEach(s => {
    byParent[s.parent_id] = (byParent[s.parent_id] || 0) + 1;
  });
  let cambios = false;
  top.forEach(parent => {
    if (!byParent[parent.id]) {
      crearSubcuentaDefault_(parent);
      cambios = true;
    }
  });
  return cambios;
}

function guardarCuenta(cuenta) {
  const owner = username_();
  if (!cuenta || !cuenta.nombre) throw new Error('Nombre de cuenta obligatorio');
  const todas = leerHoja('Cuentas');
  let parent = null;
  if (cuenta.parent_id) {
    parent = todas.find(c => c.id === cuenta.parent_id);
    if (!parent) throw new Error('Cuenta padre no encontrada');
    if (parent.parent_id) throw new Error('Una subcuenta no puede tener subcuentas');
    // Subcuentas heredan tipo y moneda del padre; ignoramos lo que mande el cliente.
    cuenta.tipo = parent.tipo;
    cuenta.moneda = parent.moneda;
  } else {
    if (!['activo', 'pasivo'].includes(cuenta.tipo)) throw new Error('Tipo inválido');
    if (!cuenta.moneda) throw new Error('Moneda obligatoria');
  }
  // Al editar, preservamos metadatos que el cliente no envía (orden, icono y
  // fecha de creación) para no reiniciarlos y romper el orden de subcuentas.
  const existente = cuenta.id ? todas.find(c => c.id === cuenta.id) : null;
  // Subcuentas no admiten establecimiento propio (sólo lo gestiona la cuenta padre).
  const estId = cuenta.parent_id
    ? ''
    : (function () {
        const v = String(cuenta.establecimiento_id == null ? '' : cuenta.establecimiento_id).trim();
        if (!v) return '';
        const ok = leerHoja('Establecimientos').some(e => e.id === v);
        if (!ok) throw new Error('Establecimiento no encontrado');
        return v;
      })();
  const fila = {
    owner: (existente && existente.owner) || owner,
    id: cuenta.id || uid_('cta'),
    parent_id: cuenta.parent_id || '',
    nombre: String(cuenta.nombre).trim(),
    tipo: cuenta.tipo,
    moneda: String(cuenta.moneda).toUpperCase(),
    icono: cuenta.icono || (existente && existente.icono) || 'account_balance_wallet',
    saldo_inicial: Number(cuenta.saldo_inicial || 0),
    orden: cuenta.orden || (existente && existente.orden) || 99,
    oculta: !!cuenta.oculta,
    establecimiento_id: estId,
    fecha_creacion: (existente && existente.fecha_creacion) || isoAhora_()
  };
  upsertFila('Cuentas', fila);
  if (!fila.parent_id) asegurarSubcuentaDefaultCuenta_(fila.id);
  return obtenerCuentas();
}

function eliminarCuenta(id) {
  const todas = leerHoja('Cuentas');
  const cuenta = todas.find(c => c.id === id);
  if (!cuenta) throw new Error('Cuenta no encontrada');
  const esSub = !!cuenta.parent_id;
  const idsAEliminar = new Set([id]);
  if (!esSub) {
    todas.filter(c => c.parent_id === id).forEach(s => idsAEliminar.add(s.id));
  }
  if (esSub) {
    const subsHermanas = todas.filter(c => c.parent_id === cuenta.parent_id);
    if (subsHermanas.length <= 1) throw new Error('Cada cuenta debe tener al menos una subcuenta.');
  }
  // Top-level: se permite borrar junto con sus subcuentas para no romper
  // la regla de "siempre al menos una subcuenta" mientras la cuenta exista.
  const txs = leerHoja('Transacciones').filter(t => (
    (t.cuenta_id === id || t.cuenta_destino_id === id) ||
    (esSub && t.subcuenta_id === id && t.cuenta_id === cuenta.parent_id) ||
    (esSub && t.subcuenta_destino_id === id && t.cuenta_destino_id === cuenta.parent_id)
  ));
  if (txs.length) throw new Error((esSub ? 'La subcuenta' : 'La cuenta') + ' tiene ' + txs.length + ' movimientos. Reasígnalos o elimínalos primero.');
  if (esSub) {
    eliminarFila('Cuentas', id);
  } else {
    const restantes = leerHoja('Cuentas').filter(c => !idsAEliminar.has(c.id));
    escribirHoja('Cuentas', restantes);
  }
  return obtenerCuentas();
}

function reordenarSubcuentas(parentId, ids) {
  if (!parentId || !Array.isArray(ids)) throw new Error('Parámetros inválidos');
  const todas = leerHoja('Cuentas');
  const padre = todas.find(c => c.id === parentId);
  if (!padre || padre.parent_id) throw new Error('Cuenta padre inválida');
  ids.forEach(id => {
    if (!todas.some(c => c.id === id && c.parent_id === parentId)) {
      throw new Error('Subcuenta ' + id + ' no pertenece a la cuenta padre');
    }
  });
  ids.forEach((id, i) => {
    const idx = todas.findIndex(c => c.id === id);
    todas[idx].orden = i + 1;
  });
  escribirHoja('Cuentas', todas);
  return obtenerCuentas();
}

function reordenarCuentas(ids) {
  if (!Array.isArray(ids)) throw new Error('Parámetros inválidos');
  const todas = leerHoja('Cuentas');
  ids.forEach(id => {
    if (!todas.some(c => c.id === id && !c.parent_id)) {
      throw new Error('Cuenta inválida');
    }
  });
  ids.forEach((id, i) => {
    const idx = todas.findIndex(c => c.id === id);
    todas[idx].orden = i + 1;
  });
  escribirHoja('Cuentas', todas);
  return obtenerCuentas();
}

// ───────── Categorías ─────────
function reordenarCategorias(ids) {
  if (!Array.isArray(ids)) throw new Error('Parámetros inválidos');
  const todas = leerHoja('Categorias');
  ids.forEach(id => {
    if (!todas.some(c => c.id === id)) throw new Error('Categoría inválida');
  });
  ids.forEach((id, i) => {
    todas.find(c => c.id === id).orden = i + 1;
  });
  escribirHoja('Categorias', todas);
  return obtenerCategorias();
}

function obtenerCategorias() {
  return filasVisibles_('Categorias').sort((a, b) => a.orden - b.orden);
}

function guardarCategoria(cat) {
  const owner = username_();
  const existente = cat.id ? leerHoja('Categorias').find(c => c.id === cat.id) : null;
  const fila = {
    owner: (existente && existente.owner) || owner, id: cat.id || uid_('cat'),
    nombre: String(cat.nombre).trim(), color: cat.color || '#00613e',
    icono: cat.icono || (existente && existente.icono) || 'category',
    tipo: cat.tipo || 'gasto', orden: cat.orden || (existente && existente.orden) || 99
  };
  upsertFila('Categorias', fila);
  return obtenerCategorias();
}

function eliminarCategoria(id) {
  eliminarFila('Categorias', id);
  return obtenerCategorias();
}

// ───────── Establecimientos ─────────
function obtenerEstablecimientos() {
  return filasVisibles_('Establecimientos')
    .sort((a, b) => String(a.nombre || '').localeCompare(String(b.nombre || ''), 'es', { sensitivity: 'base' }));
}

function guardarEstablecimiento(est) {
  const owner = username_();
  const input = est || {};
  const nombre = String(input.nombre || '').trim();
  if (!nombre) throw new Error('Nombre de establecimiento obligatorio');

  const todos = leerHoja('Establecimientos');
  const existente = input.id ? todos.find(e => e.id === input.id) : null;
  if (input.id && !existente) throw new Error('Establecimiento no encontrado');
  const nombreClave = nombre.toLowerCase();
  if (todos.some(e => e.id !== input.id && String(e.nombre || '').trim().toLowerCase() === nombreClave)) {
    throw new Error('Ya existe un establecimiento con ese nombre');
  }

  upsertFila('Establecimientos', {
    owner: (existente && existente.owner) || owner,
    id: existente ? existente.id : uid_('est'),
    nombre: nombre,
    web: normalizarWeb_(input.web)
  });
  return obtenerEstablecimientos();
}

function eliminarEstablecimiento(id) {
  const establecimientoId = String(id || '').trim();
  if (!leerHoja('Establecimientos').some(e => e.id === establecimientoId)) {
    throw new Error('Establecimiento no encontrado');
  }
  if (leerHoja('Transacciones').some(t => t.establecimiento_id === establecimientoId)) {
    throw new Error('No se puede eliminar: tiene movimientos asociados');
  }
  const usadoEnRecurrente = leerHoja('Recurrentes').some(r => {
    try { return JSON.parse(r.plantilla || '{}').establecimiento_id === establecimientoId; }
    catch (_) { return false; }
  });
  if (usadoEnRecurrente) throw new Error('No se puede eliminar: tiene plantillas recurrentes asociadas');

  eliminarFila('Establecimientos', establecimientoId);
  return obtenerEstablecimientos();
}

// ───────── Transacciones ─────────
// Parsea "yyyy-MM-dd" como fecha de calendario local (no UTC).
// new Date("yyyy-MM-dd") interpreta medianoche UTC y desplaza el día en España.
function parseFecha(s) {
  if (!s) return null;
  if (Object.prototype.toString.call(s) === '[object Date]') return s;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function iso_(d) {
  return Utilities.formatDate(new Date(d), tz_(), 'yyyy-MM-dd');
}

// Reparto destino: array de {subcuenta_id, importe} guardado como JSON en
// Transacciones.reparto_destino. Acepta string, array o null; devuelve array
// filtrando entradas inválidas. Vacío o null = reparto a nivel de cuenta.
function parseRepartoDestino_(raw) {
  if (!raw) return [];
  let parsed = raw;
  if (typeof raw === 'string') {
    try { parsed = JSON.parse(raw); } catch (e) { return []; }
  }
  if (!Array.isArray(parsed)) return [];
  return parsed
    .map(r => {
      const out = {
        subcuenta_id: r && r.subcuenta_id,
        importe: Number(r && r.importe || 0),
        categoria_id: (r && r.categoria_id) || ''
      };
      const periodoUso = Number(r && r.periodo_uso || 0);
      const montoUsar = Number(r && r.monto_a_usar || 0);
      if (periodoUso > 0) out.periodo_uso = periodoUso;
      if (montoUsar > 0) out.monto_a_usar = montoUsar;
      return out;
    })
    .filter(r => r.subcuenta_id && r.importe > 0);
}

function tipoTransferenciaPresupuesto_(cuentasById, cuentaOrigenId, cuentaDestinoId) {
  const origen = cuentasById[String(cuentaOrigenId || '')];
  const destino = cuentasById[String(cuentaDestinoId || '')];
  if (!origen || !destino) return 'neutro';
  if (origen.tipo === 'activo' && destino.tipo === 'pasivo') return 'gasto';
  if (origen.tipo === 'pasivo' && destino.tipo === 'activo') return 'ingreso';
  return 'neutro';
}

function tipoTransferenciaPresupuestoTx_(t, cuentasById) {
  if (!t || t.tipo !== 'transferencia') return 'neutro';
  return tipoTransferenciaPresupuesto_(cuentasById, t.cuenta_id, t.cuenta_destino_id);
}

function validarCategoriaPorTipo_(catsById, categoriaId, tipoEsperado, label) {
  const id = String(categoriaId || '').trim();
  if (!id) throw new Error((label || 'Categoría') + ' obligatoria para transferencias de ' + tipoEsperado);
  const cat = catsById[id];
  if (!cat) throw new Error((label || 'Categoría') + ' no encontrada');
  if (cat.tipo !== tipoEsperado) {
    throw new Error((label || 'Categoría') + ' debe ser de tipo ' + tipoEsperado);
  }
}

function validarCategoriasTransferencia_(tipoPresupuesto, categoriaId, reparto) {
  if (tipoPresupuesto !== 'gasto' && tipoPresupuesto !== 'ingreso') return;
  const catsById = {};
  leerHoja('Categorias')
    .forEach(c => { catsById[c.id] = c; });

  if (reparto && reparto.length) {
    reparto.forEach((r, i) => {
      validarCategoriaPorTipo_(catsById, r.categoria_id, tipoPresupuesto, 'Categoría del destino ' + (i + 1));
    });
    return;
  }
  const categoriaFallback = String(categoriaId || '').trim();
  if (categoriaFallback) {
    validarCategoriaPorTipo_(catsById, categoriaFallback, tipoPresupuesto, 'Categoría');
  }
}

function obtenerTransacciones(filtro) {
  filtro = filtro || {};
  let txs = filasVisibles_('Transacciones');
  if (filtro.cuenta_id) txs = txs.filter(t => t.cuenta_id === filtro.cuenta_id || t.cuenta_destino_id === filtro.cuenta_id);
  if (filtro.tipo) txs = txs.filter(t => t.tipo === filtro.tipo);
  if (filtro.estado) txs = txs.filter(t => t.estado === filtro.estado);
  if (filtro.categoria_id) txs = txs.filter(t => t.categoria_id === filtro.categoria_id);
  if (filtro.desde) txs = txs.filter(t => String(t.fecha) >= filtro.desde);
  if (filtro.hasta) txs = txs.filter(t => String(t.fecha) <= filtro.hasta);
  if (filtro.q) {
    const q = String(filtro.q).toLowerCase();
    txs = txs.filter(t => String(t.descripcion || '').toLowerCase().includes(q) || String(t.notas || '').toLowerCase().includes(q));
  }
  return txs
    .map(t => Object.assign({}, t, {
      importe: Number(t.importe || 0),
      importe_destino: Number(t.importe_destino || 0),
      ratio_conversion: Number(t.ratio_conversion || 0)
    }))
    .sort((a, b) => String(b.fecha).localeCompare(String(a.fecha)));
}

function guardarTransaccion(tx) {
  const owner = username_();
  const actor = requireUsuario_();
  if (!tx) throw new Error('Transacción vacía');
  const tipo = tx.tipo;
  if (!['gasto', 'ingreso', 'transferencia', 'devolucion'].includes(tipo)) throw new Error('Tipo inválido');
  if (!tx.cuenta_id) throw new Error('Cuenta obligatoria');
  if (!(Number(tx.importe) > 0)) throw new Error('Importe debe ser > 0');
  const fecha = parseFecha(tx.fecha);
  if (!fecha) throw new Error('Fecha inválida');
  const txs = leerHoja('Transacciones');
  const existente = tx.id ? txs.find(t => t.id === tx.id) : null;
  if (tx.id && !existente) throw new Error('Transacción no encontrada');
  // Concurrencia optimista: si el cliente envía la fecha_ultima_edicion que
  // conocía al abrir el formulario y el servidor tiene otra, hay conflicto.
  if (existente && tx.base_fecha_ultima_edicion != null && tx.base_fecha_ultima_edicion !== '') {
    const base = String(tx.base_fecha_ultima_edicion || '');
    const actual = String(existente.fecha_ultima_edicion || '');
    if (base !== actual) {
      throw new Error('CONFLICT:' + JSON.stringify({
        tipo: 'transaccion',
        id: existente.id,
        mensaje: 'Esta transacción fue modificada por otro usuario o en otra pestaña.',
        fecha_ultima_edicion: actual,
        ultima_edicion_por: existente.ultima_edicion_por || '',
        transaccion: existente
      }));
    }
  }
  const ownerTx = existente ? String(existente.owner || owner) : owner;
  const cuentasHoja = leerHoja('Cuentas');
  let establecimiento_id = '';
  const establecimientoSolicitado = String(tx.establecimiento_id || '').trim();
  if (tipo !== 'transferencia' && establecimientoSolicitado) {
    const valido = leerHoja('Establecimientos').some(e => e.id === establecimientoSolicitado);
    if (!valido) throw new Error('Establecimiento no encontrado');
    establecimiento_id = establecimientoSolicitado;
  }
  let subcuenta_id = '';
  if (tx.subcuenta_id) {
    const sub = cuentasHoja.find(c => c.id === tx.subcuenta_id && c.parent_id === tx.cuenta_id);
    if (!sub) throw new Error('Subcuenta no encontrada o no pertenece a la cuenta');
    subcuenta_id = tx.subcuenta_id;
  }
  let subcuenta_destino_id = '';
  let repartoDestinoJson = '';
  let repartoDestinoNormalizado = [];
  if (tipo === 'transferencia' && tx.cuenta_destino_id) {
    const repartoRaw = Array.isArray(tx.reparto_destino) ? tx.reparto_destino : null;
    if (repartoRaw && repartoRaw.length) {
      const subValidas = new Set(
        cuentasHoja.filter(c => c.parent_id === tx.cuenta_destino_id).map(c => c.id)
      );
      const visto = new Set();
      const normalizado = [];
      let suma = 0;
      repartoRaw.forEach(r => {
        const sid = r && r.subcuenta_id;
        if (!sid) throw new Error('Cada destino requiere subcuenta');
        if (!subValidas.has(sid)) throw new Error('Subcuenta destino no pertenece a la cuenta destino');
        if (visto.has(sid)) throw new Error('Subcuenta destino duplicada');
        const imp = Number(r.importe);
        if (!(imp > 0)) throw new Error('Importe de destino debe ser > 0');
        const periodoUso = Number(r.periodo_uso || 0);
        const montoUsar = Number(r.monto_a_usar || 0);
        if (periodoUso < 0) throw new Error('Periodo de uso no puede ser negativo');
        if (montoUsar < 0) throw new Error('Monto a usar no puede ser negativo');
        visto.add(sid);
        const filaReparto = { subcuenta_id: sid, importe: imp, categoria_id: (r.categoria_id || '') };
        if (periodoUso > 0) filaReparto.periodo_uso = periodoUso;
        if (montoUsar > 0) filaReparto.monto_a_usar = montoUsar;
        normalizado.push(filaReparto);
        suma += imp;
      });
      const totalEsperado = tx.importe_destino ? Number(tx.importe_destino) : Number(tx.importe);
      if (Math.abs(suma - totalEsperado) > 0.01) {
        throw new Error('La suma del reparto (' + suma.toFixed(2) + ') no coincide con el importe destino (' + totalEsperado.toFixed(2) + ')');
      }
      repartoDestinoJson = JSON.stringify(normalizado);
      repartoDestinoNormalizado = normalizado;
      // Backfill del campo legacy para que filtros por subcuenta sigan
      // encontrando la tx mientras conviven datos nuevos y viejos.
      subcuenta_destino_id = normalizado[0].subcuenta_id;
    } else if (tx.subcuenta_destino_id) {
      const subd = cuentasHoja.find(c => c.id === tx.subcuenta_destino_id && c.parent_id === tx.cuenta_destino_id);
      if (!subd) throw new Error('Subcuenta destino no encontrada o no pertenece a la cuenta destino');
      subcuenta_destino_id = tx.subcuenta_destino_id;
    }
  } else if (tx.subcuenta_destino_id) {
    throw new Error('Subcuenta destino solo aplica a transferencias');
  }
  if (tipo === 'transferencia') {
    const cuentasById = {};
    cuentasHoja.forEach(c => { cuentasById[c.id] = c; });
    const tipoPresupuesto = tipoTransferenciaPresupuesto_(cuentasById, tx.cuenta_id, tx.cuenta_destino_id);
    validarCategoriasTransferencia_(tipoPresupuesto, tx.categoria_id || '', repartoDestinoNormalizado);
  }
  if (tipo === 'devolucion') {
    // ponytail: la categoría de una devolución debe ser de tipo gasto
    // (revierte un gasto previo). Validación inline para evitar el texto
    // "para transferencias de" que mete validarCategoriaPorTipo_ para txs
    // de transferencia.
    const catsById = {};
    leerHoja('Categorias').forEach(c => { catsById[c.id] = c; });
    const idCat = String(tx.categoria_id || '').trim();
    if (!idCat) throw new Error('Categoría de devolución obligatoria');
    const cat = catsById[idCat];
    if (!cat) throw new Error('Categoría de devolución no encontrada');
    if (cat.tipo !== 'gasto') throw new Error('Categoría de devolución debe ser de tipo gasto');
  }
  const recExistenteId = existente ? (existente.recurrente_id || '') : '';
  const fila = {
    owner: ownerTx,
    id: tx.id || uid_('tx'),
    fecha: iso_(fecha),
    tipo,
    importe: Number(tx.importe),
    moneda: String(tx.moneda || 'EUR').toUpperCase(),
    cuenta_id: tx.cuenta_id,
    subcuenta_id: subcuenta_id,
    cuenta_destino_id: tx.cuenta_destino_id || '',
    subcuenta_destino_id: subcuenta_destino_id,
    importe_destino: tx.importe_destino ? Number(tx.importe_destino) : '',
    ratio_conversion: tx.ratio_conversion ? Number(tx.ratio_conversion) : '',
    reparto_destino: repartoDestinoJson,
    categoria_id: tx.categoria_id || '',
    establecimiento_id: establecimiento_id,
    descripcion: String(tx.descripcion || '').trim(),
    estado: tx.estado || 'pendiente',
    // ponytail: preservar el recurrente existente cuando el payload no lo
    // incluye (caso típico: edit de una tx recurrente sin tocar el switch).
    // Si el frontend manda tx.recurrente_id === '' explícitamente, eso gana
    // y desvincula la tx (la UI por separado llama eliminarRecurrente).
    recurrente_id: tx.recurrente_id !== undefined ? (tx.recurrente_id || '') : recExistenteId,
    fecha_pago: tx.fecha_pago ? iso_(tx.fecha_pago) : '',
    conciliada_con: tx.conciliada_con || '',
    notas: tx.notas || '',
    fecha_creacion: (existente && existente.fecha_creacion) || tx.fecha_creacion || isoAhora_(),
    ultima_edicion_por: actor,
    fecha_ultima_edicion: isoAhora_()
  };
  if (fila.tipo === 'transferencia' && !fila.cuenta_destino_id) throw new Error('Cuenta destino obligatoria en transferencia');
  // ponytail: misma cuenta se permite para mover entre subcuentas del mismo
  // padre. Si no hay subcuenta origen o el destino coincide, sería un no-op.
  if (fila.tipo === 'transferencia' && fila.cuenta_destino_id === fila.cuenta_id) {
    if (!subcuenta_id) throw new Error('Transferencia interna requiere subcuenta de origen');
    const destIds = repartoDestinoNormalizado.length
      ? repartoDestinoNormalizado.map(r => r.subcuenta_id)
      : (subcuenta_destino_id ? [subcuenta_destino_id] : []);
    if (!destIds.length) throw new Error('Transferencia interna requiere subcuenta de destino');
    if (destIds.includes(subcuenta_id)) throw new Error('La subcuenta destino no puede ser la misma que la de origen');
  }
  if (fila.tipo === 'transferencia' && fila.ratio_conversion && !(fila.importe_destino > 0)) {
    throw new Error('Falta importe destino o ratio de conversión');
  }
  upsertFila('Transacciones', fila);
  // ponytail: si la tx ya está vinculada a un recurrente, actualizar ese en
  // lugar de crear uno nuevo. Mantiene las txs generadas previas apuntando al
  // mismo id.
  let recurrentesActualizados = null;
  if (tx.recurrente_plantilla) {
    if (fila.recurrente_id) {
      actualizarPlantillaRecurrente_(owner, fila.recurrente_id, tx, fila);
      recurrentesActualizados = obtenerRecurrentes();
    } else {
      const nuevoId = upsertRecurrenteBase_(owner, tx);
      if (nuevoId) {
        fila.recurrente_id = nuevoId;
        upsertFila('Transacciones', fila);
      }
      recurrentesActualizados = obtenerRecurrentes();
    }
  } else if (tx.recurrente_id === '') {
    // Desvinculación explícita: el frontend puede haber eliminado el recurrente
    // por separado; devolvemos la lista actual para que la UI no quede stale.
    recurrentesActualizados = obtenerRecurrentes();
  }
  // Payload enriquecido para refresco ligero en el cliente.
  // Campos de la tx se mantienen en el top-level por compatibilidad (self-tests).
  // No recalculamos obtenerCuentas() aquí: es caro tras invalidar el cache de
  // saldos. El frontend mantiene un cache local de saldos y aplica el delta
  // de la tx (crear/editar) sin esperar este recálculo.
  return Object.assign({}, fila, {
    transaccion: fila,
    cuentas: null,
    recurrentes: recurrentesActualizados,
    data_version: getDataVersion_()
  });
}

// ponytail: actualiza la plantilla de un recurrente existente a partir de la
// tx. Preserva id, owner, activa, inicio, fin y ultima_generacion.
function actualizarPlantillaRecurrente_(owner, recurrenteId, tx, filaTx) {
  const datos = leerHoja('Recurrentes');
  const idx = datos.findIndex(r => r.id === recurrenteId);
  if (idx < 0) return;
  const actual = datos[idx];
  let plantilla;
  try { plantilla = JSON.parse(actual.plantilla); } catch (_) { plantilla = {}; }
  const nuevoReparto = Array.isArray(tx.reparto_destino)
    ? JSON.stringify(tx.reparto_destino)
    : (tx.reparto_destino || '');
  const upd = {
    tipo: filaTx.tipo,
    importe: Number(filaTx.importe),
    cuenta_id: filaTx.cuenta_id,
    subcuenta_id: filaTx.subcuenta_id || '',
    cuenta_destino_id: filaTx.cuenta_destino_id || '',
    subcuenta_destino_id: filaTx.subcuenta_destino_id || '',
    importe_destino: filaTx.importe_destino || '',
    ratio_conversion: filaTx.ratio_conversion || '',
    reparto_destino: nuevoReparto,
    categoria_id: filaTx.categoria_id || '',
    establecimiento_id: filaTx.establecimiento_id || '',
    descripcion: filaTx.descripcion || plantilla.descripcion || '',
    periodo_meses: Number(tx.recurrente_periodo_meses || plantilla.periodo_meses || 1),
    dia_mes: Number(tx.recurrente_dia || plantilla.dia_mes || 1)
  };
  // Conservar campos historicos que la tx no expone.
  const merged = Object.assign({}, plantilla, upd);
  if (!merged.inicio) merged.inicio = actual.ultima_generacion || iso_(filaTx.fecha) || plantilla.inicio || '';
  validarPlantillaRecurrente_(owner, merged);
  datos[idx].plantilla = JSON.stringify(merged);
  // ponytail: editar la tx original también reposiciona mes_inicio/anio_inicio
  // del recurrente a la fecha de esa tx. Útil cuando el usuario crea un
  // recurrente y luego decide "esto arrancó el mes X" sin tocar el sheet.
  const fechaRef = parseFecha(filaTx.fecha);
  if (fechaRef && !isNaN(fechaRef)) {
    datos[idx].mes_inicio = fechaRef.getMonth() + 1;
    datos[idx].anio_inicio = fechaRef.getFullYear();
  }
  // Actualizar solo la fila afectada (evita reescribir toda la hoja).
  upsertFila('Recurrentes', datos[idx]);
}

function eliminarTransaccion(id, baseFechaUltimaEdicion) {
  const txs = leerHoja('Transacciones');
  const existente = txs.find(t => t.id === id) || null;
  if (!existente) throw new Error('Transacción no encontrada');
  // Concurrencia optimista opcional en borrado.
  if (baseFechaUltimaEdicion != null && baseFechaUltimaEdicion !== '') {
    const base = String(baseFechaUltimaEdicion || '');
    const actual = String(existente.fecha_ultima_edicion || '');
    if (base !== actual) {
      throw new Error('CONFLICT:' + JSON.stringify({
        tipo: 'transaccion',
        id: existente.id,
        mensaje: 'Esta transacción fue modificada por otro usuario o en otra pestaña y ya no se puede eliminar con la versión local.',
        fecha_ultima_edicion: actual,
        ultima_edicion_por: existente.ultima_edicion_por || '',
        transaccion: existente
      }));
    }
  }
  eliminarFila('Transacciones', id);
  return { ok: true, data_version: getDataVersion_() };
}

// ───────── Recurrentes ─────────
function obtenerRecurrentes() {
  return filasVisibles_('Recurrentes').map(r => ({
    id: r.id, plantilla: r.plantilla, ultima_generacion: r.ultima_generacion, activa: r.activa === true || r.activa === 'true',
    mes_inicio: r.mes_inicio != null && r.mes_inicio !== '' ? Number(r.mes_inicio) : null,
    anio_inicio: r.anio_inicio != null && r.anio_inicio !== '' ? Number(r.anio_inicio) : null
  }));
}

function validarPlantillaRecurrente_(owner, plantilla) {
  const p = plantilla || {};
  if (!['gasto', 'ingreso', 'transferencia', 'devolucion'].includes(p.tipo)) throw new Error('Tipo inválido');
  if (!p.cuenta_id) throw new Error('Cuenta obligatoria');
  if (!(Number(p.importe) > 0)) throw new Error('Importe debe ser > 0');
  // ponytail: periodo en meses, entero 1-12. Compatibilidad con plantillas
  // legacy: periodicidad === 'mensual' cae a 1 mes; semanal/diario/anual
  // también (los valores imposibles de expresar en meses se aplanan a 1
  // para no romper recurrencias activas, pero el usuario debe re-editarlas).
  const periodoBruto = p.periodo_meses != null && p.periodo_meses !== ''
    ? p.periodo_meses
    : 1;
  const periodo = Number(periodoBruto);
  if (!Number.isInteger(periodo) || periodo < 1 || periodo > 12) {
    throw new Error('El periodo debe ser un entero entre 1 y 12 meses');
  }
  p.periodo_meses = periodo;
  const establecimientoId = String(p.establecimiento_id || '').trim();
  if (p.tipo === 'transferencia') {
    p.establecimiento_id = '';
  } else if (establecimientoId) {
    if (!leerHoja('Establecimientos').some(e => e.id === establecimientoId)) {
      throw new Error('Establecimiento no encontrado');
    }
    p.establecimiento_id = establecimientoId;
  } else {
    p.establecimiento_id = '';
  }
  if (p.tipo !== 'transferencia') {
    // ponytail: si trae subcuenta de origen, debe pertenecer a la cuenta.
    // Sin esto, el saldo se imputa a una subcuenta ajena hasta el siguiente
    // bootstrap (cuando normalizarSubcuentasHuerfanas_ la limpia).
    if (p.subcuenta_id) {
      const subO = leerHoja('Cuentas').find(c => c.id === p.subcuenta_id && c.parent_id === p.cuenta_id);
      if (!subO) throw new Error('Subcuenta no encontrada o no pertenece a la cuenta');
    }
    if (p.tipo === 'devolucion') {
      const catsById = {};
      leerHoja('Categorias').forEach(c => { catsById[c.id] = c; });
      validarCategoriaPorTipo_(catsById, p.categoria_id || '', 'gasto', 'Categoría de devolución');
    }
    return p;
  }

  if (!p.cuenta_destino_id) throw new Error('Cuenta destino obligatoria en transferencia recurrente');
  // ponytail: misma cuenta se permite para mover entre subcuentas del mismo
  // padre. La validación fina (subcuentas distintas) la hace guardarTransaccion
  // al materializar la plantilla; aquí solo dejamos pasar la intención.
  if (String(p.cuenta_destino_id) === String(p.cuenta_id)) {
    if (!p.subcuenta_id) throw new Error('Transferencia interna recurrente requiere subcuenta de origen');
    if (!p.subcuenta_destino_id && !(p.reparto_destino && parseRepartoDestino_(p.reparto_destino).length)) {
      throw new Error('Transferencia interna recurrente requiere subcuenta de destino');
    }
  }

  const cuentas = leerHoja('Cuentas');
  if (p.subcuenta_destino_id) {
    const sub = cuentas.find(c => c.id === p.subcuenta_destino_id && c.parent_id === p.cuenta_destino_id);
    if (!sub) throw new Error('Subcuenta destino no encontrada o no pertenece a la cuenta destino');
  }

  const reparto = parseRepartoDestino_(p.reparto_destino);
  if (reparto.length) {
    const subValidas = new Set(cuentas.filter(c => c.parent_id === p.cuenta_destino_id).map(c => c.id));
    let suma = 0;
    reparto.forEach(r => {
      if (!subValidas.has(r.subcuenta_id)) throw new Error('Subcuenta destino no pertenece a la cuenta destino');
      if (!(Number(r.importe) > 0)) throw new Error('Importe de destino debe ser > 0');
      suma += Number(r.importe || 0);
    });
    const totalEsperado = p.importe_destino ? Number(p.importe_destino) : Number(p.importe || 0);
    if (Math.abs(suma - totalEsperado) > 0.01) {
      throw new Error('La suma del reparto no coincide con el importe destino');
    }
  }

  const cuentasById = {};
  cuentas.forEach(c => { cuentasById[c.id] = c; });
  const tipoPresupuesto = tipoTransferenciaPresupuesto_(cuentasById, p.cuenta_id, p.cuenta_destino_id);
  validarCategoriasTransferencia_(tipoPresupuesto, p.categoria_id || '', reparto);

  if (p.ratio_conversion && !(Number(p.importe_destino) > 0)) {
    throw new Error('Falta importe destino o ratio de conversión');
  }
  return p;
}

function guardarRecurrente(rec) {
  const owner = username_();
  if (!rec.plantilla) throw new Error('Falta plantilla');
  const plantilla = typeof rec.plantilla === 'string' ? JSON.parse(rec.plantilla) : Object.assign({}, rec.plantilla);
  validarPlantillaRecurrente_(owner, plantilla);
  const datos = leerHoja('Recurrentes');
  const existente = rec.id ? datos.find(r => r.id === rec.id) : null;
  // ponytail: mes_inicio/anio_inicio se aceptan como columnas propias o como
  // campos dentro de la plantilla; lo que llegue escrito gana. Si no llegan
  // ni en payload ni en la plantilla, conservan lo que ya tenía la fila.
  let mes = normalizarMesInicio_(rec.mes_inicio != null ? rec.mes_inicio : (plantilla.mes_inicio != null ? plantilla.mes_inicio : null));
  let anio = normalizarAnioInicio_(rec.anio_inicio != null ? rec.anio_inicio : (plantilla.anio_inicio != null ? plantilla.anio_inicio : null));
  if ((mes == null || anio == null) && existente) {
    if (mes == null) mes = existente.mes_inicio != null && existente.mes_inicio !== '' ? Number(existente.mes_inicio) : null;
    if (anio == null) anio = existente.anio_inicio != null && existente.anio_inicio !== '' ? Number(existente.anio_inicio) : null;
  }
  if ((mes == null) !== (anio == null)) {
    throw new Error('mes_inicio y anio_inicio deben ir los dos o ninguno');
  }
  const fila = {
    owner: owner,
    id: rec.id || uid_('rec'),
    plantilla: JSON.stringify(plantilla),
    ultima_generacion: rec.ultima_generacion || (existente && existente.ultima_generacion) || '',
    activa: rec.activa !== false,
    mes_inicio: mes == null ? '' : mes,
    anio_inicio: anio == null ? '' : anio
  };
  upsertFila('Recurrentes', fila);
  return obtenerRecurrentes();
}

function normalizarMesInicio_(v) {
  if (v == null || v === '') return null;
  const n = Number(v);
  if (!Number.isInteger(n) || n < 1 || n > 12) throw new Error('mes_inicio debe ser un entero entre 1 y 12');
  return n;
}

function normalizarAnioInicio_(v) {
  if (v == null || v === '') return null;
  const n = Number(v);
  if (!Number.isInteger(n) || n < 1900 || n > 3000) throw new Error('anio_inicio debe ser un año válido');
  return n;
}

function eliminarRecurrente(id) {
  eliminarFila('Recurrentes', id);
  return obtenerRecurrentes();
}

function upsertRecurrenteBase_(owner, tx) {
  const fechaIso = iso_(tx.fecha);
  const fechaObj = parseFecha(fechaIso) || new Date();
  const plantilla = {
    tipo: tx.tipo, importe: Number(tx.importe), cuenta_id: tx.cuenta_id,
    subcuenta_id: tx.subcuenta_id || '',
    cuenta_destino_id: tx.cuenta_destino_id || '',
    subcuenta_destino_id: tx.subcuenta_destino_id || '',
    importe_destino: tx.importe_destino ? Number(tx.importe_destino) : '',
    ratio_conversion: tx.ratio_conversion ? Number(tx.ratio_conversion) : '',
    reparto_destino: Array.isArray(tx.reparto_destino) ? JSON.stringify(tx.reparto_destino) : (tx.reparto_destino || ''),
    categoria_id: tx.categoria_id || '', establecimiento_id: tx.establecimiento_id || '', descripcion: tx.descripcion || '',
    periodo_meses: Number(tx.recurrente_periodo_meses || 1), dia_mes: tx.recurrente_dia || Number(String(tx.fecha).slice(8, 10)),
    inicio: fechaIso, fin: tx.recurrente_fin || ''
  };
  const id = uid_('rec');
  const fila = {
    owner: owner, id,
    plantilla: JSON.stringify(plantilla), ultima_generacion: fechaIso, activa: true,
    mes_inicio: fechaObj.getMonth() + 1, anio_inicio: fechaObj.getFullYear()
  };
  upsertFila('Recurrentes', fila);
  return id;
}

function generarRecurrentesPendientes_(owner, fechaCorte) {
  const actor = currentUser_() || owner;
  const recs = leerHoja('Recurrentes').filter(r => r.owner === owner && r.activa);
  const txs = leerHoja('Transacciones');
  let cambios = false;
  recs.forEach(r => {
    try {
      const p = JSON.parse(r.plantilla);
      const periodo = Math.max(1, Number(p.periodo_meses) || 1);
      const dia = Math.min(31, Math.max(1, Number(p.dia_mes) || 1));
      const inicio = parseFecha(p.inicio) || new Date();
      // ponytail: mes_inicio/anio_inicio definen desde qué mes/año se generan
      // transacciones. Si están, ganan sobre el mes de inicio legacy; el día
      // sigue siendo el de dia_mes. Vacío = comportamiento heredado.
      const mesIniRaw = r.mes_inicio != null && r.mes_inicio !== '' ? Number(r.mes_inicio) : null;
      const anioIniRaw = r.anio_inicio != null && r.anio_inicio !== '' ? Number(r.anio_inicio) : null;
      let base;
      if (r.ultima_generacion) {
        const u = parseFecha(r.ultima_generacion);
        base = u && !isNaN(u) ? new Date(u) : new Date(inicio);
      } else {
        base = new Date(inicio);
      }
      if (mesIniRaw && anioIniRaw) {
        const minY = anioIniRaw * 12 + (mesIniRaw - 1);
        const curY = base.getFullYear() * 12 + base.getMonth();
        if (curY < minY) base = new Date(anioIniRaw, mesIniRaw - 1, 1);
      }
      // Primera ocurrencia del día en el mes de base (o último día del mes si no existe).
      let cursor = fechaConDiaMes_(base.getFullYear(), base.getMonth(), dia);
      // Si esa fecha ya quedó cubierta por ultima_generacion, avanzar periodos
      // hasta pasar de ella (evita regenerar el mismo mes).
      if (r.ultima_generacion) {
        const uIso = iso_(base);
        while (iso_(cursor) <= uIso && cursor <= fechaCorte) {
          cursor = siguienteCursor_(cursor, periodo, dia);
        }
      }
      // Normalizar corte a medianoche local para comparación de día fiable.
      const corte = new Date(fechaCorte.getFullYear(), fechaCorte.getMonth(), fechaCorte.getDate());
      while (cursor <= corte) {
        const isoCursor = iso_(cursor);
        // ponytail: idempotencia robusta. Match por (recurrente_id, mes) o
        // por (mes, cuenta, importe, descripcion) para no duplicar el
        // "original" manual previo sin recurrente_id.
        const ya = txConflictaEnMesRecurrente_(txs, r, p, isoCursor);
        if (!ya) {
          txs.push({
            owner: owner, id: uid_('tx'),
            fecha: isoCursor, tipo: p.tipo, importe: Number(p.importe), moneda: 'EUR',
            cuenta_id: p.cuenta_id,
            subcuenta_id: p.subcuenta_id || '',
            cuenta_destino_id: p.cuenta_destino_id || '',
            subcuenta_destino_id: p.subcuenta_destino_id || '',
            importe_destino: p.importe_destino || '',
            ratio_conversion: p.ratio_conversion || '',
            reparto_destino: p.reparto_destino || '',
            categoria_id: p.categoria_id || '',
            establecimiento_id: p.tipo === 'transferencia' ? '' : (p.establecimiento_id || ''),
            descripcion: p.descripcion || '',
            estado: 'pendiente', recurrente_id: r.id, fecha_pago: '', conciliada_con: '', notas: '',
            fecha_creacion: isoAhora_(),
            ultima_edicion_por: actor,
            fecha_ultima_edicion: isoAhora_()
          });
          cambios = true;
        }
        cursor = siguienteCursor_(cursor, periodo, dia);
      }
      r.ultima_generacion = iso_(corte);
    } catch (e) {
      // ponytail: plantilla corrupta no debe romper el bootstrap
      Logger.log('plantilla corrupta ' + r.id + ': ' + e.message);
    }
  });
  if (cambios) escribirHoja('Transacciones', txs);
  escribirHoja('Recurrentes', leerHoja('Recurrentes').map(r => {
    if (r.owner !== owner) return r;
    const nuevo = recs.find(x => x.id === r.id);
    return nuevo || r;
  }));
}

// ponytail: dedup al generar recurrentes. Idempotente por (recurrente_id, mes)
// y por (mes, cuenta, importe, descripcion) para atrapar el "original" manual
// que existía antes de crear el recurrente y no tiene recurrente_id asignado.
function txConflictaEnMesRecurrente_(txs, r, p, isoCursor) {
  const ym = isoCursor.slice(0, 7);
  const desc = String(p.descripcion || '').trim();
  const imp = Number(p.importe);
  const cta = p.cuenta_id || '';
  return txs.some(t => {
    if (typeof t.fecha !== 'string' || t.fecha.slice(0, 7) !== ym) return false;
    if (t.recurrente_id === r.id) return true;
    return t.recurrente_id === '' &&
      t.cuenta_id === cta &&
      Number(t.importe) === imp &&
      String(t.descripcion || '').trim() === desc;
  });
}

// ponytail: construye una fecha en año/mes con el día pedido. Si el mes no
// tiene ese día (31 en feb, 30 en feb, etc.), usa el último día del mes.
// Así "día 31" se interpreta como "último día del mes".
function fechaConDiaMes_(year, month, dia) {
  const last = new Date(year, month + 1, 0).getDate();
  const d = Math.min(Math.max(1, Number(dia) || 1), last);
  return new Date(year, month, d);
}

// Avanza N meses manteniendo el día de la plantilla (o el último día del mes
// si ese día no existe). Evita el desborde clásico de setMonth con días 29-31.
function siguienteCursor_(d, periodoMeses, dia) {
  const meses = Number(periodoMeses) || 1;
  const targetMonth = d.getMonth() + meses;
  const year = d.getFullYear() + Math.floor(targetMonth / 12);
  const month = ((targetMonth % 12) + 12) % 12;
  return fechaConDiaMes_(year, month, dia != null ? dia : d.getDate());
}

function generarRecurrentesPendientes(fechaCorte) {
  generarRecurrentesPendientes_(username_(), parseFecha(fechaCorte) || new Date());
  return obtenerTransacciones({});
}

// ───────── Presupuestos ─────────
// ponytail: one budget per category, same every month
function obtenerPresupuestos() {
  const all = filasVisibles_('Presupuestos');
  const seen = new Map();
  all.forEach(p => seen.set(p.categoria_id, p));
  return [...seen.values()];
}

function guardarPresupuesto(p) {
  const owner = username_();
  if (!p.categoria_id) throw new Error('Categoría requerida');
  const fila = {
    owner: owner, id: p.id || uid_('ppto'),
    categoria_id: p.categoria_id, importe_esperado: Number(p.importe_esperado || 0)
  };
  upsertFila('Presupuestos', fila);
  return obtenerPresupuestos();
}

function eliminarPresupuesto(id) {
  eliminarFila('Presupuestos', id);
  return obtenerPresupuestos();
}

// ───────── Conciliación ─────────
function conciliar(cuenta_id, saldo_banco, fecha) {
  const owner = username_();
  const cuenta = leerHoja('Cuentas').find(c => c.id === cuenta_id);
  if (!cuenta) throw new Error('Cuenta no encontrada');
  if (cuenta.parent_id) throw new Error('Solo se pueden conciliar cuentas, no subcuentas');
  // Saldo canónico (mismo cálculo que la UI).
  const cta = obtenerCuentas().find(c => c.id === cuenta_id);
  const sistema = cta ? cta.saldo : 0;
  const banco = Number(saldo_banco);
  const diferencia = +(banco - sistema).toFixed(2);
  const txs = leerHoja('Transacciones').filter(t => (t.cuenta_id === cuenta_id || t.cuenta_destino_id === cuenta_id));
  const nuevasTxs = txs.map(t => Object.assign({}, t, { estado: 'conciliado', fecha_pago: fecha || isoHoy_() }));
  const todas = leerHoja('Transacciones').map(t => {
    const nueva = nuevasTxs.find(n => n.id === t.id);
    return nueva || t;
  });
  escribirHoja('Transacciones', todas);
  const eventos = leerHoja('Conciliaciones');
  eventos.push({
    owner: owner, id: uid_('con'),
    fecha: fecha || isoHoy_(), cuenta_id,
    saldo_sistema: sistema, saldo_banco: banco, diferencia, notas: ''
  });
  escribirHoja('Conciliaciones', eventos);
  return { sistema, banco, diferencia, fecha: fecha || isoHoy_() };
}

function obtenerConciliaciones(opts) {
  opts = opts || {};
  const limit = Number(opts.limit) || 0;
  const offset = Number(opts.offset) || 0;
  let rows = leerHoja('Conciliaciones');
  rows.sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
  const total = rows.length;
  if (limit > 0) rows = rows.slice(offset, offset + limit);
  return { total, rows };
}

function editarConciliacion(id, cambios) {
  const owner = username_();
  const eventos = leerHoja('Conciliaciones');
  const idx = eventos.findIndex(c => c.id === id && c.owner === owner);
  if (idx === -1) throw new Error('Conciliación no encontrada');
  const actual = eventos[idx];
  const banco = cambios.saldo_banco !== undefined ? Number(cambios.saldo_banco) : Number(actual.saldo_banco);
  const sistema = Number(actual.saldo_sistema);
  const diferencia = +(banco - sistema).toFixed(2);
  eventos[idx] = Object.assign({}, actual, {
    saldo_banco: banco,
    diferencia,
    notas: cambios.notas !== undefined ? String(cambios.notas || '') : actual.notas
  });
  escribirHoja('Conciliaciones', eventos);
  return eventos[idx];
}

function eliminarConciliacion(id) {
  const owner = username_();
  const eventos = leerHoja('Conciliaciones');
  const idx = eventos.findIndex(c => c.id === id && c.owner === owner);
  if (idx === -1) throw new Error('Conciliación no encontrada');
  // ponytail: no revertimos estado=conciliado en tx porque no sabemos qué eventos marcaron cuáles.
  eventos.splice(idx, 1);
  escribirHoja('Conciliaciones', eventos);
  return { ok: true };
}

// ───────── Resumen y evolución ─────────
function obtenerResumen(anio, mes) {
  const a = anio || new Date().getFullYear();
  const m = mes || (new Date().getMonth() + 1);
  const txs = leerHoja('Transacciones');
  const cuentasById = {};
  filasVisibles_('Cuentas')
    .forEach(c => { cuentasById[c.id] = c; });
  const enMes = txs.filter(t => {
    const f = new Date(t.fecha);
    return f.getFullYear() === Number(a) && (f.getMonth() + 1) === Number(m);
  });
  const ingresos = enMes.filter(t => t.tipo === 'ingreso').reduce((s, t) => s + Number(t.importe || 0), 0);
  const gastos = enMes.filter(t => t.tipo === 'gasto').reduce((s, t) => s + Number(t.importe || 0), 0);
  // ponytail: devoluciones restan del total de gastos del mes (revierte gastos).
  const devoluciones = enMes.filter(t => t.tipo === 'devolucion').reduce((s, t) => s + Number(t.importe || 0), 0);
  const transferenciasGasto = enMes
    .filter(t => t.tipo === 'transferencia' && tipoTransferenciaPresupuestoTx_(t, cuentasById) === 'gasto')
    .reduce((s, t) => s + Number(t.importe || 0), 0);
  const transferenciasIngreso = enMes
    .filter(t => t.tipo === 'transferencia' && tipoTransferenciaPresupuestoTx_(t, cuentasById) === 'ingreso')
    .reduce((s, t) => s + Number(t.importe || 0), 0);
  const pendiente = enMes.filter(t => t.estado === 'pendiente').length;
  const vencido = enMes.filter(t => t.estado === 'vencido').length;
  // Evolución últimos 12 meses
  const evol = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(a, m - 1 - i, 1);
    // ponytail: usa componentes locales (script tz) — formatDate con ss_tz podía devolver
    // el mes anterior si la hoja estaba en UTC; los slice(0,7) de txs almacenados no matcheaban
    const k = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
    const en = txs.filter(t => String(t.fecha).slice(0, 7) === k);
    evol.push({
      mes: k,
      ingresos: en.filter(t => t.tipo === 'ingreso').reduce((s, t) => s + Number(t.importe || 0), 0),
      gastos: en.filter(t => t.tipo === 'gasto').reduce((s, t) => s + Number(t.importe || 0), 0)
              - en.filter(t => t.tipo === 'devolucion').reduce((s, t) => s + Number(t.importe || 0), 0)
    });
  }
  // Próximos recurrentes
  const recs = filasVisibles_('Recurrentes').filter(r => r.activa);
  const proximos = recs.map(r => {
    try {
      const p = JSON.parse(r.plantilla);
      const periodo = Math.max(1, Number(p.periodo_meses) || 1);
      const dia = Math.min(31, Math.max(1, Number(p.dia_mes) || 1));
      const inicio = parseFecha(p.inicio) || new Date();
      const mesIniRaw = r.mes_inicio != null && r.mes_inicio !== '' ? Number(r.mes_inicio) : null;
      const anioIniRaw = r.anio_inicio != null && r.anio_inicio !== '' ? Number(r.anio_inicio) : null;
      let base;
      if (r.ultima_generacion) {
        const u = parseFecha(r.ultima_generacion);
        base = u && !isNaN(u) ? new Date(u) : new Date(inicio);
      } else {
        base = new Date(inicio);
      }
      if (mesIniRaw && anioIniRaw) {
        const minY = anioIniRaw * 12 + (mesIniRaw - 1);
        const curY = base.getFullYear() * 12 + base.getMonth();
        if (curY < minY) base = new Date(anioIniRaw, mesIniRaw - 1, 1);
      }
      let cursor = fechaConDiaMes_(base.getFullYear(), base.getMonth(), dia);
      // ponytail: avanza el cursor hasta el primer disparo estrictamente posterior a hoy
      const hoy = new Date();
      const hoyMid = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
      while (cursor <= hoyMid) cursor = siguienteCursor_(cursor, periodo, dia);
      return {
        id: r.id, descripcion: p.descripcion, importe: Number(p.importe),
        dia_mes: p.dia_mes, periodo_meses: p.periodo_meses || 1,
        proximo_disparo: iso_(cursor)
      };
    } catch (e) {
      return null;
    }
  }).filter(Boolean);
  return {
    anio: a,
    mes: m,
    ingresos,
    ingresosPresupuesto: ingresos + transferenciasIngreso,
    gastos: gastos - devoluciones,
    gastosPresupuesto: (gastos - devoluciones) + transferenciasGasto,
    neto: ingresos - (gastos - devoluciones),
    pendiente,
    vencido,
    evol,
    proximos
  };
}

function obtenerResumenEstablecimientos(anio, mes) {
  const a = Number(anio || new Date().getFullYear());
  const m = Number(mes || (new Date().getMonth() + 1));
  const filas = obtenerEstablecimientos().map(e => ({
    id: e.id,
    nombre: e.nombre,
    ingresos: 0,
    gastos: 0,
    neto: 0
  }));
  filas.push({ id: '', nombre: 'Sin establecimiento', ingresos: 0, gastos: 0, neto: 0 });
  const porId = {};
  filas.forEach(f => { porId[f.id] = f; });
  const periodo = a + '-' + String(m).padStart(2, '0');

  leerHoja('Transacciones')
    .filter(t => String(t.fecha).slice(0, 7) === periodo && (t.tipo === 'gasto' || t.tipo === 'ingreso' || t.tipo === 'devolucion'))
    .forEach(t => {
      const fila = porId[String(t.establecimiento_id || '')] || porId[''];
      if (t.tipo === 'ingreso') fila.ingresos += Number(t.importe || 0);
      else if (t.tipo === 'gasto') fila.gastos += Number(t.importe || 0);
      else if (t.tipo === 'devolucion') fila.gastos -= Number(t.importe || 0); // ponytail: revierten gasto
    });
  filas.forEach(f => { f.neto = f.ingresos - f.gastos; });
  return { anio: a, mes: m, filas: filas };
}

function obtenerCategoriasResumen(anio, mes) {
  const a = anio || new Date().getFullYear();
  const m = mes || (new Date().getMonth() + 1);
  const txs = leerHoja('Transacciones').filter(t => ['gasto', 'ingreso', 'transferencia', 'devolucion'].includes(t.tipo));
  const ps = filasVisibles_('Presupuestos');
  const cats = obtenerCategorias();
  const catsById = {};
  cats.forEach(c => { catsById[c.id] = c; });
  const cuentasById = {};
  filasVisibles_('Cuentas')
    .forEach(c => { cuentasById[c.id] = c; });
  // Reparto destino de transferencias: cada subcuenta se imputa a su propia
  // categoría (gasto o ingreso según la dirección activo/pasivo).
  // No se usa categoria_id del tx como fallback para evitar el campo separado.
  const porCatRep = {};
  txs.forEach(t => {
    if (t.tipo !== 'transferencia') return;
    const f = new Date(t.fecha);
    if (f.getFullYear() !== Number(a) || (f.getMonth() + 1) !== Number(m)) return;
    const tipoPresupuesto = tipoTransferenciaPresupuestoTx_(t, cuentasById);
    if (tipoPresupuesto === 'neutro') return;
    const rep = parseRepartoDestino_(t.reparto_destino);
    if (rep.length) {
      rep.forEach(r => {
        const imp = Number(r.importe || 0);
        if (!(imp > 0)) return;
        const catId = r.categoria_id || '';
        const cat = catsById[catId];
        if (!cat || cat.tipo !== tipoPresupuesto) return;
        porCatRep[catId] = (porCatRep[catId] || 0) + imp;
      });
    }
  });
  return cats.map(c => {
    const esperado = ps.filter(p => p.categoria_id === c.id).reduce((s, p) => s + Number(p.importe_esperado || 0), 0);
    // ponytail: devoluciones (importe positivo) restan del real de su categoría
    // de gasto. Se cuentan junto al caso !transferencia para reutilizar el
    // filtro por categoria_id y mes.
    const real = txs.filter(t => {
      if (t.tipo === 'transferencia') return false; // transferencias se cuentan por porCatRep
      if (t.categoria_id !== c.id) return false;
      const f = new Date(t.fecha);
      return f.getFullYear() === Number(a) && (f.getMonth() + 1) === Number(m);
    }).reduce((s, t) => {
      const imp = Number(t.importe || 0);
      return s + (t.tipo === 'devolucion' ? -imp : imp);
    }, 0) + (porCatRep[c.id] || 0);
    return { id: c.id, nombre: c.nombre, color: c.color, esperado, real, diferencia: real - esperado };
  });
}

function guardarTipoCambio(base, destino, ratio) {
  const owner = username_();
  if (!(ratio > 0)) throw new Error('Ratio debe ser > 0');
  const fila = {
    owner: owner, id: uid_('tc'),
    fecha: isoHoy_(), base: String(base).toUpperCase(), destino: String(destino).toUpperCase(),
    ratio: Number(ratio)
  };
  upsertFila('TiposCambio', fila);
  return fila;
}

// ───────── Self-test mínimo ─────────
function __selfTest() {
  _selfTestActive = true;
  const owner = SELF_TEST_USER;
  try {
    return __selfTestBody_(owner);
  } finally {
    _selfTestActive = false;
  }
}

function __selfTestBody_(owner) {
  sembrar(owner);
  const cuentas = obtenerCuentas();
  const cat = obtenerCategorias();
  if (!cuentas.length) throw new Error('Sin cuentas');
  if (!cat.length) throw new Error('Sin categorías');

  const crossOwnerCatId = uid_('cross-cat');
  upsertFila('Categorias', {
    owner: '__otro_owner__', id: crossOwnerCatId, nombre: 'self-cross-owner',
    color: '#000000', icono: 'category', tipo: 'gasto', orden: 999
  });
  guardarCategoria({ id: crossOwnerCatId, nombre: 'self-cross-owner-edit' });
  const crossOwnerCat = leerHoja('Categorias').find(c => c.id === crossOwnerCatId);
  if (!crossOwnerCat || crossOwnerCat.nombre !== 'self-cross-owner-edit' || crossOwnerCat.owner !== '__otro_owner__') {
    throw new Error('No permitió editar categoría de otro owner');
  }
  eliminarCategoria(crossOwnerCatId);
  if (leerHoja('Categorias').some(c => c.id === crossOwnerCatId)) {
    throw new Error('No permitió eliminar categoría de otro owner');
  }

  const t = guardarTransaccion({
    tipo: 'gasto', importe: 12.5, cuenta_id: cuentas[0].id, categoria_id: cat[0].id,
    descripcion: 'self-test', fecha: isoHoy_()
  });
  if (!t.id) throw new Error('Fallo al guardar transacción');
  eliminarTransaccion(t.id);

  const estNombre = 'self-est-' + Utilities.getUuid().slice(0, 8);
  const establecimiento = guardarEstablecimiento({ nombre: estNombre }).find(e => e.nombre === estNombre);
  if (!establecimiento) throw new Error('Fallo al crear establecimiento');
  const txEstGasto = guardarTransaccion({
    tipo: 'gasto', importe: 12.5, cuenta_id: cuentas[0].id, categoria_id: cat[0].id,
    establecimiento_id: establecimiento.id, descripcion: 'self-est-gasto', fecha: isoHoy_()
  });
  const txEstIngreso = guardarTransaccion({
    tipo: 'ingreso', importe: 20, cuenta_id: cuentas[0].id,
    establecimiento_id: establecimiento.id, descripcion: 'self-est-ingreso', fecha: isoHoy_()
  });
  const resumenEst = obtenerResumenEstablecimientos().filas.find(e => e.id === establecimiento.id);
  if (!resumenEst || resumenEst.gastos !== 12.5 || resumenEst.ingresos !== 20 || resumenEst.neto !== 7.5) {
    throw new Error('Resumen de establecimiento incorrecto: ' + JSON.stringify(resumenEst));
  }
  const resumenSinEstAntes = obtenerResumenEstablecimientos().filas.find(e => e.id === '') || { ingresos: 0, gastos: 0, neto: 0 };
  const txSinEstGasto = guardarTransaccion({
    tipo: 'gasto', importe: 3, cuenta_id: cuentas[0].id, categoria_id: cat[0].id,
    descripcion: 'self-est-sin-gasto', fecha: isoHoy_()
  });
  const txSinEstIngreso = guardarTransaccion({
    tipo: 'ingreso', importe: 5, cuenta_id: cuentas[0].id,
    descripcion: 'self-est-sin-ingreso', fecha: isoHoy_()
  });
  const resumenSinEst = obtenerResumenEstablecimientos().filas.find(e => e.id === '');
  if (!resumenSinEst) {
    throw new Error('Falta la fila Sin establecimiento');
  }
  if (resumenSinEst.gastos !== resumenSinEstAntes.gastos + 3 || resumenSinEst.ingresos !== resumenSinEstAntes.ingresos + 5 || resumenSinEst.neto !== resumenSinEstAntes.neto + 2) {
    throw new Error('Resumen sin establecimiento incorrecto: ' + JSON.stringify(resumenSinEst));
  }
  let estErr = null;
  try { eliminarEstablecimiento(establecimiento.id); }
  catch (e) { estErr = e.message; }
  if (!estErr || !/movimientos asociados/.test(estErr)) throw new Error('No bloqueó establecimiento con movimientos: ' + estErr);

  const destinoEstNombre = 'self-est-dest-' + Utilities.getUuid().slice(0, 8);
  const destinoEst = guardarCuenta({
    nombre: destinoEstNombre, tipo: cuentas[0].tipo, moneda: cuentas[0].moneda || 'EUR', saldo_inicial: 0
  }).find(c => c.nombre === destinoEstNombre);
  const txEstTransfer = guardarTransaccion({
    tipo: 'transferencia', importe: 1, cuenta_id: cuentas[0].id, cuenta_destino_id: destinoEst.id,
    establecimiento_id: establecimiento.id, descripcion: 'self-est-transfer', fecha: isoHoy_()
  });
  if (txEstTransfer.establecimiento_id !== '') throw new Error('Transferencia conservó establecimiento');
  eliminarTransaccion(txEstGasto.id);
  eliminarTransaccion(txEstIngreso.id);
  eliminarTransaccion(txSinEstGasto.id);
  eliminarTransaccion(txSinEstIngreso.id);
  eliminarTransaccion(txEstTransfer.id);
  eliminarCuenta(destinoEst.id);

  const recEstId = uid_('rec');
  const recEstDesc = 'self-est-rec-' + Utilities.getUuid().slice(0, 8);
  guardarRecurrente({
    id: recEstId,
    plantilla: {
      tipo: 'gasto', importe: 3.21, cuenta_id: cuentas[0].id, categoria_id: cat[0].id,
      establecimiento_id: establecimiento.id, descripcion: recEstDesc,
      periodo_meses: 1, dia_mes: Number(isoHoy_().slice(8, 10)), inicio: isoHoy_()
    },
    ultima_generacion: '', activa: true
  });
  const recEst = leerHoja('Recurrentes').find(r => r.owner === owner && r.id === recEstId);
  if (!recEst || JSON.parse(recEst.plantilla).establecimiento_id !== establecimiento.id) {
    throw new Error('Plantilla recurrente perdió establecimiento');
  }
  estErr = null;
  try { eliminarEstablecimiento(establecimiento.id); }
  catch (e) { estErr = e.message; }
  if (!estErr || !/plantillas recurrentes/.test(estErr)) throw new Error('No bloqueó establecimiento recurrente: ' + estErr);
  generarRecurrentesPendientes_(owner, new Date());
  const txEstGenerada = leerHoja('Transacciones').find(x => x.owner === owner && x.recurrente_id === recEstId);
  if (!txEstGenerada || txEstGenerada.establecimiento_id !== establecimiento.id) {
    throw new Error('Transacción recurrente perdió establecimiento');
  }
  eliminarTransaccion(txEstGenerada.id);
  eliminarRecurrente(recEstId);
  eliminarEstablecimiento(establecimiento.id);

  const conciliado = conciliar(cuentas[0].id, cuentas[0].saldo || 0, isoHoy_());
  if (typeof conciliado.diferencia !== 'number') throw new Error('Fallo en conciliación');

  // ponytail: smoke-test subcuentas + reorder + tx con subcuenta_id
  const parent = cuentas[0];
  const subName = 'self-sub-' + Utilities.getUuid().slice(0, 8);
  const withSub = guardarCuenta({
    parent_id: parent.id, nombre: subName, saldo_inicial: 100
  });
  const subId = withSub.find(c => c.id === parent.id).subcuentas.find(s => s.nombre === subName).id;
  const txSub = guardarTransaccion({
    tipo: 'gasto', importe: 30, cuenta_id: parent.id, subcuenta_id: subId,
    categoria_id: cat[0].id, descripcion: 'self-sub-tx', fecha: isoHoy_()
  });
  const fresh = obtenerCuentas().find(c => c.id === parent.id);
  const sub = (fresh.subcuentas || []).find(s => s.id === subId);
  if (!sub || sub.saldo !== 70) throw new Error('Saldo subcuenta incorrecto: ' + (sub && sub.saldo));
  // Modelo de partición: el saldo_inicial (100) de la subcuenta ya forma parte del
  // saldo del padre, así que crearla no cambia el total; solo el gasto (30) lo baja.
  if (fresh.saldo !== parent.saldo - 30) throw new Error('Saldo padre incorrecto: ' + fresh.saldo);
  // Reorder y segunda alta sin reemplazar la primera
  const anotherName = 'self-sub-' + Utilities.getUuid().slice(0, 8);
  const withAnother = guardarCuenta({
    parent_id: parent.id, nombre: anotherName, saldo_inicial: 0
  });
  const parentWithBoth = withAnother.find(c => c.id === parent.id);
  const anotherId = parentWithBoth.subcuentas.find(s => s.nombre === anotherName).id;
  if (!parentWithBoth.subcuentas.some(s => s.id === subId)) throw new Error('Crear una subcuenta reemplazó la anterior');
  // Creamos una cuenta destino pasiva para validar el caso activo -> pasivo
  // (debe computar como gasto en presupuesto).
  const destParent = guardarCuenta({
    nombre: 'self-dest-' + Utilities.getUuid().slice(0, 8),
    tipo: 'pasivo', moneda: 'EUR', saldo_inicial: 0
  }).find(c => c.nombre.startsWith('self-dest-'));
  const destSubName = 'self-dest-sub-' + Utilities.getUuid().slice(0, 8);
  const withDestSub = guardarCuenta({ parent_id: destParent.id, nombre: destSubName, saldo_inicial: 0 });
  const destSubId = withDestSub.find(c => c.id === destParent.id).subcuentas.find(s => s.nombre === destSubName).id;
  const realAntesTransferencia = obtenerCategoriasResumen().find(c => c.id === cat[0].id).real;
  const gastoPresupuestoAntes = obtenerResumen().gastosPresupuesto;
  const txTransfer = guardarTransaccion({
    tipo: 'transferencia', importe: 20,
    cuenta_id: parent.id, subcuenta_id: subId,
    cuenta_destino_id: destParent.id, subcuenta_destino_id: destSubId,
    categoria_id: cat[0].id, descripcion: 'self-sub-transfer', fecha: isoHoy_()
  });
  const realDespuesTransferencia = obtenerCategoriasResumen().find(c => c.id === cat[0].id).real;
  if (realDespuesTransferencia !== realAntesTransferencia + 20) throw new Error('Transferencia no sumó al presupuesto: ' + realDespuesTransferencia);
  const gastoPresupuestoDespues = obtenerResumen().gastosPresupuesto;
  if (gastoPresupuestoDespues !== gastoPresupuestoAntes + 20) throw new Error('Transferencia no sumó al gasto actual: ' + gastoPresupuestoDespues);
  const targetAfter = obtenerCuentas().find(c => c.id === destParent.id).subcuentas.find(s => s.id === destSubId);
  if (targetAfter.saldo !== 20) throw new Error('Transferencia no acreditó subcuenta destino: ' + targetAfter.saldo);

  // Transferencia interna entre subcuentas del mismo padre: no debe contar en
  // presupuesto y debe mover saldo entre subcuentas sin tocar el total del padre.
  const parentPre = obtenerCuentas().find(c => c.id === parent.id);
  const subPre = parentPre.subcuentas.find(s => s.id === subId);
  const otherPre = parentPre.subcuentas.find(s => s.id === anotherId);
  const ingresoPresIntAntes = obtenerResumen().ingresosPresupuesto;
  const gastoPresIntAntes = obtenerResumen().gastosPresupuesto;
  const txInternal = guardarTransaccion({
    tipo: 'transferencia', importe: 8,
    cuenta_id: parent.id, subcuenta_id: subId,
    cuenta_destino_id: parent.id, subcuenta_destino_id: anotherId,
    descripcion: 'self-internal-transfer', fecha: isoHoy_()
  });
  const parentPost = obtenerCuentas().find(c => c.id === parent.id);
  const subPost = parentPost.subcuentas.find(s => s.id === subId);
  const otherPost = parentPost.subcuentas.find(s => s.id === anotherId);
  if (subPost.saldo !== subPre.saldo - 8) throw new Error('Interna no debitó sub origen: ' + subPost.saldo + ' vs ' + (subPre.saldo - 8));
  if (otherPost.saldo !== otherPre.saldo + 8) throw new Error('Interna no acreditó sub destino: ' + otherPost.saldo + ' vs ' + (otherPre.saldo + 8));
  if (Math.abs(parentPost.saldo - parentPre.saldo) > 0.01) throw new Error('Interna alteró saldo del padre: ' + parentPost.saldo + ' vs ' + parentPre.saldo);
  if (obtenerResumen().ingresosPresupuesto !== ingresoPresIntAntes) throw new Error('Interna alteró ingresos presupuesto');
  if (obtenerResumen().gastosPresupuesto !== gastoPresIntAntes) throw new Error('Interna alteró gastos presupuesto');
  // Misma cuenta sin subcuenta origen o destino debe rechazarse.
  let intErr;
  intErr = null;
  try { guardarTransaccion({ tipo: 'transferencia', importe: 1, cuenta_id: parent.id, cuenta_destino_id: parent.id, descripcion: 'self-internal-no-sub', fecha: isoHoy_() }); }
  catch (e) { intErr = e.message; }
  if (!intErr || !/subcuenta de origen/.test(intErr)) throw new Error('No rechazó interna sin sub origen: ' + intErr);
  intErr = null;
  try { guardarTransaccion({ tipo: 'transferencia', importe: 1, cuenta_id: parent.id, subcuenta_id: subId, cuenta_destino_id: parent.id, descripcion: 'self-internal-self', fecha: isoHoy_() }); }
  catch (e) { intErr = e.message; }
  if (!intErr || !/misma que la de origen/.test(intErr)) throw new Error('No rechazó interna origen==destino: ' + intErr);

  // Transferencia pasivo -> activo: debe computar como ingreso de presupuesto.
  const catIngreso = obtenerCategorias().find(c => c.tipo === 'ingreso');
  if (!catIngreso) throw new Error('No hay categoría de ingreso para self-test');
  const ingresoPresupuestoAntes = obtenerResumen().ingresosPresupuesto;
  const txTransferIngreso = guardarTransaccion({
    tipo: 'transferencia', importe: 15,
    cuenta_id: destParent.id,
    cuenta_destino_id: parent.id,
    categoria_id: catIngreso.id,
    descripcion: 'self-income-transfer', fecha: isoHoy_()
  });
  const ingresoPresupuestoDespues = obtenerResumen().ingresosPresupuesto;
  if (ingresoPresupuestoDespues !== ingresoPresupuestoAntes + 15) {
    throw new Error('Transferencia pasivo→activo no sumó al ingreso presupuesto: ' + ingresoPresupuestoDespues);
  }
  const after = reordenarSubcuentas(parent.id, [anotherId, subId]);
  const parentAfter = after.find(c => c.id === parent.id);
  const testOrder = parentAfter.subcuentas.filter(s => s.id === anotherId || s.id === subId);
  if (testOrder.length !== 2 || testOrder[0].id !== anotherId) throw new Error('Reorder falló');
  if (cat.length >= 2) {
    const ordenOriginal = cat.map(c => c.id);
    const reordenadas = reordenarCategorias([cat[1].id, cat[0].id, ...ordenOriginal.slice(2)]);
    if (reordenadas[0].id !== cat[1].id) throw new Error('Reorder categorías falló');
    reordenarCategorias(ordenOriginal);
  }

  // Transferencia con reparto a varias subcuentas destino: la suma del reparto
  // debe coincidir con el importe total, y cada subcuenta destino recibe su parte.
  const splitSubAName = 'self-split-a-' + Utilities.getUuid().slice(0, 8);
  const splitSubBName = 'self-split-b-' + Utilities.getUuid().slice(0, 8);
  guardarCuenta({ parent_id: destParent.id, nombre: splitSubAName, saldo_inicial: 0 });
  guardarCuenta({ parent_id: destParent.id, nombre: splitSubBName, saldo_inicial: 0 });
  const destCtas = obtenerCuentas().find(c => c.id === destParent.id);
  const splitSubA = destCtas.subcuentas.find(s => s.nombre === splitSubAName).id;
  const splitSubB = destCtas.subcuentas.find(s => s.nombre === splitSubBName).id;
  const txSplit = guardarTransaccion({
    tipo: 'transferencia', importe: 50,
    cuenta_id: parent.id,
    cuenta_destino_id: destParent.id,
    reparto_destino: [
      { subcuenta_id: splitSubA, importe: 30, categoria_id: cat[0].id },
      { subcuenta_id: splitSubB, importe: 20, categoria_id: cat[0].id }
    ],
    descripcion: 'self-split-transfer', fecha: isoHoy_()
  });
  const afterSplit = obtenerCuentas().find(c => c.id === destParent.id).subcuentas;
  const subAAfter = afterSplit.find(s => s.id === splitSubA);
  const subBAfter = afterSplit.find(s => s.id === splitSubB);
  if (!subAAfter || subAAfter.saldo !== 30) throw new Error('Reparto no acreditó subA: ' + (subAAfter && subAAfter.saldo));
  if (!subBAfter || subBAfter.saldo !== 20) throw new Error('Reparto no acreditó subB: ' + (subBAfter && subBAfter.saldo));
  // Validación: la suma debe ser exactamente el total del importe.
  let fallido = null;
  try {
    guardarTransaccion({
      tipo: 'transferencia', importe: 50,
      cuenta_id: parent.id, cuenta_destino_id: destParent.id,
      reparto_destino: [
        { subcuenta_id: splitSubA, importe: 30, categoria_id: cat[0].id },
        { subcuenta_id: splitSubB, importe: 15, categoria_id: cat[0].id }
      ],
      descripcion: 'self-split-bad', fecha: isoHoy_()
    });
  } catch (e) { fallido = e.message; }
  if (!fallido || !/suma del reparto/.test(fallido)) throw new Error('No se rechazó reparto con suma incorrecta: ' + fallido);

  // Reparto destino con categorías por subcuenta: cada subcuenta se imputa a
  // su propia categoría, no en una sola para todo el importe.
  const catAName = 'self-catA-' + Utilities.getUuid().slice(0, 8);
  const catBName = 'self-catB-' + Utilities.getUuid().slice(0, 8);
  const catAId = guardarCategoria({ nombre: catAName, color: '#000000', icono: 'tag', tipo: 'gasto' }).find(c => c.nombre === catAName).id;
  const catBId = guardarCategoria({ nombre: catBName, color: '#000000', icono: 'tag', tipo: 'gasto' }).find(c => c.nombre === catBName).id;
  const realAAAntes = obtenerCategoriasResumen().find(c => c.id === catAId).real;
  const realBAntes = obtenerCategoriasResumen().find(c => c.id === catBId).real;
  const txCatSplit = guardarTransaccion({
    tipo: 'transferencia', importe: 50,
    cuenta_id: parent.id,
    cuenta_destino_id: destParent.id,
    reparto_destino: [
      { subcuenta_id: splitSubA, importe: 30, categoria_id: catAId },
      { subcuenta_id: splitSubB, importe: 20, categoria_id: catBId }
    ],
    descripcion: 'self-cat-split', fecha: isoHoy_()
  });
  const realAADespues = obtenerCategoriasResumen().find(c => c.id === catAId).real;
  const realBDespues = obtenerCategoriasResumen().find(c => c.id === catBId).real;
  if (realAADespues !== realAAAntes + 30) throw new Error('Reparto no atribuyó 30 a catA: ' + realAADespues);
  if (realBDespues !== realBAntes + 20) throw new Error('Reparto no atribuyó 20 a catB: ' + realBDespues);
  eliminarTransaccion(txCatSplit.id);
  eliminarCategoria(catAId);
  eliminarCategoria(catBId);

  // Cleanup
  eliminarTransaccion(txSub.id);
  eliminarTransaccion(txTransfer.id);
  eliminarTransaccion(txTransferIngreso.id);
  eliminarTransaccion(txSplit.id);
  eliminarTransaccion(txInternal.id);
  eliminarCuenta(subId);
  eliminarCuenta(anotherId);
  eliminarCuenta(destSubId);
  eliminarCuenta(splitSubA);
  eliminarCuenta(splitSubB);
  eliminarCuenta(destParent.id);

  // Validación de periodo_meses en plantillas recurrentes (1-12 inclusive,
  // entero; valores fuera de rango o no enteros se rechazan; legacy sin
  // periodo_meses cae a 1).
  const baseRec = { tipo: 'gasto', importe: 1, cuenta_id: cuentas[0].id, categoria_id: cat[0].id, dia_mes: 1, inicio: '2024-01-01' };
  let recErr;
  recErr = null;
  try { validarPlantillaRecurrente_(owner, Object.assign({}, baseRec, { periodo_meses: 13 })); }
  catch (e) { recErr = e.message; }
  if (!recErr || !/entre 1 y 12/.test(recErr)) throw new Error('No rechazó periodo_meses=13: ' + recErr);
  recErr = null;
  try { validarPlantillaRecurrente_(owner, Object.assign({}, baseRec, { periodo_meses: 0 })); }
  catch (e) { recErr = e.message; }
  if (!recErr || !/entre 1 y 12/.test(recErr)) throw new Error('No rechazó periodo_meses=0: ' + recErr);
  recErr = null;
  try { validarPlantillaRecurrente_(owner, Object.assign({}, baseRec, { periodo_meses: 1.5 })); }
  catch (e) { recErr = e.message; }
  if (!recErr || !/entre 1 y 12/.test(recErr)) throw new Error('No rechazó periodo_meses=1.5: ' + recErr);
  const legacyRec = Object.assign({}, baseRec);
  delete legacyRec.periodo_meses;
  legacyRec.periodicidad = 'semanal';
  validarPlantillaRecurrente_(owner, legacyRec);
  if (legacyRec.periodo_meses !== 1) throw new Error('Legacy periodicidad no cayó a 1 mes: ' + legacyRec.periodo_meses);
  const okRec = Object.assign({}, baseRec, { periodo_meses: 6 });
  validarPlantillaRecurrente_(owner, okRec);
  if (okRec.periodo_meses !== 6) throw new Error('periodo_meses=6 fue mutado: ' + okRec.periodo_meses);

  // Validación de subcuenta_id en plantilla: si está, debe pertenecer a la
  // cuenta. Sin esta guarda, las txs generadas imputarían saldo a una
  // subcuenta ajena hasta el siguiente bootstrap.
  const recConSubOk = Object.assign({}, baseRec, { subcuenta_id: subId });
  validarPlantillaRecurrente_(owner, recConSubOk);
  let recSubErr;
  recSubErr = null;
  try { validarPlantillaRecurrente_(owner, Object.assign({}, baseRec, { subcuenta_id: destSubId })); }
  catch (e) { recSubErr = e.message; }
  if (!recSubErr || !/no pertenece a la cuenta/.test(recSubErr)) throw new Error('No rechazó subcuenta ajena: ' + recSubErr);
  recSubErr = null;
  try { validarPlantillaRecurrente_(owner, Object.assign({}, baseRec, { subcuenta_id: 'cta_inexistente' })); }
  catch (e) { recSubErr = e.message; }
  if (!recSubErr || !/Subcuenta no encontrada/.test(recSubErr)) throw new Error('No rechazó subcuenta inexistente: ' + recSubErr);

  // Dedup de recurrentes: detectar el "original" manual sin recurrente_id y
  // no duplicarlo al generar.
  const fakeRec = { id: 'rec_dedup', owner };
  const plantAlquiler = { tipo: 'gasto', importe: 100, cuenta_id: cuentas[0].id,
                          descripcion: 'Alquiler', inicio: '2026-08-15', dia_mes: 15, periodo_meses: 1 };
  const txsDedup = [
    { id: 'tx_a', fecha: '2026-08-15', importe: 100, cuenta_id: cuentas[0].id,
      descripcion: 'Alquiler', recurrente_id: '' },
    { id: 'tx_b', fecha: '2026-07-15', importe: 100, cuenta_id: cuentas[0].id,
      descripcion: 'Alquiler', recurrente_id: '' },
    { id: 'tx_c', fecha: '2026-09-15', importe: 50, cuenta_id: cuentas[0].id,
      descripcion: 'Otro gasto', recurrente_id: '' }
  ];
  if (txConflictaEnMesRecurrente_(txsDedup, fakeRec, plantAlquiler, '2026-08-15') !== true)
    throw new Error('No detectó original manual mismo mes');
  if (txConflictaEnMesRecurrente_(txsDedup, fakeRec, plantAlquiler, '2026-10-15') !== false)
    throw new Error('Falso positivo en mes vacío');
  if (txConflictaEnMesRecurrente_(
      [{ id: 'x', fecha: '2026-08-20', importe: 999, cuenta_id: cuentas[0].id, descripcion: 'Otra cosa', recurrente_id: '' }],
      fakeRec, plantAlquiler, '2026-08-20') !== false)
    throw new Error('Match por mes pero cuenta/importe/desc distinto no debería chocar');
  if (txConflictaEnMesRecurrente_(
      [{ id: 'y', fecha: '2026-08-15', importe: 100, cuenta_id: cuentas[0].id,
         descripcion: 'Alquiler', recurrente_id: 'rec_OTRO' }],
      fakeRec, plantAlquiler, '2026-08-15') !== false)
    throw new Error('No debe colisionar con otro recurrente del mismo mes');

  // Actualizar una tx recurrente debe actualizar el mismo recurrente y
  // NO crear uno nuevo. Evita la fuga de recurrentes duplicados.
  const antesUpdate = leerHoja('Recurrentes').length;
  const txRec = guardarTransaccion({
    tipo: 'gasto', importe: 50, cuenta_id: cuentas[0].id,
    categoria_id: cat[0].id, descripcion: 'self-rec-update',
    fecha: '2026-08-15', recurrente_plantilla: true,
    recurrente_periodo_meses: 1, recurrente_dia: 15
  });
  if (!txRec.id || !txRec.recurrente_id)
    throw new Error('Crear tx recurrente no produjo recurrente_id');
  const recIdOriginal = txRec.recurrente_id;
  const txRecUpdate = guardarTransaccion({
    id: txRec.id, tipo: 'gasto', importe: 77, cuenta_id: cuentas[0].id,
    categoria_id: cat[0].id, descripcion: 'self-rec-update-V2',
    fecha: '2026-08-15', recurrente_plantilla: true,
    recurrente_periodo_meses: 1, recurrente_dia: 15
  });
  if (txRecUpdate.recurrente_id !== recIdOriginal)
    throw new Error('Update de tx recurrente cambió el id: ' + txRecUpdate.recurrente_id + ' vs ' + recIdOriginal);
  const despuesUpdate = leerHoja('Recurrentes').length;
  if (despuesUpdate !== antesUpdate + 1)
    throw new Error('Update creó un recurrente extra: ' + antesUpdate + ' → ' + despuesUpdate);
  const recFinal = leerHoja('Recurrentes').find(r => r.id === recIdOriginal);
  const plantFinal = JSON.parse(recFinal.plantilla);
  if (plantFinal.importe !== 77 || plantFinal.descripcion !== 'self-rec-update-V2')
    throw new Error('Plantilla no se actualizó: ' + JSON.stringify(plantFinal));
  if (recFinal.ultima_generacion !== '2026-08-15')
    throw new Error('ultima_generacion se perdió/alteró: ' + recFinal.ultima_generacion);
  // Limpieza
  eliminarRecurrente(recIdOriginal);
  eliminarTransaccion(txRec.id);

  // ponytail: smoke-test de devoluciones — verifica que la categoría de gasto
  // queda como real = gasto_original − devolucion, y que el saldo de la cuenta
  // recupera lo devuelto. Mismo mes para que el cálculo del resumen aplique.
  const catGasto = cat.find(c => c.tipo === 'gasto');
  const hoy = new Date();
  const mesResumen = hoy.getFullYear() + '-' + String(hoy.getMonth() + 1).padStart(2, '0');
  const txBase = guardarTransaccion({
    tipo: 'gasto', importe: 50, cuenta_id: cuentas[0].id,
    categoria_id: catGasto.id, descripcion: 'self-devol-base', fecha: isoHoy_()
  });
  const txDev = guardarTransaccion({
    tipo: 'devolucion', importe: 20, cuenta_id: cuentas[0].id,
    categoria_id: catGasto.id, descripcion: 'self-devol-rev', fecha: isoHoy_()
  });
  const categoriasMes = obtenerCategoriasResumen(hoy.getFullYear(), hoy.getMonth() + 1);
  const realGasto = categoriasMes.find(r => r.id === catGasto.id).real;
  if (realGasto !== 30) throw new Error('Devolución no restó del real: ' + realGasto);
  const saldoTras = obtenerCuentas().find(c => c.id === cuentas[0].id).saldo;
  const saldoEsperado = cuentas[0].saldo - 50 + 20;
  if (Math.abs(saldoTras - saldoEsperado) > 0.01) {
    throw new Error('Saldo tras devolución incorrecto: ' + saldoTras + ' vs ' + saldoEsperado);
  }
  // Y la categoría de ingreso no debe haberse tocado.
  const resumenMes = obtenerResumen(hoy.getFullYear(), hoy.getMonth() + 1);
  if (resumenMes.ingresos !== 0) throw new Error('Devolución no debe contar como ingreso');
  if (resumenMes.gastos !== 30) throw new Error('Resumen gastos sin aplicar devolución: ' + resumenMes.gastos);
  // Limpieza
  eliminarTransaccion(txDev.id);
  eliminarTransaccion(txBase.id);

  return 'ok ' + cuentas.length + ' cuentas, ' + cat.length + ' categorías, diferencia=' + conciliado.diferencia;
}

// ───────── Self-test routing de hojas (auth) ─────────
// Cubre resolverHojaActivaId_, vincularHojaUsuarioAdmin (duplicado, demote,
// auto-defecto), desvincularHojaUsuarioAdmin (promoción de defecto) y
// cambiarHojaActiva (persistencia + rechazo). Snapshot/restore al final para
// no dejar rastro en las páginas auth aunque un assert falle. Las pruebas
// tocan solo las hojas de auth, no SpreadsheetApp.openById: los IDs son fake.
function __selfTestHojas() {
  _selfTestActive = true;
  try {
    return __selfTestHojasBody_();
  } finally {
    _selfTestActive = false;
  }
}

function __selfTestHojasBody_() {
  const owner = SELF_TEST_USER;

  const snapSpreads = leerSpreadsheets_();
  const snapLinks = leerHojasUsuarios_();
  const snapUsers = leerUsuariosAuth_();

  try {
    const tag = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    const idA = 'fake_sheet_' + tag + '_a';
    const idB = 'fake_sheet_' + tag + '_b';
    const idC = 'fake_sheet_' + tag + '_c';
    const fakeUser = '__selftest_' + tag;

    const spreads = leerSpreadsheets_().slice();
    spreads.push({ spreadsheet_id: idA, nombre: 'fakeA', descripcion: '', fecha_alta: isoAhora_() });
    spreads.push({ spreadsheet_id: idB, nombre: 'fakeB', descripcion: '', fecha_alta: isoAhora_() });
    spreads.push({ spreadsheet_id: idC, nombre: 'fakeC', descripcion: '', fecha_alta: isoAhora_() });
    escribirSpreadsheets_(spreads);

    // Sembrar un usuario fake vía escritura directa para sortear las
    // comprobaciones internas de vincularHojaUsuarioAdmin.
    const users = leerUsuariosAuth_().slice();
    if (!buscarUsuario_(fakeUser)) {
      const salt = Utilities.getUuid().replace(/-/g, '');
      users.push({
        username: fakeUser,
        password_hash: passwordHash_('selftest1234', salt),
        salt: salt, rol: ROLES.BASICO, activo: true, fecha_creacion: isoAhora_()
      });
      escribirUsuariosAuth_(users);
    }

    let mensajeError;

    // 1. resolverHojaActivaId_: con varias hojas y sin defecto, devuelve la primera.
    limpiarLinksDe_(fakeUser);
    vincularHojaUsuarioInternal_(fakeUser, idA, false);
    vincularHojaUsuarioInternal_(fakeUser, idB, false);
    if (resolverHojaActivaId_(fakeUser) !== idA) {
      throw new Error('Sin defecto debe devolver la primera vinculada');
    }
    // 2. resolverHojaActivaId_: prevalece la marcada como defecto.
    marcarDefecto_(fakeUser, idB);
    if (resolverHojaActivaId_(fakeUser) !== idB) {
      throw new Error('Con defecto debe prevalecer la marcada');
    }
    // 3. resolverHojaActivaId_: con una sola hoja, esa es la activa.
    limpiarLinksDe_(fakeUser);
    vincularHojaUsuarioInternal_(fakeUser, idA, false);
    if (resolverHojaActivaId_(fakeUser) !== idA) {
      throw new Error('Única hoja debe ser la activa');
    }

    // 4. vincularHojaUsuarioAdmin rechaza duplicado explícito.
    mensajeError = null;
    try { vincularHojaUsuarioAdmin(fakeUser, idA, false); }
    catch (e) { mensajeError = e.message; }
    if (!mensajeError || !/ya tiene vinculada/.test(mensajeError)) {
      throw new Error('Debió rechazar duplicado: ' + mensajeError);
    }

    // 5. vincularHojaUsuarioAdmin como defecto demota el defecto anterior.
    limpiarLinksDe_(fakeUser);
    vincularHojaUsuarioInternal_(fakeUser, idA, true);
    vincularHojaUsuarioAdmin(fakeUser, idB, true);
    const links5 = linksDe_(fakeUser);
    if (!links5.find(l => l.spreadsheet_id === idB && l.por_defecto)) {
      throw new Error('Nueva hoja debe haber quedado como defecto');
    }
    if (links5.find(l => l.spreadsheet_id === idA && l.por_defecto)) {
      throw new Error('Defecto anterior no fue desmarcado');
    }

    // 6. La primera vinculación del usuario se marca defecto automáticamente.
    limpiarLinksDe_(fakeUser);
    vincularHojaUsuarioAdmin(fakeUser, idC, false);
    const links6 = linksDe_(fakeUser);
    if (links6.length !== 1 || !links6[0].por_defecto) {
      throw new Error('Primera vinculación debe pasar a defecto auto');
    }

    // 7. desvincularHojaUsuarioAdmin promueve otra hoja al quitar el defecto.
    limpiarLinksDe_(fakeUser);
    vincularHojaUsuarioInternal_(fakeUser, idA, true);
    vincularHojaUsuarioInternal_(fakeUser, idB, false);
    desvincularHojaUsuarioAdmin(fakeUser, idA);
    const links7 = linksDe_(fakeUser);
    if (links7.length !== 1) throw new Error('Debe quedar una sola tras desvinculación');
    if (links7[0].spreadsheet_id !== idB || !links7[0].por_defecto) {
      throw new Error('Hoja restante debe haber sido promovida a defecto');
    }

    // 8. cambiarHojaActiva persiste por_defecto: resolverHojaActivaId_ la ve.
    limpiarLinksDe_(owner);
    vincularHojaUsuarioInternal_(owner, idA, true);
    vincularHojaUsuarioInternal_(owner, idB, false);
    cambiarHojaActiva(idB);
    if (resolverHojaActivaId_(owner) !== idB) {
      throw new Error('cambiarHojaActiva no actualizó el defecto');
    }

    // 9. cambiarHojaActiva rechaza IDs no vinculados al usuario.
    mensajeError = null;
    const idRara = 'fake_no_vinculado_' + Utilities.getUuid().replace(/-/g, '');
    try { cambiarHojaActiva(idRara); }
    catch (e) { mensajeError = e.message; }
    if (!mensajeError || !/No tienes vinculada/.test(mensajeError)) {
      throw new Error('Debió rechazar hoja no vinculada: ' + mensajeError);
    }

    return 'ok — 9 casos de routing de hojas validados';
  } finally {
    escribirSpreadsheets_(snapSpreads);
    escribirHojasUsuarios_(snapLinks);
    escribirUsuariosAuth_(snapUsers);
    _currentSheetId = '';
  }
}

function ejecutarSelfTestAdmin(spreadsheetId) {
  requireAdmin_();
  const sid = String(spreadsheetId || '').trim();
  const hojas = sid ? listarSpreadsheetsAdmin().filter(h => String(h.spreadsheet_id) === sid) : [];
  if (sid && !hojas.length) throw new Error('Hoja no encontrada: ' + sid);
  if (sid) _currentSheetId = sid;
  const resultados = {};
  const errores = [];
  const ejecutar = (nombre, prueba) => {
    const inicio = Date.now();
    try {
      resultados[nombre] = { ok: true, mensaje: String(prueba() || ''), ms: Date.now() - inicio };
    } catch (e) {
      const mensaje = e && e.message ? e.message : String(e);
      resultados[nombre] = { ok: false, mensaje, ms: Date.now() - inicio };
      errores.push({ prueba: nombre, mensaje });
    }
  };
  ejecutar('datos', __selfTest);
  ejecutar('hojas', __selfTestHojas);
  return { ok: errores.length === 0, hoja: sid || null, resultados, errores };
}

function limpiarLinksDe_(username) {
  const restantes = leerHojasUsuarios_().filter(l =>
    String(l.username || '').trim().toLowerCase() !== String(username || '').trim().toLowerCase());
  escribirHojasUsuarios_(restantes);
}

function linksDe_(username) {
  const target = String(username || '').trim().toLowerCase();
  return leerHojasUsuarios_().filter(l => String(l.username || '').trim().toLowerCase() === target);
}

function marcarDefecto_(username, spreadsheetId) {
  const target = String(username || '').trim().toLowerCase();
  const sid = String(spreadsheetId);
  const links = leerHojasUsuarios_();
  links.forEach(l => {
    if (String(l.username || '').trim().toLowerCase() === target) {
      l.por_defecto = String(l.spreadsheet_id) === sid;
    }
  });
  escribirHojasUsuarios_(links);
}

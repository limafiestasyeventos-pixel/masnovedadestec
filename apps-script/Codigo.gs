/**
 * MasNovedadesTEC - API del panel de administración
 * Pega este código en tu hoja de Google Sheets: Extensiones > Apps Script.
 *
 * 1) Cambia ADMIN_USUARIO y ADMIN_CLAVE por tus datos.
 * 2) Guarda (icono de disquete).
 * 3) Implementar > Nueva implementación > Aplicación web
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier persona
 * 4) Copia la "URL de la aplicación web" y pásala para conectar el panel.
 */

// ===== CONFIGURA AQUÍ TU ACCESO =====
const ADMIN_USUARIO = 'admin';
const ADMIN_CLAVE = 'CAMBIA-ESTA-CLAVE';
// ====================================

const CARPETA_FOTOS = 'MasNovedadesTEC - fotos';
const HORAS_SESION = 6;

function doGet() {
  return json({ ok: true, mensaje: 'API de MasNovedadesTEC activa' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.accion === 'login') return json(login(req));
    if (!sesionValida(req.token)) return json({ ok: false, error: 'SESION' });
    switch (req.accion) {
      case 'listar': return json({ ok: true, productos: leerProductos() });
      case 'guardar': return json(guardarProducto(req.producto || {}));
      case 'eliminar': return json(eliminarProducto(req.id));
      case 'destacar': return json(destacarProducto(req.id));
      case 'subirFoto': return json(subirFoto(req));
    }
    return json({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Sesión ---------- */
function login(req) {
  if (ADMIN_CLAVE === 'CAMBIA-ESTA-CLAVE') {
    return { ok: false, error: 'Primero cambia ADMIN_CLAVE en el código de Apps Script y vuelve a implementar.' };
  }
  if (req.usuario !== ADMIN_USUARIO || req.clave !== ADMIN_CLAVE) {
    Utilities.sleep(1500);
    return { ok: false, error: 'Usuario o clave incorrectos' };
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sesion_' + token, '1', HORAS_SESION * 3600);
  return { ok: true, token: token };
}

function sesionValida(token) {
  return !!token && CacheService.getScriptCache().get('sesion_' + token) === '1';
}

/* ---------- Hoja ---------- */
function hoja() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function encabezados(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) {
    return String(h).trim().toLowerCase();
  });
}

function leerProductos() {
  const datos = hoja().getDataRange().getValues();
  const cols = datos.shift().map(function (h) { return String(h).trim().toLowerCase(); });
  return datos
    .filter(function (fila) { return fila.some(function (c) { return c !== ''; }); })
    .map(function (fila) {
      const o = {};
      cols.forEach(function (k, i) { o[k] = fila[i]; });
      return o;
    });
}

function buscarFila(datos, colId, id) {
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][colId]) === String(id)) return i + 1;
  }
  return -1;
}

function guardarProducto(p) {
  if (!String(p.nombre || '').trim()) return { ok: false, error: 'Falta el nombre del producto' };
  if (!(Number(p.precio) > 0)) return { ok: false, error: 'El precio debe ser un número mayor que 0' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja();
    const cols = encabezados(sh);
    const datos = sh.getDataRange().getValues();
    const colId = cols.indexOf('id');
    let id = String(p.id || '').trim();
    let fila = id ? buscarFila(datos, colId, id) : -1;

    if (fila < 0) {
      const usados = {};
      datos.slice(1).forEach(function (r) { usados[String(r[colId])] = true; });
      const base = id || slug(p.nombre) || 'producto';
      id = base;
      let n = 2;
      while (usados[id]) id = base + '-' + (n++);
      if (!p.activo) p.activo = 'SI';
      if (!p.destacado) p.destacado = 'NO';
    }
    p.id = id;

    const nueva = cols.map(function (k, i) {
      if (p[k] !== undefined) return p[k];
      return fila > 0 ? datos[fila - 1][i] : '';
    });
    if (fila > 0) sh.getRange(fila, 1, 1, cols.length).setValues([nueva]);
    else sh.appendRow(nueva);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function eliminarProducto(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja();
    const fila = buscarFila(sh.getDataRange().getValues(), encabezados(sh).indexOf('id'), id);
    if (fila < 0) return { ok: false, error: 'No se encontró el producto' };
    sh.deleteRow(fila);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function destacarProducto(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja();
    const cols = encabezados(sh);
    const colId = cols.indexOf('id');
    const colDest = cols.indexOf('destacado');
    const datos = sh.getDataRange().getValues();
    if (buscarFila(datos, colId, id) < 0) return { ok: false, error: 'No se encontró el producto' };
    const valores = datos.slice(1).map(function (r) { return [String(r[colId]) === String(id) ? 'SI' : 'NO']; });
    if (valores.length) sh.getRange(2, colDest + 1, valores.length, 1).setValues(valores);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Fotos ---------- */
function carpetaFotos() {
  const it = DriveApp.getFoldersByName(CARPETA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FOTOS);
}

function subirFoto(req) {
  if (!req.datos) return { ok: false, error: 'No llegó la imagen' };
  const tipo = /^image\/(jpeg|png|webp)$/.test(req.tipo) ? req.tipo : 'image/jpeg';
  const blob = Utilities.newBlob(Utilities.base64Decode(req.datos), tipo, slug(req.nombre || 'foto') + '-' + Date.now());
  const archivo = carpetaFotos().createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + archivo.getId() + '&sz=w1000' };
}

function slug(texto) {
  return String(texto || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
    .slice(0, 40);
}

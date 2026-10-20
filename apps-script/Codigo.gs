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
 *
 * Pagos con Mercado Pago:
 *   Configuración del proyecto (engranaje) > Propiedades de la secuencia de comandos >
 *   Agregar propiedad:  MP_ACCESS_TOKEN  =  tu Access Token de Mercado Pago.
 *   Nunca pegues el Access Token dentro del código ni lo compartas.
 */

// ===== CONFIGURA AQUÍ TU ACCESO =====
const ADMIN_USUARIO = 'admin';
const ADMIN_CLAVE = 'CAMBIA-ESTA-CLAVE';
// ====================================

const CARPETA_FOTOS = 'MasNovedadesTEC - fotos';
const HORAS_SESION = 6;
const TIENDA_URL = 'https://limafiestasyeventos-pixel.github.io/masnovedadestec/';
const ENVIO_LIMA = 10;
const ENVIO_PROVINCIA = 15;
const ENVIO_GRATIS_DESDE = 99;
const HOJA_PEDIDOS = 'Pedidos';
const COLUMNAS_PEDIDOS = ['fecha', 'pedido', 'estado', 'metodo', 'nombre', 'celular', 'dni', 'departamento',
  'distrito', 'direccion', 'productos', 'subtotal', 'envio', 'total', 'pago_id', 'stock_descontado', 'items'];

function doGet() {
  return json({ ok: true, mensaje: 'API de MasNovedadesTEC activa' });
}

function doPost(e) {
  try {
    const param = (e && e.parameter) || {};
    let req = {};
    try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { req = {}; }

    // Aviso automático de Mercado Pago (webhook)
    if (param.mp || req.type === 'payment' || req.topic === 'payment' || param.topic === 'payment') {
      const pagoId = (req.data && req.data.id) || param['data.id'] || param.id;
      if (pagoId) procesarPago(pagoId);
      return json({ ok: true });
    }

    // Acciones públicas de la tienda
    if (req.accion === 'login') return json(login(req));
    if (req.accion === 'crearPedido') return json(crearPedido(req));
    if (req.accion === 'verificarPedido') return json(verificarPedido(req.pedido));

    // Acciones del panel (requieren sesión)
    const s = sesion(req.token);
    if (!s) return json({ ok: false, error: 'SESION' });
    if (SOLO_DUENO.indexOf(req.accion) >= 0 && s.rol !== 'dueno') {
      return json({ ok: false, error: 'Solo el dueño de la tienda puede hacer esto' });
    }
    switch (req.accion) {
      case 'listar': return json({ ok: true, productos: leerProductos() });
      case 'guardar': return json(guardarProducto(req.producto || {}));
      case 'eliminar': return json(eliminarProducto(req.id));
      case 'destacar': return json(destacarProducto(req.id));
      case 'subirFoto': return json(subirFoto(req));
      case 'listarPedidos': return json({ ok: true, pedidos: listarPedidos() });
      case 'estadoPedido': return json(cambiarEstadoPedido(req.pedido, req.estado));
      case 'listarUsuarios': return json({ ok: true, usuarios: listarUsuarios() });
      case 'crearUsuario': return json(crearUsuario(req.usuario, req.clave));
      case 'claveUsuario': return json(cambiarClaveUsuario(req.usuario, req.clave));
      case 'activarUsuario': return json(activarUsuario(req.usuario, req.activo));
      case 'eliminarUsuario': return json(eliminarUsuario(req.usuario));
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
  const usuario = String(req.usuario || '').trim();
  let info = null;
  if (usuario === ADMIN_USUARIO && req.clave === ADMIN_CLAVE) {
    info = { usuario: ADMIN_USUARIO, rol: 'dueno' };
  } else {
    const u = buscarUsuario(usuario.toLowerCase());
    if (u && u.datos.activo !== 'NO' && hashClave(req.clave, u.datos.sal) === u.datos.clave) {
      info = { usuario: u.datos.usuario, rol: 'ayudante' };
    }
  }
  if (!info) {
    Utilities.sleep(1500);
    return { ok: false, error: 'Usuario o clave incorrectos' };
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sesion_' + token, JSON.stringify(info), HORAS_SESION * 3600);
  return { ok: true, token: token, usuario: info.usuario, rol: info.rol };
}

/* Devuelve {usuario, rol} o null. Un ayudante pausado o eliminado pierde el acceso al instante. */
function sesion(token) {
  if (!token) return null;
  const valor = CacheService.getScriptCache().get('sesion_' + token);
  if (!valor || valor === '1') return null;
  let info;
  try { info = JSON.parse(valor); } catch (x) { return null; }
  if (info.rol === 'ayudante') {
    const u = buscarUsuario(info.usuario);
    if (!u || u.datos.activo === 'NO') return null;
  }
  return info;
}

/* ---------- Administradores (ayudantes) ---------- */
const SOLO_DUENO = ['eliminar', 'listarUsuarios', 'crearUsuario', 'claveUsuario', 'activarUsuario', 'eliminarUsuario'];
const HOJA_USUARIOS = 'Usuarios';

function hojaUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_USUARIOS);
  if (!sh) {
    sh = ss.insertSheet(HOJA_USUARIOS, ss.getNumSheets());
    sh.appendRow(['usuario', 'clave', 'sal', 'rol', 'activo', 'creado']);
    sh.hideSheet();
  }
  return sh;
}

function hashClave(clave, sal) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + String(clave || ''), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function buscarUsuario(usuario) {
  const sh = hojaUsuarios();
  const datos = sh.getDataRange().getValues();
  const cols = datos[0].map(function (h) { return String(h).trim().toLowerCase(); });
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).toLowerCase() === String(usuario).toLowerCase()) {
      const o = {};
      cols.forEach(function (k, j) { o[k] = datos[i][j]; });
      return { fila: i + 1, cols: cols, datos: o, hoja: sh };
    }
  }
  return null;
}

function validarClave(clave) {
  return String(clave || '').length >= 6 ? '' : 'La clave debe tener al menos 6 caracteres';
}

function crearUsuario(usuario, clave) {
  usuario = String(usuario || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) return { ok: false, error: 'El usuario debe tener de 3 a 30 letras o números, sin espacios ni tildes' };
  const errorClave = validarClave(clave);
  if (errorClave) return { ok: false, error: errorClave };
  if (usuario === String(ADMIN_USUARIO).toLowerCase()) return { ok: false, error: 'Ese usuario ya es el tuyo (dueño)' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (buscarUsuario(usuario)) return { ok: false, error: 'Ya existe un administrador con ese usuario' };
    const sal = Utilities.getUuid();
    hojaUsuarios().appendRow([usuario, hashClave(clave, sal), sal, 'ayudante', 'SI',
      Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm')]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function listarUsuarios() {
  const datos = hojaUsuarios().getDataRange().getValues();
  datos.shift();
  const ayudantes = datos.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { usuario: r[0], rol: 'ayudante', activo: r[4] !== 'NO',
      creado: r[5] instanceof Date ? Utilities.formatDate(r[5], 'America/Lima', 'yyyy-MM-dd HH:mm') : r[5] };
  });
  return [{ usuario: ADMIN_USUARIO, rol: 'dueno', activo: true, creado: '' }].concat(ayudantes);
}

function cambiarClaveUsuario(usuario, clave) {
  const errorClave = validarClave(clave);
  if (errorClave) return { ok: false, error: errorClave };
  const u = buscarUsuario(usuario);
  if (!u) return { ok: false, error: 'No se encontró ese administrador' };
  const sal = Utilities.getUuid();
  u.hoja.getRange(u.fila, u.cols.indexOf('clave') + 1).setValue(hashClave(clave, sal));
  u.hoja.getRange(u.fila, u.cols.indexOf('sal') + 1).setValue(sal);
  return { ok: true };
}

function activarUsuario(usuario, activo) {
  const u = buscarUsuario(usuario);
  if (!u) return { ok: false, error: 'No se encontró ese administrador' };
  u.hoja.getRange(u.fila, u.cols.indexOf('activo') + 1).setValue(activo ? 'SI' : 'NO');
  return { ok: true };
}

function eliminarUsuario(usuario) {
  const u = buscarUsuario(usuario);
  if (!u) return { ok: false, error: 'No se encontró ese administrador' };
  u.hoja.deleteRow(u.fila);
  return { ok: true };
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

/* ---------- Pedidos y pagos (Mercado Pago) ---------- */
const ESTADOS_CON_STOCK = ['Pagado', 'Confirmado', 'Enviado', 'Entregado'];

function hojaPedidos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_PEDIDOS);
  if (!sh) {
    sh = ss.insertSheet(HOJA_PEDIDOS, ss.getNumSheets());
    sh.appendRow(COLUMNAS_PEDIDOS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function mercadoPago(metodo, ruta, cuerpo) {
  const token = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
  if (!token) throw new Error('Los pagos en línea aún no están configurados. Puedes pedir por WhatsApp o pagar contraentrega en Lima.');
  const opciones = { method: metodo, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  if (cuerpo) { opciones.contentType = 'application/json'; opciones.payload = JSON.stringify(cuerpo); }
  const res = UrlFetchApp.fetch('https://api.mercadopago.com' + ruta, opciones);
  const datos = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() >= 300) throw new Error('Mercado Pago no respondió bien: ' + (datos.message || res.getResponseCode()));
  return datos;
}

function crearPedido(req) {
  const c = req.cliente || {};
  const metodo = req.metodo === 'contraentrega' ? 'Contraentrega' : 'Mercado Pago';
  const limpio = function (s) { return String(s || '').trim().slice(0, 200); };
  if (!limpio(c.nombre)) return { ok: false, error: 'Escribe tu nombre completo' };
  if (!/^9\d{8}$/.test(limpio(c.celular))) return { ok: false, error: 'Escribe un celular válido de 9 dígitos que empiece con 9' };
  if (!limpio(c.distrito) || !limpio(c.direccion)) return { ok: false, error: 'Completa tu distrito y dirección' };
  const lima = c.departamento === 'Lima' || c.departamento === 'Callao';
  if (metodo === 'Contraentrega' && !lima) return { ok: false, error: 'El pago contraentrega solo está disponible en Lima y Callao' };
  if (metodo !== 'Contraentrega' && !PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN')) {
    return { ok: false, error: 'Los pagos en línea aún no están configurados. Puedes pedir por WhatsApp o pagar contraentrega en Lima.' };
  }

  // Precios y stock se toman de la hoja, nunca del navegador del cliente
  const productos = {};
  leerProductos().forEach(function (p) { productos[String(p.id)] = p; });
  const porProducto = {};
  const lineas = [];
  (req.items || []).forEach(function (it) {
    const p = productos[String(it.id)];
    if (!p || /^no$/i.test(String(p.activo).trim())) throw new Error('Un producto de tu carrito ya no está disponible. Actualiza la página.');
    const q = Math.max(1, Math.min(9, parseInt(it.q, 10) || 1));
    porProducto[p.id] = (porProducto[p.id] || 0) + q;
    const stock = String(p.stock === undefined ? '' : p.stock).trim();
    if (stock !== '' && porProducto[p.id] > Number(stock)) {
      throw new Error(Number(stock) > 0 ? 'Solo quedan ' + stock + ' unidades de ' + p.nombre : p.nombre + ' está agotado');
    }
    lineas.push({ id: String(p.id), titulo: String(p.nombre) + (it.v ? ' (' + limpio(it.v) + ')' : ''), q: q, precio: Number(p.precio) });
  });
  if (!lineas.length) return { ok: false, error: 'Tu carrito está vacío' };

  const subtotal = lineas.reduce(function (s, l) { return s + l.precio * l.q; }, 0);
  const envio = subtotal >= ENVIO_GRATIS_DESDE ? 0 : (lima ? ENVIO_LIMA : ENVIO_PROVINCIA);
  const total = Math.round((subtotal + envio) * 100) / 100;
  const pedido = 'MNT-' + Utilities.formatDate(new Date(), 'America/Lima', 'yyMMdd') + '-' + Math.floor(1000 + Math.random() * 9000);
  const estado = metodo === 'Contraentrega' ? 'Por confirmar' : 'Pendiente de pago';

  const fila = {
    fecha: Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm'), pedido: pedido, estado: estado, metodo: metodo,
    nombre: limpio(c.nombre), celular: limpio(c.celular), dni: limpio(c.dni), departamento: limpio(c.departamento),
    distrito: limpio(c.distrito), direccion: limpio(c.direccion),
    productos: lineas.map(function (l) { return l.q + ' x ' + l.titulo; }).join(' | '),
    subtotal: subtotal, envio: envio, total: total, pago_id: '', stock_descontado: 'NO',
    items: JSON.stringify(lineas.map(function (l) { return { id: l.id, q: l.q }; }))
  };
  const sh = hojaPedidos();
  const cols = encabezados(sh);
  sh.appendRow(cols.map(function (k) { return fila[k] !== undefined ? fila[k] : ''; }));

  if (metodo === 'Contraentrega') return { ok: true, pedido: pedido, total: total };

  const items = lineas.map(function (l) {
    return { id: l.id, title: l.titulo.slice(0, 250), quantity: l.q, unit_price: l.precio, currency_id: 'PEN' };
  });
  if (envio) items.push({ id: 'envio', title: 'Envío', quantity: 1, unit_price: envio, currency_id: 'PEN' });
  const vuelta = TIENDA_URL + '?pedido=' + encodeURIComponent(pedido);
  const preferencia = mercadoPago('post', '/checkout/preferences', {
    items: items,
    external_reference: pedido,
    payer: { name: fila.nombre },
    back_urls: { success: vuelta, failure: vuelta, pending: vuelta },
    auto_return: 'approved',
    notification_url: ScriptApp.getService().getUrl() + '?mp=1',
    statement_descriptor: 'MASNOVEDADESTEC'
  });
  return { ok: true, pedido: pedido, total: total, url: preferencia.init_point };
}

function buscarPedido(sh, pedido) {
  const datos = sh.getDataRange().getValues();
  const cols = datos[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const colPedido = cols.indexOf('pedido');
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][colPedido]) === String(pedido)) {
      const o = {};
      cols.forEach(function (k, j) { o[k] = datos[i][j]; });
      return { fila: i + 1, cols: cols, datos: o };
    }
  }
  return null;
}

function verificarPedido(pedido) {
  if (!pedido) return { ok: false, error: 'Falta el número de pedido' };
  const sh = hojaPedidos();
  let p = buscarPedido(sh, pedido);
  if (!p) return { ok: false, error: 'No encontramos ese pedido' };
  if (p.datos.estado !== 'Pagado' && p.datos.metodo === 'Mercado Pago') {
    const busqueda = mercadoPago('get', '/v1/payments/search?external_reference=' + encodeURIComponent(pedido) + '&sort=date_created&criteria=desc');
    const pagos = busqueda.results || [];
    const aprobado = pagos.filter(function (x) { return x.status === 'approved'; })[0];
    if (aprobado) {
      confirmarPago(pedido, aprobado);
      p = buscarPedido(sh, pedido);
    } else if (pagos[0]) {
      return { ok: true, pedido: pedido, estado: pagos[0].status, total: p.datos.total };
    }
  }
  return { ok: true, pedido: pedido, estado: p.datos.estado === 'Pagado' ? 'approved' : 'pending', total: p.datos.total };
}

function procesarPago(pagoId) {
  const pago = mercadoPago('get', '/v1/payments/' + encodeURIComponent(pagoId));
  if (pago.status === 'approved' && pago.external_reference) confirmarPago(pago.external_reference, pago);
}

function confirmarPago(pedido, pago) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = hojaPedidos();
    const p = buscarPedido(sh, pedido);
    if (!p || p.datos.estado === 'Pagado') return;
    const montoOk = Number(pago.transaction_amount) + 0.01 >= Number(p.datos.total);
    escribirPedido(sh, p, { estado: montoOk ? 'Pagado' : 'Revisar: monto distinto', pago_id: String(pago.id) });
    if (montoOk) ajustarStock(sh, p, -1);
  } finally {
    lock.releaseLock();
  }
}

function escribirPedido(sh, p, cambios) {
  Object.keys(cambios).forEach(function (k) {
    const col = p.cols.indexOf(k);
    if (col >= 0) sh.getRange(p.fila, col + 1).setValue(cambios[k]);
    p.datos[k] = cambios[k];
  });
}

/* signo -1 descuenta stock (venta), +1 lo devuelve (cancelación). Nunca dos veces. */
function ajustarStock(shPedidos, p, signo) {
  const yaDescontado = String(p.datos.stock_descontado).toUpperCase() === 'SI';
  if ((signo < 0 && yaDescontado) || (signo > 0 && !yaDescontado)) return;
  let items = [];
  try { items = JSON.parse(p.datos.items || '[]'); } catch (x) { items = []; }
  const sh = hoja();
  const cols = encabezados(sh);
  const colId = cols.indexOf('id');
  const colStock = cols.indexOf('stock');
  if (colStock >= 0) {
    const datos = sh.getDataRange().getValues();
    items.forEach(function (it) {
      const fila = buscarFila(datos, colId, it.id);
      if (fila < 0) return;
      const actual = datos[fila - 1][colStock];
      if (String(actual).trim() === '') return;
      sh.getRange(fila, colStock + 1).setValue(Math.max(0, Number(actual) + signo * Number(it.q)));
    });
  }
  escribirPedido(shPedidos, p, { stock_descontado: signo < 0 ? 'SI' : 'NO' });
}

function listarPedidos() {
  const datos = hojaPedidos().getDataRange().getValues();
  const cols = datos.shift().map(function (h) { return String(h).trim().toLowerCase(); });
  return datos.slice(-100).reverse().map(function (fila) {
    const o = {};
    cols.forEach(function (k, i) {
      o[k] = fila[i] instanceof Date ? Utilities.formatDate(fila[i], 'America/Lima', 'yyyy-MM-dd HH:mm') : fila[i];
    });
    delete o.items;
    return o;
  });
}

function cambiarEstadoPedido(pedido, estado) {
  const permitidos = ['Por confirmar', 'Pendiente de pago', 'Pagado', 'Confirmado', 'Enviado', 'Entregado', 'Cancelado'];
  if (permitidos.indexOf(estado) < 0) return { ok: false, error: 'Estado no válido' };
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = hojaPedidos();
    const p = buscarPedido(sh, pedido);
    if (!p) return { ok: false, error: 'No se encontró el pedido' };
    escribirPedido(sh, p, { estado: estado });
    if (ESTADOS_CON_STOCK.indexOf(estado) >= 0) ajustarStock(sh, p, -1);
    if (estado === 'Cancelado') ajustarStock(sh, p, +1);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

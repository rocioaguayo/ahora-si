/**
 * Ahora Sí · Control financiero del hogar
 * Backend en Google Apps Script sobre una planilla de Google Sheets.
 *
 * Configuración (una sola vez):
 *   Editor de Apps Script > Configuración del proyecto > Propiedades del script
 *   Agregar la propiedad SHEET_ID con el ID de la planilla.
 *   El ID nunca va escrito en el código: así el repositorio puede ser público.
 *
 * Hojas que usa:
 *   Usuarios:     id | nombre | pinHash | ingreso | metas | fechaRegistro | correo | config (JSON)
 *   Movimientos:  id | userId | tipo | categoria | monto | descripcion | metaAhorro | fecha
 */

const TIPOS = ["INGRESO", "GASTO", "AHORRO"];
const METAS_POR_DEFECTO = "Casa Propia,Vehículo,Fondo Emergencia,Otros";
const SESION_SEGUNDOS = 6 * 60 * 60;      // la sesión dura 6 horas
const MAX_INTENTOS = 5;                   // intentos fallidos antes de bloquear
const BLOQUEO_SEGUNDOS = 15 * 60;         // bloqueo de 15 minutos
const MAX_TEXTO = 120;

// ---------------------------------------------------------------------------
// Acceso a la planilla
// ---------------------------------------------------------------------------

function openSS_() {
  const id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!id) throw new Error("Falta configurar la propiedad SHEET_ID del script.");
  return SpreadsheetApp.openById(id);
}

// Busca la pestaña sin importar espacios ni mayúsculas
function getSheet_(name) {
  const ss = openSS_();
  const target = name.trim().toLowerCase();
  const found = ss.getSheets().find(s => s.getName().trim().toLowerCase() === target);
  return found || ss.insertSheet(name);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Ahora Sí · Control financiero")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ---------------------------------------------------------------------------
// Utilidades de seguridad
// ---------------------------------------------------------------------------

function hashPin_(pin, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ":" + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function crearHash_(pin) {
  const salt = Utilities.getUuid();
  return "sha256$" + salt + "$" + hashPin_(pin, salt);
}

function pinCorrecto_(pin, guardado) {
  guardado = String(guardado);
  if (guardado.indexOf("sha256$") === 0) {
    const partes = guardado.split("$");
    return hashPin_(pin, partes[1]) === partes[2];
  }
  // Compatibilidad: cuentas antiguas con el PIN guardado sin cifrar
  return guardado === String(pin);
}

// Evita que un texto se interprete como fórmula en Sheets (=, +, -, @)
function textoSeguro_(valor) {
  let t = String(valor == null ? "" : valor).trim().slice(0, MAX_TEXTO);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function escaparHtml_(t) {
  return String(t).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function correoValido_(c) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c);
}

function pinValido_(p) {
  return /^\d{4,6}$/.test(String(p));
}

// Sesiones en caché: el navegador solo conoce un token temporal, nunca el id interno
function crearSesion_(userId) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put("ses_" + token, userId, SESION_SEGUNDOS);
  return token;
}

function usuarioDeSesion_(token) {
  if (!token) throw new Error("Sesión no válida. Vuelve a ingresar.");
  const cache = CacheService.getScriptCache();
  const userId = cache.get("ses_" + token);
  if (!userId) throw new Error("La sesión expiró. Vuelve a ingresar.");
  cache.put("ses_" + token, userId, SESION_SEGUNDOS); // renueva la sesión
  return userId;
}

function buscarUsuarioPorId_(userId) {
  const data = getSheet_("Usuarios").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) return { fila: i + 1, row: data[i] };
  }
  return null;
}

function usuarioPublico_(row) {
  return {
    nombre: row[1],
    ingreso: Number(row[3] || 0),
    metas: row[4] ? String(row[4]).split(",").map(m => m.trim()).filter(Boolean) : METAS_POR_DEFECTO.split(","),
    correo: row[6] ? String(row[6]) : "",
    config: leerConfig_(row)
  };
}

// ---------------------------------------------------------------------------
// Registro y acceso
// ---------------------------------------------------------------------------

function registrarUsuario(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const nombre = textoSeguro_(data && data.nombre);
    const correo = String((data && data.correo) || "").trim().toLowerCase();
    const pin = String((data && data.pin) || "").trim();
    const ingreso = Math.max(0, Math.round(Number(data && data.ingreso) || 0));

    if (!nombre) return { success: false, error: "Ingresa tu nombre." };
    if (!correoValido_(correo)) return { success: false, error: "El correo no es válido." };
    if (!pinValido_(pin)) return { success: false, error: "El PIN debe tener entre 4 y 6 números." };

    const sheet = getSheet_("Usuarios");
    const filas = sheet.getDataRange().getValues();
    const existe = filas.slice(1).some(r => String(r[6]).trim().toLowerCase() === correo);
    if (existe) return { success: false, error: "Ya existe una cuenta con ese correo." };

    sheet.appendRow([Utilities.getUuid(), nombre, crearHash_(pin), ingreso, METAS_POR_DEFECTO, new Date(), correo, "{}"]);
    enviarMailBienvenida_(correo, nombre);
    return { success: true };
  } catch (err) {
    return { success: false, error: "No se pudo crear la cuenta. Intenta de nuevo." };
  } finally {
    lock.releaseLock();
  }
}

function enviarMailBienvenida_(correo, nombre) {
  try {
    const n = escaparHtml_(nombre);
    const html = `
      <div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:24px;border:1px solid #eee3d9;border-radius:16px;background:#fffcf9;color:#3b2a20">
        <h2 style="color:#c2185b;margin-top:0">Hola, ${n}</h2>
        <p style="line-height:1.6">Tu cuenta en <b>Ahora Sí</b> quedó creada. Ya puedes registrar ingresos, gastos y ahorros, y ver tu resumen del mes.</p>
        <p style="line-height:1.6">Para entrar usa este correo y el PIN que elegiste. Por seguridad, nunca te lo enviaremos por correo.</p>
        <p style="font-size:12px;color:#8a776b;margin-bottom:0">Si no creaste esta cuenta, puedes ignorar este mensaje.</p>
      </div>`;
    GmailApp.sendEmail(correo, "Tu cuenta en Ahora Sí está lista", "Tu cuenta en Ahora Sí quedó creada.", { htmlBody: html });
  } catch (e) {
    Logger.log("Error al enviar correo: " + e.message);
  }
}

/**
 * Login con correo + PIN. Se acepta el nombre para cuentas antiguas sin correo.
 * Bloquea 15 minutos tras 5 intentos fallidos y migra PIN antiguos a hash.
 */
function login(usuario, pin) {
  try {
    const id = String(usuario || "").trim().toLowerCase();
    pin = String(pin || "").trim();
    if (!id || !pin) return { ok: false, error: "Ingresa tu correo y tu PIN." };

    const cache = CacheService.getScriptCache();
    const claveIntentos = "int_" + Utilities.base64EncodeWebSafe(id);
    const intentos = Number(cache.get(claveIntentos) || 0);
    if (intentos >= MAX_INTENTOS) return { ok: false, error: "Demasiados intentos. Espera 15 minutos." };

    const sheet = getSheet_("Usuarios");
    const data = sheet.getDataRange().getValues();
    const porCorreo = id.indexOf("@") > -1;

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const coincide = porCorreo
        ? String(r[6]).trim().toLowerCase() === id
        : String(r[1]).trim().toLowerCase() === id;
      if (coincide && pinCorrecto_(pin, r[2])) {
        if (String(r[2]).indexOf("sha256$") !== 0) {
          sheet.getRange(i + 1, 3).setValue(crearHash_(pin)); // migración transparente
        }
        cache.remove(claveIntentos);
        return { ok: true, token: crearSesion_(r[0]), user: usuarioPublico_(r) };
      }
    }
    cache.put(claveIntentos, String(intentos + 1), BLOQUEO_SEGUNDOS);
    return { ok: false, error: "Correo o PIN incorrectos." };
  } catch (e) {
    return { ok: false, error: "No se pudo ingresar. Intenta de nuevo." };
  }
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove("ses_" + token);
  return { ok: true };
}

// ---------------------------------------------------------------------------
// Configuración del usuario: presupuesto por categoría y objetivo de cada meta
// ---------------------------------------------------------------------------

function leerConfig_(row) {
  try { return row[7] ? JSON.parse(row[7]) : {}; } catch (e) { return {}; }
}

function guardarConfig(token, config) {
  try {
    const userId = usuarioDeSesion_(token);
    const u = buscarUsuarioPorId_(userId);
    if (!u) throw new Error("Usuario no encontrado.");
    const limpio = { presupuestos: {}, objetivos: {} };
    const metas = usuarioPublico_(u.row).metas;
    Object.keys((config && config.presupuestos) || {}).forEach(function (c) {
      if (CATEGORIAS.indexOf(c) === -1) return;
      const v = Math.round(Number(config.presupuestos[c]));
      if (isFinite(v) && v > 0 && v < 1e10) limpio.presupuestos[c] = v;
    });
    Object.keys((config && config.objetivos) || {}).forEach(function (m) {
      if (metas.indexOf(m) === -1) return;
      const v = Math.round(Number(config.objetivos[m]));
      if (isFinite(v) && v > 0 && v < 1e11) limpio.objetivos[m] = v;
    });
    getSheet_("Usuarios").getRange(u.fila, 8).setValue(JSON.stringify(limpio));
    return { success: true, config: limpio };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---------------------------------------------------------------------------
// Movimientos
// ---------------------------------------------------------------------------

// Valida y normaliza un movimiento recibido desde la pantalla.
function validarMovimiento_(mov, metas) {
  const tipo = String(mov && mov.tipo);
  const monto = Math.round(Number(mov && mov.monto));
  if (TIPOS.indexOf(tipo) === -1) throw new Error("Tipo de movimiento no válido.");
  if (!isFinite(monto) || monto <= 0 || monto > 1e10) throw new Error("Ingresa un monto válido.");
  let meta = "";
  if (tipo === "AHORRO") {
    meta = String(mov.metaAhorro || "");
    if (metas.indexOf(meta) === -1) throw new Error("Meta de ahorro no válida.");
  }
  let categoria = tipo;
  if (tipo === "GASTO") {
    categoria = String(mov.categoria || "Otros");
    if (CATEGORIAS.indexOf(categoria) === -1) categoria = "Otros";
  }
  let fecha = new Date();
  if (mov.fecha && /^\d{4}-\d{2}-\d{2}$/.test(String(mov.fecha))) {
    const p = String(mov.fecha).split("-");
    fecha = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
    const hace = (Date.now() - fecha.getTime()) / 86400000;
    if (hace > 400 || hace < -1) throw new Error("La fecha debe estar dentro del último año.");
  }
  return { tipo: tipo, monto: monto, categoria: categoria, meta: meta, descripcion: textoSeguro_(mov.descripcion), fecha: fecha };
}

function registrarMovimiento(token, mov) {
  try {
    const userId = usuarioDeSesion_(token);
    const u = buscarUsuarioPorId_(userId);
    if (!u) throw new Error("Usuario no encontrado.");
    const m = validarMovimiento_(mov, usuarioPublico_(u.row).metas);
    getSheet_("Movimientos").appendRow([Utilities.getUuid(), userId, m.tipo, m.categoria, m.monto, m.descripcion, m.meta, m.fecha]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// Busca la fila de un movimiento y comprueba que sea del usuario de la sesión.
function filaMovimiento_(userId, id) {
  const data = getSheet_("Movimientos").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (data[i][1] !== userId) throw new Error("Movimiento no encontrado.");
      return i + 1;
    }
  }
  throw new Error("Movimiento no encontrado.");
}

function actualizarMovimiento(token, id, mov) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const userId = usuarioDeSesion_(token);
    const u = buscarUsuarioPorId_(userId);
    if (!u) throw new Error("Usuario no encontrado.");
    const m = validarMovimiento_(mov, usuarioPublico_(u.row).metas);
    const fila = filaMovimiento_(userId, String(id));
    getSheet_("Movimientos").getRange(fila, 3, 1, 6).setValues([[m.tipo, m.categoria, m.monto, m.descripcion, m.meta, m.fecha]]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function eliminarMovimiento(token, id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const userId = usuarioDeSesion_(token);
    const fila = filaMovimiento_(userId, String(id));
    getSheet_("Movimientos").deleteRow(fila);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function leerMovimientos_(userId) {
  const tz = Session.getScriptTimeZone();
  const data = getSheet_("Movimientos").getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[1] !== userId || !(r[7] instanceof Date)) continue;
    lista.push({
      id: r[0],
      mes: Utilities.formatDate(r[7], tz, "yyyy-MM"),
      fecha: Utilities.formatDate(r[7], tz, "dd/MM"),
      fechaISO: Utilities.formatDate(r[7], tz, "yyyy-MM-dd"),
      orden: r[7].getTime() + i,
      tipo: r[2], categoria: String(r[3]), monto: Number(r[4] || 0),
      descripcion: String(r[5] || ""), meta: String(r[6] || "")
    });
  }
  return lista;
}

/**
 * Resumen de un mes (mes = "AAAA-MM"; por defecto el actual).
 * Ingresos, gastos y saldo son del mes. Las metas de ahorro son acumuladas.
 */
function getDashboard(token, mes) {
  try {
    const userId = usuarioDeSesion_(token);
    const u = buscarUsuarioPorId_(userId);
    if (!u) throw new Error("Usuario no encontrado.");
    const usuario = usuarioPublico_(u.row);
    const hoy = new Date();
    const mesActual = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM");
    mes = /^\d{4}-\d{2}$/.test(String(mes)) ? String(mes) : mesActual;
    return calcularResumen(leerMovimientos_(userId), usuario, mes, mesActual, hoy);
  } catch (e) {
    return { error: e.message };
  }
}

// Todos los movimientos del usuario, para descargarlos en Excel (CSV).
function exportarMovimientos(token) {
  try {
    const userId = usuarioDeSesion_(token);
    const lista = leerMovimientos_(userId).sort(function (a, b) { return a.orden - b.orden; });
    return { success: true, movimientos: lista.map(function (m) {
      return { fechaISO: m.fechaISO, tipo: m.tipo, categoria: m.categoria, monto: m.monto, descripcion: m.descripcion, meta: m.meta };
    }) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---------------------------------------------------------------------------
// Cálculo (función pura, también la usa el modo demo de index.html)
// ---------------------------------------------------------------------------

/* ===== Cálculo del resumen (idéntico en Code.gs y en el modo demo) ===== */
var CATEGORIAS = ["Arriendo/Dividendo", "Cuentas Básicas", "Alimentación/Feria", "Educación/Hijos", "Transporte", "Salud", "Ocio", "Otros"];
var NOMBRES_MES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

function mesesAtras_(mes, n) {
  var p = mes.split("-"), a = Number(p[0]), m = Number(p[1]) - 1 - n;
  while (m < 0) { m += 12; a--; }
  while (m > 11) { m -= 12; a++; }
  return a + "-" + ("0" + (m + 1)).slice(-2);
}

function calcularResumen(movimientos, usuario, mes, mesActual, hoy) {
  var cfg = usuario.config || {};
  var presupuestos = cfg.presupuestos || {};
  var objetivos = cfg.objetivos || {};
  var ingresos = 0, gastos = 0, ahorroMes = 0, porCat = {}, historial = [], meses = {}, porMes = {};
  var ahorroMeta = {}, ahorroMetaMes = {};
  usuario.metas.forEach(function (m) { ahorroMeta[m] = 0; ahorroMetaMes[m] = {}; });

  movimientos.forEach(function (m) {
    meses[m.mes] = true;
    var pm = porMes[m.mes] || (porMes[m.mes] = { ingresos: 0, gastos: 0, ahorro: 0 });
    if (m.tipo === "INGRESO") pm.ingresos += m.monto;
    if (m.tipo === "GASTO") pm.gastos += m.monto;
    if (m.tipo === "AHORRO") {
      pm.ahorro += m.monto;
      if (m.mes <= mes) {
        ahorroMeta[m.meta] = (ahorroMeta[m.meta] || 0) + m.monto;
        if (!ahorroMetaMes[m.meta]) ahorroMetaMes[m.meta] = {};
        ahorroMetaMes[m.meta][m.mes] = (ahorroMetaMes[m.meta][m.mes] || 0) + m.monto;
      }
    }
    if (m.mes !== mes) return;
    if (m.tipo === "INGRESO") ingresos += m.monto;
    if (m.tipo === "GASTO") { gastos += m.monto; porCat[m.categoria] = (porCat[m.categoria] || 0) + m.monto; }
    if (m.tipo === "AHORRO") ahorroMes += m.monto;
    historial.push(m);
  });
  historial.sort(function (a, b) { return b.orden - a.orden; });

  var base = ingresos > 0 ? ingresos : usuario.ingreso;
  var saldo = base - gastos - ahorroMes;
  var semanas = 4;
  if (mes === mesActual) {
    var ultimo = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate();
    semanas = Math.max(1, Math.ceil((ultimo - hoy.getDate() + 1) / 7));
  }

  // Presupuesto por categoría
  var nombresCat = CATEGORIAS.slice();
  Object.keys(porCat).concat(Object.keys(presupuestos)).forEach(function (c) { if (nombresCat.indexOf(c) === -1) nombresCat.push(c); });
  var alertas = [], presupuestoTotal = 0;
  var categorias = nombresCat.map(function (c) {
    var g = porCat[c] || 0, p = Number(presupuestos[c]) || 0;
    presupuestoTotal += p;
    var pct = p > 0 ? g / p : 0;
    var estado = p > 0 ? (pct > 1 ? "excedido" : pct >= 0.8 ? "alerta" : "ok") : "sin";
    if (estado === "excedido") alertas.push(c + ": superaste el presupuesto por $" + Math.round(g - p).toLocaleString("es-CL"));
    else if (estado === "alerta") alertas.push(c + ": llevas el " + Math.round(pct * 100) + "% del presupuesto");
    return { categoria: c, gastado: g, presupuesto: p, pct: pct, estado: estado };
  }).filter(function (x) { return x.gastado > 0 || x.presupuesto > 0; })
    .sort(function (a, b) { return b.gastado - a.gastado; });
  if (base > 0 && gastos >= base * 0.9) alertas.unshift("Tus gastos ya llegan al " + Math.round(gastos / base * 100) + "% de tus ingresos del mes");

  // Últimos 6 meses
  var serie = [];
  for (var i = 5; i >= 0; i--) {
    var k = mesesAtras_(mes, i), d = porMes[k] || { ingresos: 0, gastos: 0, ahorro: 0 };
    serie.push({ mes: k, etiqueta: NOMBRES_MES[Number(k.split("-")[1]) - 1], ingresos: d.ingresos, gastos: d.gastos, ahorro: d.ahorro });
  }
  var anterior = porMes[mesesAtras_(mes, 1)];
  var gastosAnterior = anterior ? anterior.gastos : 0;

  // Metas con objetivo y fecha estimada (ritmo = promedio de los últimos 3 meses)
  var metas = usuario.metas.map(function (n) {
    var ahorrado = ahorroMeta[n] || 0, objetivo = Number(objetivos[n]) || 0;
    var ritmo = 0;
    for (var j = 0; j < 3; j++) ritmo += (ahorroMetaMes[n] || {})[mesesAtras_(mes, j)] || 0;
    ritmo = ritmo / 3;
    var faltante = Math.max(0, objetivo - ahorrado), mesesEst = null, fecha = null;
    if (objetivo > 0 && faltante === 0) mesesEst = 0;
    else if (objetivo > 0 && ritmo > 0) {
      mesesEst = Math.ceil(faltante / ritmo);
      var f = mesesAtras_(mes, -mesesEst).split("-");
      fecha = NOMBRES_MES[Number(f[1]) - 1] + " " + f[0];
    }
    return { nombre: n, ahorrado: ahorrado, objetivo: objetivo, pct: objetivo > 0 ? Math.min(1, ahorrado / objetivo) : 0,
             ritmo: Math.round(ritmo), faltante: faltante, mesesEstimados: mesesEst, fechaEstimada: fecha };
  });

  return {
    mes: mes,
    meses: Object.keys(meses).concat([mesActual]).filter(function (v, i, a) { return a.indexOf(v) === i; }).sort().reverse(),
    ingresos: ingresos, ingresoEstimado: usuario.ingreso, usaEstimado: ingresos === 0,
    gastos: gastos, ahorroMes: ahorroMes, saldoDisponible: saldo,
    presupuestoSemanal: saldo > 0 ? Math.floor(saldo / semanas) : 0,
    usoPresupuesto: base > 0 ? gastos / base : 0,
    alertaActiva: base > 0 && gastos >= base * 0.9,
    presupuestoTotal: presupuestoTotal,
    categorias: categorias, alertas: alertas, serie: serie,
    gastosMesAnterior: gastosAnterior,
    variacionGastos: gastosAnterior > 0 ? (gastos - gastosAnterior) / gastosAnterior : null,
    metas: metas,
    historial: historial.slice(0, 300).map(function (m) {
      return { id: m.id, fecha: m.fecha, fechaISO: m.fechaISO, tipo: m.tipo, categoria: m.categoria, monto: m.monto, descripcion: m.descripcion, meta: m.meta };
    })
  };
}

if (typeof module !== "undefined") module.exports = { calcularResumen: calcularResumen };

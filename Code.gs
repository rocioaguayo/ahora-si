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
 *   Usuarios:     id | nombre | pinHash | ingreso | metas | fechaRegistro | correo
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
    correo: row[6] ? String(row[6]) : ""
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

    sheet.appendRow([Utilities.getUuid(), nombre, crearHash_(pin), ingreso, METAS_POR_DEFECTO, new Date(), correo]);
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
// Movimientos y resumen
// ---------------------------------------------------------------------------

function registrarMovimiento(token, mov) {
  try {
    const userId = usuarioDeSesion_(token);
    const u = buscarUsuarioPorId_(userId);
    if (!u) throw new Error("Usuario no encontrado.");

    const tipo = String(mov && mov.tipo);
    const monto = Math.round(Number(mov && mov.monto));
    if (TIPOS.indexOf(tipo) === -1) return { success: false, error: "Tipo de movimiento no válido." };
    if (!isFinite(monto) || monto <= 0 || monto > 1e10) return { success: false, error: "Ingresa un monto válido." };

    let meta = "";
    if (tipo === "AHORRO") {
      meta = String(mov.metaAhorro || "");
      if (usuarioPublico_(u.row).metas.indexOf(meta) === -1) return { success: false, error: "Meta de ahorro no válida." };
    }
    const categoria = tipo === "GASTO" ? (textoSeguro_(mov.categoria) || "Gasto general") : tipo;

    getSheet_("Movimientos").appendRow([
      Utilities.getUuid(), userId, tipo, categoria, monto, textoSeguro_(mov.descripcion), meta, new Date()
    ]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
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
    const tz = Session.getScriptTimeZone();
    const hoy = new Date();
    const mesActual = Utilities.formatDate(hoy, tz, "yyyy-MM");
    mes = /^\d{4}-\d{2}$/.test(String(mes)) ? String(mes) : mesActual;

    const data = getSheet_("Movimientos").getDataRange().getValues();
    const movimientos = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (r[1] !== userId || !(r[7] instanceof Date)) continue;
      movimientos.push({
        mes: Utilities.formatDate(r[7], tz, "yyyy-MM"),
        fecha: Utilities.formatDate(r[7], tz, "dd/MM"),
        orden: r[7].getTime(),
        tipo: r[2], categoria: String(r[3]), monto: Number(r[4] || 0),
        descripcion: String(r[5] || ""), meta: String(r[6] || "")
      });
    }
    return calcularResumen(movimientos, usuario, mes, mesActual, hoy);
  } catch (e) {
    return { error: e.message };
  }
}

/** Cálculo puro (sin Sheets): se puede probar fuera de Apps Script. */
function calcularResumen(movimientos, usuario, mes, mesActual, hoy) {
  let ingresos = 0, gastos = 0, ahorroMes = 0;
  const metasAhorro = {};
  usuario.metas.forEach(m => metasAhorro[m] = 0);
  const porCategoria = {};
  const historial = [];
  const meses = {};

  movimientos.forEach(m => {
    meses[m.mes] = true;
    if (m.tipo === "AHORRO" && m.mes <= mes) metasAhorro[m.meta] = (metasAhorro[m.meta] || 0) + m.monto;
    if (m.mes !== mes) return;
    if (m.tipo === "INGRESO") ingresos += m.monto;
    if (m.tipo === "GASTO") { gastos += m.monto; porCategoria[m.categoria] = (porCategoria[m.categoria] || 0) + m.monto; }
    if (m.tipo === "AHORRO") ahorroMes += m.monto;
    historial.push(m);
  });

  historial.sort((a, b) => b.orden - a.orden);
  const base = ingresos > 0 ? ingresos : usuario.ingreso;
  const saldo = base - gastos - ahorroMes;

  // Presupuesto semanal: saldo repartido en las semanas que quedan del mes
  let semanasRestantes = 4;
  if (mes === mesActual) {
    const ultimoDia = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate();
    const diasRestantes = ultimoDia - hoy.getDate() + 1;
    semanasRestantes = Math.max(1, Math.ceil(diasRestantes / 7));
  }

  return {
    mes: mes,
    meses: Object.keys(meses).concat(mesActual).filter((v, i, a) => a.indexOf(v) === i).sort().reverse(),
    ingresos: ingresos,
    ingresoEstimado: usuario.ingreso,
    usaEstimado: ingresos === 0,
    gastos: gastos,
    ahorroMes: ahorroMes,
    ahorroTotal: Object.keys(metasAhorro).reduce((s, k) => s + metasAhorro[k], 0),
    metasAhorro: metasAhorro,
    saldoDisponible: saldo,
    presupuestoSemanal: saldo > 0 ? Math.floor(saldo / semanasRestantes) : 0,
    usoPresupuesto: base > 0 ? gastos / base : 0,
    alertaActiva: base > 0 && gastos >= base * 0.9,
    gastosPorCategoria: Object.keys(porCategoria).map(k => ({ categoria: k, monto: porCategoria[k] }))
      .sort((a, b) => b.monto - a.monto),
    historial: historial.slice(0, 25).map(m => ({
      fecha: m.fecha, tipo: m.tipo, categoria: m.tipo === "AHORRO" ? "Ahorro · " + m.meta : (m.tipo === "INGRESO" ? "Ingreso" : m.categoria),
      monto: m.monto, descripcion: m.descripcion
    }))
  };
}

if (typeof module !== "undefined") module.exports = { calcularResumen };

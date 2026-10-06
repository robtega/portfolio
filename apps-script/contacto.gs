/**
 * Formulario de contacto de robtegavisual
 * Recibe los envíos del portafolio, te manda un correo y guarda cada respuesta en una hoja de Google Sheets.
 *
 * Se pega en script.google.com y se publica como "Aplicación web" (ver pasos en el chat).
 * No contiene contraseñas ni tokens: el link publicado solo puede recibir formularios.
 */

const CORREO_DESTINO = "robtegacontact@gmail.com";
const GUARDAR_EN_HOJA = true;
const NOMBRE_HOJA = "Contactos del portafolio";

const COLUMNAS = ["Fecha", "Nombre", "Marca / Instagram", "Email", "WhatsApp", "Qué necesita", "Videos al mes", "Presupuesto", "Mensaje", "Estado"];
const ANCHOS = [125, 160, 170, 230, 145, 170, 110, 150, 420, 140];
const ESTADOS = ["Nuevo", "Respondido", "En conversación", "Cliente", "Descartado"];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Campo trampa: las personas no lo ven, los bots lo rellenan
  if (p.sitio_web) return respuesta({ ok: true });

  const dato = (k, max) => String(p[k] || "").trim().slice(0, max || 300);
  const envio = {
    fecha: new Date(),
    nombre: dato("nombre", 120),
    marca: dato("marca", 160),
    email: dato("email", 160),
    whatsapp: dato("whatsapp", 40),
    servicios: dato("servicios", 300),
    videos: dato("videos", 40),
    presupuesto: dato("presupuesto", 60),
    mensaje: dato("mensaje", 3000),
  };

  if (!envio.nombre || !envio.email) return respuesta({ ok: false, error: "Faltan nombre o email" });

  // Evita que la misma persona mande muchos envíos seguidos
  const cache = CacheService.getScriptCache();
  const clave = "envio_" + envio.email.toLowerCase();
  if (cache.get(clave)) return respuesta({ ok: true, repetido: true });
  cache.put(clave, "1", 60);

  enviarCorreo(envio);
  if (GUARDAR_EN_HOJA) guardarEnHoja(envio);
  return respuesta({ ok: true });
}

function enviarCorreo(x) {
  const asunto = "Nuevo proyecto: " + x.nombre + (x.marca ? " · " + x.marca : "");
  const lineas = [
    "Nombre: " + x.nombre,
    "Marca / Instagram: " + (x.marca || "-"),
    "Email: " + x.email,
    "WhatsApp: " + (x.whatsapp || "-"),
    "",
    "Qué necesita: " + (x.servicios || "-"),
    "Videos al mes: " + (x.videos || "-"),
    "Presupuesto mensual: " + (x.presupuesto || "-"),
    "",
    "Mensaje:",
    x.mensaje || "-",
    "",
    "Recibido el " + Utilities.formatDate(x.fecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
  ];
  const valido = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x.email);
  const opciones = { to: CORREO_DESTINO, subject: asunto, body: lineas.join("\n"), name: "Portafolio robtegavisual" };
  if (valido) opciones.replyTo = x.email; // Al pulsar "Responder" en Gmail le contestas al cliente
  MailApp.sendEmail(opciones);
}

/* ---------- Hoja de Google Sheets ---------- */

// Sheets convierte en fórmula lo que empieza con = + - @ (por ejemplo "+58 424…").
// El apóstrofo inicial lo guarda como texto y evita que alguien meta fórmulas desde el formulario.
function comoTexto(v) {
  v = String(v || "");
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function obtenerHoja() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("HOJA_ID");
  let libro = null;
  if (id) { try { libro = SpreadsheetApp.openById(id); } catch (err) { libro = null; } }
  if (!libro) {
    libro = SpreadsheetApp.create(NOMBRE_HOJA);
    props.setProperty("HOJA_ID", libro.getId());
    darFormato(libro.getSheets()[0]);
  }
  return libro.getSheets()[0];
}

function guardarEnHoja(x) {
  const hoja = obtenerHoja();
  hoja.appendRow([
    x.fecha, comoTexto(x.nombre), comoTexto(x.marca), comoTexto(x.email), comoTexto(x.whatsapp),
    comoTexto(x.servicios), comoTexto(x.videos), comoTexto(x.presupuesto), comoTexto(x.mensaje), "Nuevo",
  ]);
}

function darFormato(hoja) {
  const n = COLUMNAS.length, filas = hoja.getMaxRows();
  hoja.getRange(1, 1, 1, n).setValues([COLUMNAS])
    .setFontWeight("bold").setFontColor("#F2EDE3").setBackground("#151312")
    .setVerticalAlignment("middle");
  hoja.setRowHeight(1, 34);
  hoja.setFrozenRows(1);
  ANCHOS.forEach((w, i) => hoja.setColumnWidth(i + 1, w));
  if (filas > 1) {
    hoja.getRange(2, 1, filas - 1, n).setVerticalAlignment("top").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    hoja.getRange(2, 1, filas - 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
    hoja.getRange(2, 9, filas - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Mensaje completo
    const estado = hoja.getRange(2, 10, filas - 1, 1);
    estado.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(false).build());
    const reglas = [
      ["Nuevo", "#F24D2F", "#FFFFFF"], ["Respondido", "#E8E2D6", "#151312"], ["En conversación", "#F2D9A6", "#151312"],
      ["Cliente", "#CFE8D2", "#14361C"], ["Descartado", "#EEEEEE", "#8C857C"],
    ].map(([t, fondo, letra]) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t)
      .setBackground(fondo).setFontColor(letra).setRanges([estado]).build());
    hoja.setConditionalFormatRules(reglas);
  }
}

// Ejecuta esta función una vez para ordenar la hoja que ya existe:
// da formato, corrige los #ERROR! de WhatsApp, borra los envíos de prueba de example.com y marca el estado.
function arreglarHoja() {
  const hoja = obtenerHoja();
  darFormato(hoja);
  const ultima = hoja.getLastRow();
  for (let f = ultima; f >= 2; f--) {
    const email = String(hoja.getRange(f, 4).getValue());
    if (/@example\.com$/i.test(email)) { hoja.deleteRow(f); continue; }
    const celda = hoja.getRange(f, 5), formula = celda.getFormula();
    if (formula) celda.setValue("'" + formula.replace(/^=/, ""));
    const est = hoja.getRange(f, 10);
    if (!est.getValue()) est.setValue("Nuevo");
  }
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Abrir el link en el navegador muestra esto: sirve para comprobar que está publicado
function doGet() {
  return respuesta({ ok: true, mensaje: "Formulario de robtegavisual activo" });
}

// Ejecuta esta función desde el editor para recibir un correo de prueba
function probar() {
  doPost({ parameter: {
    nombre: "Prueba", marca: "@marca_de_prueba", email: CORREO_DESTINO, whatsapp: "+58 000 000 0000",
    servicios: "Reels, Anuncios", videos: "5 a 10", presupuesto: "300 a 800 USD", mensaje: "Este es un envío de prueba.",
  } });
}

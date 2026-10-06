/**
 * Formulario de contacto de robtega.visual
 * Recibe los envíos del portafolio, te manda un correo y guarda cada respuesta en una hoja de Google Sheets.
 *
 * Se pega en script.google.com y se publica como "Aplicación web" (ver pasos en el chat).
 * No contiene contraseñas ni tokens: el link publicado solo puede recibir formularios.
 */

const CORREO_DESTINO = "robtegacontact@gmail.com";
const GUARDAR_EN_HOJA = true;
const NOMBRE_HOJA = "Contactos del portafolio";

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
  const opciones = { to: CORREO_DESTINO, subject: asunto, body: lineas.join("\n"), name: "Portafolio robtega.visual" };
  if (valido) opciones.replyTo = x.email; // Al pulsar "Responder" en Gmail le contestas al cliente
  MailApp.sendEmail(opciones);
}

function guardarEnHoja(x) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("HOJA_ID");
  let libro = null;
  if (id) { try { libro = SpreadsheetApp.openById(id); } catch (err) { libro = null; } }
  if (!libro) {
    libro = SpreadsheetApp.create(NOMBRE_HOJA);
    libro.getSheets()[0].appendRow(["Fecha", "Nombre", "Marca / Instagram", "Email", "WhatsApp", "Qué necesita", "Videos al mes", "Presupuesto", "Mensaje"]);
    props.setProperty("HOJA_ID", libro.getId());
  }
  libro.getSheets()[0].appendRow([x.fecha, x.nombre, x.marca, x.email, x.whatsapp, x.servicios, x.videos, x.presupuesto, x.mensaje]);
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Abrir el link en el navegador muestra esto: sirve para comprobar que está publicado
function doGet() {
  return respuesta({ ok: true, mensaje: "Formulario de robtega.visual activo" });
}

// Ejecuta esta función una vez desde el editor para dar permisos y recibir un correo de prueba
function probar() {
  doPost({ parameter: {
    nombre: "Prueba", marca: "@marca_de_prueba", email: CORREO_DESTINO, whatsapp: "+58 000 000 0000",
    servicios: "Reels, Anuncios", videos: "5 a 10", presupuesto: "300 a 800 USD", mensaje: "Este es un envío de prueba.",
  } });
}

// ============================================
// GOOGLE APPS SCRIPT — Viceconsulado Porlamar
// ============================================
// Pegar este código en:
//   Google Sheets → Extensiones → Apps Script
//
// Funciones:
//   1. Guarda la cita en Google Sheets (CRM)
//   2. Crea evento en Google Calendar automáticamente
//   3. Envía correo de confirmación al solicitante
//   4. Notifica al correo del viceconsulado
// ============================================

// ====== CONFIGURACIÓN ======
var CONFIG = {
  CALENDAR_NAME:       "Citas Viceconsulado",
  EMAIL_CONSULADO:     "espaciosigo@gmail.com",   // ← cambiar a ch.porlamar@maec.es en producción
  EMAIL_NOMBRE:        "Viceconsulado de España — Nueva Esparta",
  DURACION_CITA:       30,   // minutos
  MAX_CITAS_DIA:       10,
  HORA_APERTURA:       8,    // 8:00 AM
  HORA_CIERRE:         12,   // 12:00 PM
};
// ===========================

// -------------------------------------------------------
// doPost: Recibe datos del formulario web
// -------------------------------------------------------
function doPost(e) {
  try {
    var params = e.parameter || {};
    var nombre       = (params.nombre       || "").trim();
    var cedula       = (params.cedula       || "").trim();
    var telefono     = (params.telefono     || "").trim();
    var email        = (params.email        || "").trim();
    var tramite      = (params.tramite      || "").trim();
    var fechaPref    = (params.fecha        || "").trim();
    var observ       = (params.observaciones|| "").trim();
    var fechaRegistro = new Date();

    // 1. Guardar en Sheets
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([
      fechaRegistro,   // A
      nombre,          // B
      cedula,          // C
      telefono,        // D
      email,           // E
      tramite,         // F
      fechaPref,       // G
      observ,          // H
      "Pendiente",     // I
      ""               // J: fecha/hora asignada
    ]);

    // 2. Crear evento en Calendar
    var eventoCreado = false;
    var horaTexto = "";
    if (fechaPref) {
      var resultado = crearEventoCalendar(nombre, cedula, tramite, fechaPref, email, telefono);
      eventoCreado = resultado.ok;
      horaTexto    = resultado.hora;
    }

    // 3. Correo de confirmación al solicitante
    if (email) {
      enviarConfirmacion(nombre, email, tramite, fechaPref, horaTexto);
    }

    // 4. Notificación al viceconsulado
    notificarConsulado(nombre, cedula, telefono, email, tramite, fechaPref, horaTexto, observ);

    return jsonResponse({ status: "ok" });

  } catch (err) {
    Logger.log("Error doPost: " + err);
    return jsonResponse({ status: "error", msg: err.toString() });
  }
}

// -------------------------------------------------------
// doGet: Verificación de que el script está activo
// -------------------------------------------------------
function doGet() {
  return ContentService
    .createTextOutput("Sistema de citas Viceconsulado — activo ✓")
    .setMimeType(ContentService.MimeType.TEXT);
}

// -------------------------------------------------------
// crearEventoCalendar
// -------------------------------------------------------
function crearEventoCalendar(nombre, cedula, tramite, fecha, email, telefono) {
  try {
    // Obtener o crear el calendario
    var cals = CalendarApp.getCalendarsByName(CONFIG.CALENDAR_NAME);
    var cal  = cals.length > 0
      ? cals[0]
      : CalendarApp.createCalendar(CONFIG.CALENDAR_NAME, { color: CalendarApp.Color.RED });

    // Parsear fecha YYYY-MM-DD
    var partes   = fecha.split("-");
    var fechaObj = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));

    // Verificar fin de semana
    var dow = fechaObj.getDay();
    if (dow === 0 || dow === 6) return { ok: false, hora: "" };

    // Encontrar slot libre
    var slot = buscarSlotLibre(cal, fechaObj);
    if (!slot) return { ok: false, hora: "" };

    // Crear evento
    var inicio = new Date(fechaObj);
    inicio.setHours(slot.h, slot.m, 0, 0);
    var fin = new Date(inicio.getTime() + CONFIG.DURACION_CITA * 60000);

    var desc = [
      "TRÁMITE: "    + tramite,
      "NOMBRE: "     + nombre,
      "CÉDULA/PAS: " + cedula,
      "TELÉFONO: "   + telefono,
      "CORREO: "     + email
    ].join("\n");

    cal.createEvent("Cita: " + nombre + " — " + tramite, inicio, fin, {
      description: desc,
      location: "Viceconsulado Honorario de España — Porlamar, Nueva Esparta"
    });

    var horaFmt = pad(slot.h) + ":" + pad(slot.m);
    Logger.log("Evento creado: " + nombre + " " + fecha + " " + horaFmt);
    return { ok: true, hora: horaFmt };

  } catch (err) {
    Logger.log("Error Calendar: " + err);
    return { ok: false, hora: "" };
  }
}

// -------------------------------------------------------
// buscarSlotLibre
// -------------------------------------------------------
function buscarSlotLibre(cal, fechaObj) {
  var dInicio = new Date(fechaObj); dInicio.setHours(CONFIG.HORA_APERTURA, 0, 0, 0);
  var dFin    = new Date(fechaObj); dFin.setHours(CONFIG.HORA_CIERRE,    0, 0, 0);

  var eventos = cal.getEvents(dInicio, dFin);
  if (eventos.length >= CONFIG.MAX_CITAS_DIA) return null;

  var h = CONFIG.HORA_APERTURA, m = 0;
  while (h < CONFIG.HORA_CIERRE) {
    var slotI = new Date(fechaObj); slotI.setHours(h, m, 0, 0);
    var slotF = new Date(slotI.getTime() + CONFIG.DURACION_CITA * 60000);
    if (slotF > dFin) break;

    var libre = true;
    for (var i = 0; i < eventos.length; i++) {
      if (slotI < eventos[i].getEndTime() && slotF > eventos[i].getStartTime()) {
        libre = false; break;
      }
    }
    if (libre) return { h: h, m: m };

    m += CONFIG.DURACION_CITA;
    if (m >= 60) { h += Math.floor(m / 60); m = m % 60; }
  }
  return null;
}

// -------------------------------------------------------
// enviarConfirmacion: correo al solicitante
// -------------------------------------------------------
function enviarConfirmacion(nombre, email, tramite, fecha, hora) {
  var asunto = "Solicitud recibida — Viceconsulado de España en Nueva Esparta";

  var horaInfo = (hora)
    ? "Fecha: " + fecha + "   Hora tentativa: " + hora + "\n(Le confirmaremos la hora exacta por este correo.)"
    : (fecha ? "Fecha solicitada: " + fecha + "\nLe contactaremos para confirmar disponibilidad." : "Le contactaremos para coordinar fecha y hora.");

  var cuerpo =
    "Estimado/a " + nombre + ",\n\n" +
    "Hemos recibido su solicitud de cita para:\n" +
    "  » " + tramite + "\n\n" +
    horaInfo + "\n\n" +
    "─────────────────────────────\n" +
    "RECUERDE EL DÍA DE SU CITA:\n" +
    "  • Traer TODA la documentación requerida (originales y copias)\n" +
    "  • Su cédula venezolana debe estar vigente\n" +
    "  • La cita es personal e intransferible\n" +
    "  • Llegar puntualmente — sin documentación completa no se atiende\n" +
    "─────────────────────────────\n\n" +
    "Consulte los requisitos en:\n" +
    "https://espaciosigo-ai.github.io/viceconsulado-porlamar/\n\n" +
    "Para cancelar o reprogramar: responda este correo o escríbanos por WhatsApp.\n\n" +
    "Atentamente,\n" +
    CONFIG.EMAIL_NOMBRE + "\n" +
    "ch.porlamar@maec.es";

  MailApp.sendEmail({ to: email, subject: asunto, body: cuerpo, name: CONFIG.EMAIL_NOMBRE });
}

// -------------------------------------------------------
// notificarConsulado: aviso interno al viceconsulado
// -------------------------------------------------------
function notificarConsulado(nombre, cedula, telefono, email, tramite, fecha, hora, observ) {
  var asunto = "NUEVA CITA — " + tramite + " — " + nombre;
  var cuerpo =
    "Nueva solicitud de cita recibida en la web:\n\n" +
    "Nombre:      " + nombre   + "\n" +
    "Cédula/Pas:  " + cedula   + "\n" +
    "Teléfono:    " + telefono + "\n" +
    "Correo:      " + email    + "\n" +
    "Trámite:     " + tramite  + "\n" +
    "Fecha pref:  " + (fecha || "No especificada") + "\n" +
    "Hora asig:   " + (hora  || "Pendiente asignación") + "\n" +
    "Observ:      " + (observ || "—") + "\n\n" +
    "Ver en Google Sheets: https://docs.google.com/spreadsheets/d/" +
    SpreadsheetApp.getActiveSpreadsheet().getId();

  MailApp.sendEmail({ to: CONFIG.EMAIL_CONSULADO, subject: asunto, body: cuerpo, name: "Sistema de Citas Web" });
}

// -------------------------------------------------------
// crearEncabezados: ejecutar UNA SOLA VEZ para preparar el Sheet
// -------------------------------------------------------
function crearEncabezados() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName("Citas");

  var cols = ["Fecha Registro","Nombre Completo","Cédula / Pasaporte","Teléfono","Correo",
              "Trámite","Fecha Preferida","Observaciones","Estado","Hora Asignada"];

  sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
  var hdr = sheet.getRange(1, 1, 1, cols.length);
  hdr.setFontWeight("bold").setBackground("#AA151B").setFontColor("white");

  var anchos = [150, 200, 150, 130, 210, 200, 130, 260, 110, 130];
  anchos.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setFrozenRows(1);

  // Crear calendario si no existe
  var cals = CalendarApp.getCalendarsByName(CONFIG.CALENDAR_NAME);
  if (cals.length === 0) {
    CalendarApp.createCalendar(CONFIG.CALENDAR_NAME, { color: CalendarApp.Color.RED });
    Logger.log("Calendario '" + CONFIG.CALENDAR_NAME + "' creado.");
  } else {
    Logger.log("Calendario ya existe.");
  }

  Logger.log("Setup completo. Listo para recibir citas.");
}

// -------------------------------------------------------
// testCompleto: prueba de extremo a extremo (ejecutar manualmente)
// -------------------------------------------------------
function testCompleto() {
  var datos = {
    parameter: {
      nombre:        "Juan Pérez TEST",
      cedula:        "V-12345678",
      telefono:      "0424-0000000",
      email:         CONFIG.EMAIL_CONSULADO,
      tramite:       "Pasaporte — Renovación",
      fecha:         Utilities.formatDate(new Date(new Date().getTime() + 7*24*3600*1000), "America/Caracas", "yyyy-MM-dd"),
      observaciones: "PRUEBA AUTOMÁTICA — borrar"
    }
  };
  var result = doPost(datos);
  Logger.log("Test result: " + result.getContent());
}

// -------------------------------------------------------
// Helpers
// -------------------------------------------------------
function pad(n) { return n < 10 ? "0" + n : "" + n; }
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

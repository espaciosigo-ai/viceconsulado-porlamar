// ============================================
// GOOGLE APPS SCRIPT — Viceconsulado Porlamar
// ============================================
// Pegar este código en:
//   Google Sheets → Extensiones → Apps Script
//
// Funciones:
//   1. Guarda la cita en Google Sheets (CRM)
//   2. Crea evento en Google Calendar automáticamente
//   3. Envía correo HTML de confirmación al solicitante
//   4. Notifica al correo del viceconsulado
// ============================================

// ====== CONFIGURACIÓN ======
var CONFIG = {
  CALENDAR_NAME:       "Citas Viceconsulado",
  EMAIL_CONSULADO:     "espaciosigo@gmail.com",   // ← cambiar a ch.porlamar@maec.es en producción
  EMAIL_NOMBRE:        "Viceconsulado de España — Nueva Esparta",
  DURACION_CITA:       45,   // minutos
  MAX_CITAS_DIA:       5,    // 5 citas × 45 min = 8:00, 8:45, 9:30, 10:15, 11:00
  HORA_APERTURA:       8,    // 8:00 AM
  HORA_CIERRE:         12,   // 12:00 PM
  WEB_URL:             "https://espaciosigo-ai.github.io/viceconsulado-porlamar/",
  WHATSAPP:            "+58 424-8429665",
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
    var horaPref     = (params.hora         || "").trim();
    var observ       = (params.observaciones|| "").trim();
    var fechaRegistro = new Date();

    // 1. Guardar en Sheets
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([
      fechaRegistro,   // A: Fecha Registro
      nombre,          // B: Nombre
      cedula,          // C: Cédula/Pasaporte
      telefono,        // D: Teléfono
      email,           // E: Correo
      tramite,         // F: Trámite
      fechaPref,       // G: Fecha Preferida
      observ,          // H: Observaciones
      "Pendiente",     // I: Estado
      ""               // J: Hora Asignada (se actualiza al crear el evento)
    ]);

    // 2. Crear evento en Calendar
    var horaTexto = "";
    if (fechaPref) {
      var resultado = crearEventoCalendar(nombre, cedula, tramite, fechaPref, email, telefono, horaPref);
      horaTexto    = resultado.hora;

      // Actualizar columna J con la hora asignada
      if (horaTexto) {
        var lastRow = sheet.getLastRow();
        sheet.getRange(lastRow, 10).setValue(horaTexto);
      }
    }

    // 3. Correo HTML de confirmación al solicitante
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
function crearEventoCalendar(nombre, cedula, tramite, fecha, email, telefono, horaPref) {
  try {
    var cals = CalendarApp.getCalendarsByName(CONFIG.CALENDAR_NAME);
    var cal  = cals.length > 0
      ? cals[0]
      : CalendarApp.createCalendar(CONFIG.CALENDAR_NAME, { color: CalendarApp.Color.RED });

    // Parsear fecha YYYY-MM-DD
    var partes   = fecha.split("-");
    var fechaObj = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));

    // Verificar día laborable
    var dow = fechaObj.getDay();
    if (dow === 0 || dow === 6) return { ok: false, hora: "" };

    var slot = null;
    var dFin = new Date(fechaObj); dFin.setHours(CONFIG.HORA_CIERRE, 0, 0, 0);

    // Intentar reservar la hora preferida del usuario
    if (horaPref) {
      var hp = horaPref.split(":");
      var slotPref = { h: parseInt(hp[0]), m: parseInt(hp[1]) };
      var slotI = new Date(fechaObj); slotI.setHours(slotPref.h, slotPref.m, 0, 0);
      var slotF = new Date(slotI.getTime() + CONFIG.DURACION_CITA * 60000);

      if (slotF <= dFin) {
        var conflictos = cal.getEvents(slotI, slotF);
        if (conflictos.length === 0) {
          slot = slotPref;
        }
      }
    }

    // Si la hora preferida no está disponible, buscar el próximo slot libre
    if (!slot) slot = buscarSlotLibre(cal, fechaObj);
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
// enviarConfirmacion: correo HTML al solicitante
// -------------------------------------------------------
function enviarConfirmacion(nombre, email, tramite, fecha, hora) {
  var asunto = "Solicitud recibida — Viceconsulado de España en Nueva Esparta";

  // Formatear fecha legible
  var fechaLegible = "";
  if (fecha) {
    var p = fecha.split("-");
    var meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
    fechaLegible = parseInt(p[2]) + " de " + meses[parseInt(p[1])-1] + " de " + p[0];
  }

  var horaInfoTxt;
  if (hora) {
    horaInfoTxt = "Fecha: " + fechaLegible + "   Hora tentativa: " + hora + "\n(Le confirmaremos la hora exacta por este correo.)";
  } else if (fecha) {
    horaInfoTxt = "Fecha solicitada: " + fechaLegible + "\nLe contactaremos para confirmar disponibilidad.";
  } else {
    horaInfoTxt = "Le contactaremos para coordinar fecha y hora.";
  }

  // -------- HTML EMAIL --------
  var htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background:#FAFAF7;font-family:Arial,Helvetica,sans-serif;">' +

  // Header - bandera España
  '<div style="max-width:600px;margin:0 auto;">' +
  '<div style="height:10px;background:#AA151B;"></div>' +
  '<div style="height:16px;background:#F8CE46;"></div>' +
  '<div style="height:10px;background:#AA151B;"></div>' +

  // Logo bar
  '<div style="background:#AA151B;padding:22px 32px;text-align:center;">' +
  '<p style="color:#F8CE46;font-size:10px;letter-spacing:2px;text-transform:uppercase;margin:0 0 6px 0;">🇪🇸 Ministerio de Asuntos Exteriores</p>' +
  '<h1 style="color:white;font-size:20px;margin:0;font-weight:700;line-height:1.3;">Viceconsulado Honorario de España</h1>' +
  '<p style="color:rgba(255,255,255,0.8);font-size:13px;margin:4px 0 0 0;">Nueva Esparta, Venezuela</p>' +
  '</div>' +

  // Content
  '<div style="background:white;padding:32px;border:1px solid #e8e5df;border-top:none;">' +
  '<p style="color:#AA151B;font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;margin:0 0 20px 0;">SOLICITUD DE CITA — CONFIRMACIÓN</p>' +
  '<p style="font-size:15px;color:#1a1a1a;margin:0 0 20px 0;">Estimado/a <strong>' + nombre + '</strong>,</p>' +
  '<p style="font-size:13px;color:#555;line-height:1.7;margin:0 0 24px 0;">' +
  'Hemos recibido su solicitud de cita para el trámite indicado. Le contactaremos para confirmar la disponibilidad de su fecha.' +
  '</p>' +

  // Appointment details table
  '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;border-radius:8px;overflow:hidden;border:1px solid #e8e5df;">' +
  '<tr><td colspan="2" style="padding:12px 16px;background:#AA151B;color:white;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;">Detalles de su solicitud</td></tr>' +
  '<tr><td style="padding:11px 16px;font-size:12px;font-weight:700;color:#666;background:#fafaf7;border-bottom:1px solid #f0ede8;width:38%;">Trámite</td><td style="padding:11px 16px;font-size:13px;color:#1a1a1a;background:#fafaf7;border-bottom:1px solid #f0ede8;">' + tramite + '</td></tr>' +
  '<tr><td style="padding:11px 16px;font-size:12px;font-weight:700;color:#666;border-bottom:1px solid #f0ede8;">Nombre</td><td style="padding:11px 16px;font-size:13px;color:#1a1a1a;border-bottom:1px solid #f0ede8;">' + nombre + '</td></tr>' +
  '<tr><td style="padding:11px 16px;font-size:12px;font-weight:700;color:#666;background:#fafaf7;border-bottom:1px solid #f0ede8;">Fecha</td><td style="padding:11px 16px;font-size:13px;color:#1a1a1a;background:#fafaf7;border-bottom:1px solid #f0ede8;">' + (fechaLegible || "Por confirmar") + '</td></tr>' +
  '<tr><td style="padding:11px 16px;font-size:12px;font-weight:700;color:#666;">Hora</td><td style="padding:11px 16px;font-size:13px;color:#1a1a1a;">' + (hora ? hora + ' (tentativa)' : 'Por confirmar') + '</td></tr>' +
  '</table>' +

  // Notice box
  '<div style="background:#FFF8E1;border:1px solid #F0E0A0;border-radius:8px;padding:16px 20px;margin-bottom:24px;">' +
  '<p style="font-size:11px;font-weight:700;color:#8B6914;text-transform:uppercase;letter-spacing:1px;margin:0 0 10px 0;">⚠️ Recuerde para el día de su cita</p>' +
  '<ul style="margin:0;padding-left:18px;color:#6B5210;font-size:12px;line-height:1.9;">' +
  '<li>Traer <strong>TODA</strong> la documentación requerida (originales y copias)</li>' +
  '<li>Su cédula venezolana debe estar <strong>vigente</strong></li>' +
  '<li>La cita es <strong>personal e intransferible</strong></li>' +
  '<li>Llegar puntualmente — sin documentación completa no se atiende</li>' +
  '</ul>' +
  '</div>' +

  // Button
  '<div style="text-align:center;margin-bottom:24px;">' +
  '<a href="' + CONFIG.WEB_URL + '" style="background:#AA151B;color:white;text-decoration:none;padding:12px 28px;border-radius:6px;font-size:13px;font-weight:700;display:inline-block;">Consultar requisitos del trámite →</a>' +
  '</div>' +

  // Divider and footer note
  '<p style="font-size:11px;color:#999;border-top:1px solid #f0ede8;padding-top:16px;line-height:1.7;margin:0;">' +
  'Para cancelar o reprogramar: responda este correo o escríbanos por WhatsApp al <strong>' + CONFIG.WHATSAPP + '</strong>.<br>' +
  'Esta es una confirmación provisional. La cita queda sujeta a confirmación del Viceconsulado.' +
  '</p>' +
  '</div>' +

  // Footer
  '<div style="background:#1a1a1a;padding:18px 32px;text-align:center;">' +
  '<p style="color:white;font-size:12px;font-weight:700;margin:0 0 4px 0;">Viceconsulado Honorario de España — Nueva Esparta</p>' +
  '<p style="color:#888;font-size:11px;margin:0;">ch.porlamar@maec.es &nbsp;·&nbsp; Porlamar, Isla de Margarita, Venezuela</p>' +
  '</div>' +

  '</div>' + // max-width container
  '</body></html>';

  // -------- PLAIN TEXT fallback --------
  var cuerpoTxt =
    "Estimado/a " + nombre + ",\n\n" +
    "Hemos recibido su solicitud de cita para:\n  » " + tramite + "\n\n" +
    horaInfoTxt + "\n\n" +
    "─────────────────────────────\n" +
    "RECUERDE EL DÍA DE SU CITA:\n" +
    "  • Traer TODA la documentación requerida (originales y copias)\n" +
    "  • Su cédula venezolana debe estar vigente\n" +
    "  • La cita es personal e intransferible\n" +
    "  • Llegar puntualmente — sin documentación completa no se atiende\n" +
    "─────────────────────────────\n\n" +
    "Consulte los requisitos en:\n" + CONFIG.WEB_URL + "\n\n" +
    "Para cancelar o reprogramar: responda este correo o por WhatsApp al " + CONFIG.WHATSAPP + ".\n\n" +
    "Atentamente,\n" + CONFIG.EMAIL_NOMBRE + "\nch.porlamar@maec.es";

  MailApp.sendEmail({
    to: email,
    subject: asunto,
    body: cuerpoTxt,
    htmlBody: htmlBody,
    name: CONFIG.EMAIL_NOMBRE
  });
}

// -------------------------------------------------------
// notificarConsulado: aviso interno al viceconsulado
// -------------------------------------------------------
function notificarConsulado(nombre, cedula, telefono, email, tramite, fecha, hora, observ) {
  var asunto = "🗓 NUEVA CITA — " + tramite + " — " + nombre;
  var sheetsUrl = "https://docs.google.com/spreadsheets/d/" + SpreadsheetApp.getActiveSpreadsheet().getId();

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:600px;">' +
    '<div style="background:#AA151B;padding:16px 24px;">' +
    '<h2 style="color:white;margin:0;font-size:16px;">🗓 Nueva Solicitud de Cita</h2>' +
    '<p style="color:rgba(255,255,255,0.8);font-size:12px;margin:4px 0 0 0;">Sistema de Citas — Viceconsulado de España · Nueva Esparta</p>' +
    '</div>' +
    '<table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;border-top:none;">' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;width:35%;border-bottom:1px solid #eee;">Nombre</td><td style="padding:10px 16px;font-size:13px;border-bottom:1px solid #eee;">' + nombre + '</td></tr>' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;border-bottom:1px solid #eee;">Cédula/Pasaporte</td><td style="padding:10px 16px;font-size:13px;border-bottom:1px solid #eee;">' + cedula + '</td></tr>' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;border-bottom:1px solid #eee;">Teléfono</td><td style="padding:10px 16px;font-size:13px;border-bottom:1px solid #eee;">' + telefono + '</td></tr>' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;border-bottom:1px solid #eee;">Correo</td><td style="padding:10px 16px;font-size:13px;border-bottom:1px solid #eee;">' + email + '</td></tr>' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;border-bottom:1px solid #eee;">Trámite</td><td style="padding:10px 16px;font-size:13px;font-weight:700;color:#AA151B;border-bottom:1px solid #eee;">' + tramite + '</td></tr>' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;border-bottom:1px solid #eee;">Fecha preferida</td><td style="padding:10px 16px;font-size:13px;border-bottom:1px solid #eee;">' + (fecha || "No especificada") + '</td></tr>' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;border-bottom:1px solid #eee;">Hora asignada</td><td style="padding:10px 16px;font-size:13px;font-weight:700;border-bottom:1px solid #eee;">' + (hora || "Pendiente") + '</td></tr>' +
    '<tr><td style="padding:10px 16px;background:#fafafa;font-size:12px;font-weight:700;color:#666;">Observaciones</td><td style="padding:10px 16px;font-size:13px;">' + (observ || "—") + '</td></tr>' +
    '</table>' +
    '<div style="padding:16px;background:#f5f5f5;text-align:center;">' +
    '<a href="' + sheetsUrl + '" style="background:#AA151B;color:white;text-decoration:none;padding:10px 20px;border-radius:6px;font-size:13px;font-weight:700;">Ver en Google Sheets →</a>' +
    '</div></div>';

  var cuerpoTxt =
    "Nueva solicitud de cita recibida:\n\n" +
    "Nombre:      " + nombre   + "\n" +
    "Cédula/Pas:  " + cedula   + "\n" +
    "Teléfono:    " + telefono + "\n" +
    "Correo:      " + email    + "\n" +
    "Trámite:     " + tramite  + "\n" +
    "Fecha pref:  " + (fecha || "No especificada") + "\n" +
    "Hora asig:   " + (hora  || "Pendiente") + "\n" +
    "Observ:      " + (observ || "—") + "\n\n" +
    "Ver en Sheets: " + sheetsUrl;

  MailApp.sendEmail({
    to: CONFIG.EMAIL_CONSULADO,
    subject: asunto,
    body: cuerpoTxt,
    htmlBody: htmlBody,
    name: "Sistema de Citas Web"
  });
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

  var anchos = [150, 200, 150, 130, 210, 200, 130, 260, 130, 130];
  anchos.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setFrozenRows(1);

  // Aplicar formato dinámico
  configurarFormato(sheet);

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
// configurarFormato: aplica dropdown y colores al Sheet
// Ejecutar TAMBIÉN en sheets existentes para actualizar el formato
// -------------------------------------------------------
function configurarFormato(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Citas")
          || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  // Dropdown para columna Estado (I = col 9)
  var estadoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendiente","Confirmada","Cancelada","Atendida"], true)
    .setAllowInvalid(false)
    .setHelpText("Seleccione el estado de la cita")
    .build();
  sheet.getRange("I2:I1000").setDataValidation(estadoRule);

  // Alineación centrada en columna Estado y Hora
  sheet.getRange("I2:J1000").setHorizontalAlignment("center");

  // Formato de fecha en columna A
  sheet.getRange("A2:A1000").setNumberFormat("dd/MM/yyyy HH:mm");

  // Formato de fecha en columna G
  sheet.getRange("G2:G1000").setNumberFormat("dd/MM/yyyy");

  // Formato de hora en columna J
  sheet.getRange("J2:J1000").setNumberFormat("HH:mm");

  // Formato centrado columna J
  sheet.getRange("J2:J1000").setHorizontalAlignment("center");

  // Formato texto en columna C (cédula - evitar que interprete números)
  sheet.getRange("C2:C1000").setNumberFormat("@");

  // Formato texto en columna D (teléfono)
  sheet.getRange("D2:D1000").setNumberFormat("@");

  // Formato texto en columna E (correo)
  sheet.getRange("E2:E1000").setNumberFormat("@");

  // Formato de fuente para toda la tabla
  sheet.getRange("A2:J1000").setFontFamily("Arial").setFontSize(11);

  // Eliminar reglas previas y aplicar formato condicional por Estado
  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Pendiente")
    .setBackground("#FFF3E0").setFontColor("#E65100").setBold(false)
    .setRanges([sheet.getRange("A2:J1000")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Confirmada")
    .setBackground("#E8F5E9").setFontColor("#2E7D32").setBold(false)
    .setRanges([sheet.getRange("A2:J1000")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Cancelada")
    .setBackground("#FFEBEE").setFontColor("#C62828").setBold(false)
    .setRanges([sheet.getRange("A2:J1000")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Atendida")
    .setBackground("#E3F2FD").setFontColor("#1565C0").setBold(false)
    .setRanges([sheet.getRange("A2:J1000")])
    .build());

  sheet.setConditionalFormatRules(rules);

  // Bandas de color alternas (zebra stripes)
  try {
    var bandedRange = sheet.getRange("A1:J1000");
    var existingBandings = bandedRange.getBandings();
    existingBandings.forEach(function(b) { b.remove(); });
    bandedRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch (err) {
    Logger.log("Banding (ignorar si ya existe): " + err);
  }

  // Bordes en encabezado
  sheet.getRange(1, 1, 1, 10).setBorder(true, true, true, true, true, true, "#AA151B", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  Logger.log("Formato configurado: dropdown Estado, colores condicionales, bandas alternas.");
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
      hora:          "08:45",
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

// ============================================
// GOOGLE APPS SCRIPT — Viceconsulado Porlamar
// ============================================
// Este código va en: Google Sheets → Extensiones → Apps Script
// Recibe datos del formulario web y:
// 1. Los guarda en Google Sheets
// 2. Crea un evento en Google Calendar
// 3. Envía correo de confirmación al solicitante
// ============================================

// ====== CONFIGURACIÓN — CAMBIAR ESTOS VALORES ======
var CONFIG = {
  // Nombre del calendario (el que creaste)
  CALENDAR_NAME: "Citas Viceconsulado",
  
  // Correo desde donde se envían las confirmaciones
  EMAIL_REMITENTE_NOMBRE: "Viceconsulado de España — Nueva Esparta",
  
  // Hora por defecto si no se puede asignar automáticamente (formato 24h)
  HORA_INICIO_DEFAULT: 9,
  
  // Duración de cada cita en minutos
  DURACION_CITA: 30,
  
  // Máximo de citas por día (CAMBIAR cuando la clienta confirme)
  MAX_CITAS_DIA: 10,
  
  // Horario de atención
  HORA_APERTURA: 8,  // 8:00 AM
  HORA_CIERRE: 12,   // 12:00 PM
};
// ====================================================

/**
 * Maneja las peticiones POST del formulario web
 */
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Datos del formulario
    var nombre = e.parameter.nombre || "";
    var cedula = e.parameter.cedula || "";
    var telefono = e.parameter.telefono || "";
    var email = e.parameter.email || "";
    var tramite = e.parameter.tramite || "";
    var fechaPreferida = e.parameter.fecha || "";
    var observaciones = e.parameter.observaciones || "";
    var fechaRegistro = new Date();
    
    // Agregar fila al Sheet
    sheet.appendRow([
      fechaRegistro,          // A: Fecha de registro
      nombre,                 // B: Nombre
      cedula,                 // C: Cédula/Pasaporte
      telefono,               // D: Teléfono
      email,                  // E: Correo
      tramite,                // F: Trámite
      fechaPreferida,         // G: Fecha preferida
      observaciones,          // H: Observaciones
      "Pendiente",            // I: Estado
      ""                      // J: Fecha/hora asignada
    ]);
    
    // Intentar crear evento en Calendar
    var eventoCreado = false;
    if (fechaPreferida) {
      eventoCreado = crearEventoCalendar(nombre, cedula, tramite, fechaPreferida, email, telefono);
    }
    
    // Enviar correo de confirmación
    enviarConfirmacion(nombre, email, tramite, fechaPreferida, eventoCreado);
    
    // Respuesta exitosa
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log("Error en doPost: " + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Maneja peticiones GET (para verificar que funciona)
 */
function doGet(e) {
  return ContentService
    .createTextOutput("El sistema de citas del Viceconsulado está activo.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Crea un evento en Google Calendar
 */
function crearEventoCalendar(nombre, cedula, tramite, fecha, email, telefono) {
  try {
    // Buscar el calendario por nombre
    var calendarios = CalendarApp.getCalendarsByName(CONFIG.CALENDAR_NAME);
    if (calendarios.length === 0) {
      Logger.log("Calendario no encontrado: " + CONFIG.CALENDAR_NAME);
      return false;
    }
    var calendario = calendarios[0];
    
    // Parsear la fecha
    var partes = fecha.split("-");
    var fechaObj = new Date(partes[0], partes[1] - 1, partes[2]);
    
    // Verificar que no sea fin de semana
    var dia = fechaObj.getDay();
    if (dia === 0 || dia === 6) {
      Logger.log("Fecha cae en fin de semana");
      return false;
    }
    
    // Buscar el próximo horario disponible ese día
    var horaAsignada = buscarHorarioDisponible(calendario, fechaObj);
    if (!horaAsignada) {
      Logger.log("No hay horarios disponibles para " + fecha);
      return false;
    }
    
    // Crear el evento
    var inicio = new Date(fechaObj);
    inicio.setHours(horaAsignada.hora, horaAsignada.minutos, 0);
    
    var fin = new Date(inicio);
    fin.setMinutes(fin.getMinutes() + CONFIG.DURACION_CITA);
    
    var descripcion = "Trámite: " + tramite + "\n" +
                      "Nombre: " + nombre + "\n" +
                      "Cédula/Pasaporte: " + cedula + "\n" +
                      "Teléfono: " + telefono + "\n" +
                      "Correo: " + email;
    
    calendario.createEvent(
      "Cita: " + nombre + " — " + tramite,
      inicio,
      fin,
      {
        description: descripcion,
        location: "Viceconsulado Honorario de España — Porlamar"
      }
    );
    
    Logger.log("Evento creado: " + nombre + " a las " + horaAsignada.hora + ":" + horaAsignada.minutos);
    return true;
    
  } catch (error) {
    Logger.log("Error creando evento: " + error.toString());
    return false;
  }
}

/**
 * Busca el próximo horario disponible en un día específico
 */
function buscarHorarioDisponible(calendario, fecha) {
  var inicioDelDia = new Date(fecha);
  inicioDelDia.setHours(CONFIG.HORA_APERTURA, 0, 0);
  
  var finDelDia = new Date(fecha);
  finDelDia.setHours(CONFIG.HORA_CIERRE, 0, 0);
  
  // Obtener eventos existentes ese día
  var eventos = calendario.getEvents(inicioDelDia, finDelDia);
  
  // Verificar límite de citas
  if (eventos.length >= CONFIG.MAX_CITAS_DIA) {
    return null;
  }
  
  // Buscar el primer slot disponible
  var horaActual = CONFIG.HORA_APERTURA;
  var minutosActual = 0;
  
  while (horaActual < CONFIG.HORA_CIERRE) {
    var slotInicio = new Date(fecha);
    slotInicio.setHours(horaActual, minutosActual, 0);
    
    var slotFin = new Date(slotInicio);
    slotFin.setMinutes(slotFin.getMinutes() + CONFIG.DURACION_CITA);
    
    // Verificar que no se pase de la hora de cierre
    if (slotFin.getHours() > CONFIG.HORA_CIERRE || 
        (slotFin.getHours() === CONFIG.HORA_CIERRE && slotFin.getMinutes() > 0)) {
      break;
    }
    
    // Verificar si hay conflicto con algún evento
    var hayConflicto = false;
    for (var i = 0; i < eventos.length; i++) {
      var eventoInicio = eventos[i].getStartTime();
      var eventoFin = eventos[i].getEndTime();
      
      if (slotInicio < eventoFin && slotFin > eventoInicio) {
        hayConflicto = true;
        break;
      }
    }
    
    if (!hayConflicto) {
      return { hora: horaActual, minutos: minutosActual };
    }
    
    // Avanzar al siguiente slot
    minutosActual += CONFIG.DURACION_CITA;
    if (minutosActual >= 60) {
      horaActual += Math.floor(minutosActual / 60);
      minutosActual = minutosActual % 60;
    }
  }
  
  return null; // No hay horarios disponibles
}

/**
 * Envía correo de confirmación al solicitante
 */
function enviarConfirmacion(nombre, email, tramite, fechaPreferida, eventoCreado) {
  if (!email) return;
  
  var asunto = "Solicitud de cita recibida — Viceconsulado de España en Nueva Esparta";
  
  var cuerpo = "Estimado/a " + nombre + ",\n\n" +
    "Hemos recibido su solicitud de cita para: " + tramite + "\n\n";
  
  if (eventoCreado && fechaPreferida) {
    cuerpo += "Fecha solicitada: " + fechaPreferida + "\n" +
      "Le confirmaremos la hora exacta a la brevedad.\n\n";
  } else {
    cuerpo += "Le contactaremos para confirmar fecha y hora disponible.\n\n";
  }
  
  cuerpo += "IMPORTANTE — Recuerde para el día de su cita:\n" +
    "• Traer TODA la documentación requerida para su trámite\n" +
    "• Documentos originales y copias según corresponda\n" +
    "• Llegar puntualmente a la hora asignada\n" +
    "• La cédula venezolana debe estar vigente\n\n" +
    "Si necesita cancelar o reprogramar, responda a este correo o escríbanos por WhatsApp.\n\n" +
    "Consulte los requisitos de su trámite en:\n" +
    "https://espaciosigo-ai.github.io/viceconsulado-porlamar/\n\n" +
    "Atentamente,\n" +
    "Viceconsulado Honorario de España\n" +
    "Nueva Esparta, Venezuela\n" +
    "ch.porlamar@maec.es";
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      body: cuerpo,
      name: CONFIG.EMAIL_REMITENTE_NOMBRE
    });
    Logger.log("Correo enviado a: " + email);
  } catch (error) {
    Logger.log("Error enviando correo: " + error.toString());
  }
}

/**
 * Función para crear los encabezados del Sheet (ejecutar una sola vez)
 */
function crearEncabezados() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var encabezados = [
    "Fecha Registro",
    "Nombre Completo",
    "Cédula / Pasaporte",
    "Teléfono",
    "Correo",
    "Trámite",
    "Fecha Preferida",
    "Observaciones",
    "Estado",
    "Fecha/Hora Asignada"
  ];
  
  // Escribir encabezados en la primera fila
  sheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  
  // Formato de encabezados
  var rango = sheet.getRange(1, 1, 1, encabezados.length);
  rango.setFontWeight("bold");
  rango.setBackground("#AA151B");
  rango.setFontColor("white");
  
  // Ajustar ancho de columnas
  sheet.setColumnWidth(1, 140);  // Fecha registro
  sheet.setColumnWidth(2, 200);  // Nombre
  sheet.setColumnWidth(3, 150);  // Cédula
  sheet.setColumnWidth(4, 130);  // Teléfono
  sheet.setColumnWidth(5, 200);  // Correo
  sheet.setColumnWidth(6, 200);  // Trámite
  sheet.setColumnWidth(7, 120);  // Fecha preferida
  sheet.setColumnWidth(8, 250);  // Observaciones
  sheet.setColumnWidth(9, 100);  // Estado
  sheet.setColumnWidth(10, 150); // Fecha asignada
  
  // Congelar primera fila
  sheet.setFrozenRows(1);
  
  Logger.log("Encabezados creados correctamente");
}

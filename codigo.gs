/**
 * SISTEMA DE GENERACIÓN DE RESOLUCIONES Y REGISTRO EN EXCEL
 * Policía Nacional de Colombia - Jefatura Nacional del Servicio de Policía
 */

const SPREADSHEET_ID = "1BvoAChx88VMGn2jRytV5XHNzGAvHvq6SgDefamxn5Wk";
const SHEET_NAME = "PROYECTO_RESOLUCIONES";

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sin datos recibidos" }))
                 .setMimeType(ContentService.MimeType.JSON);
    }

    const data = JSON.parse(e.postData.contents);
    procesarYGuardarRegistro(data);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
               .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error en doPost: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
               .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("El servicio Web App de Resoluciones está activo.").setMimeType(ContentService.MimeType.TEXT);
}

function procesarYGuardarRegistro(datos) {
  if (!datos) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Fecha Registro", "Acción", "Categoría", "Porcentaje", "Grado", 
      "Nombre Uniformado", "Cédula", "Comunicado Radicación", "Fecha Radicación", 
      "Fecha Hecho", "Días Hábiles", "¿Dentro Plazo?", "Párrafo RESUELVE", "Sede Centro de Conciliación"
    ]);
    sheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#d9ead3");
  } else if (!sheet.getRange(1, 14).getValue()) {
    // Hoja ya existente: agrega el encabezado de la nueva columna N sin mover los datos previos
    sheet.getRange(1, 14).setValue("Sede Centro de Conciliación").setFontWeight("bold").setBackground("#d9ead3");
  }

  let textoResuelve = datos.textoResuelveCustom || "";
  let diasHabilesContados = 0;
  let esDentroPlazo = true;

  // Realizar el cálculo exacto de días hábiles si hay fecha de hecho y de radicación
  if (datos.fechaHecho && datos.fechaRadicacion && datos.tipoAccion !== "EXTINCION" && datos.tipoAccion !== "CONTINUIDADES") {
    const fechaHecho = parseFechaLocal(datos.fechaHecho);
    const fechaRadicacion = parseFechaLocal(datos.fechaRadicacion);
    
    diasHabilesContados = contarDiasHabilesColombia(fechaHecho, fechaRadicacion);
    esDentroPlazo = diasHabilesContados <= 90;
  }

  sheet.appendRow([
    new Date(), 
    datos.tipoAccion, 
    datos.tipoResolucion, 
    datos.tipoPorcentaje || datos.porcentajeContinuidad || "N/A",
    datos.grado, 
    datos.nombreUniformado, 
    datos.cedulaUniformado, 
    datos.comunicadoRadicacion,
    datos.fechaRadicacion, 
    datos.fechaHecho || "N/A", 
    diasHabilesContados, 
    esDentroPlazo ? "SÍ" : "NO", 
    textoResuelve,
    datos.sedeConciliacion || "N/A"
  ]);
}

/**
 * CÁLCULO PRECISO DE DÍAS HÁBILES EN COLOMBIA (EXCLUYE FINES DE SEMANA Y FESTIVOS)
 */
function contarDiasHabilesColombia(fechaInicio, fechaFin) {
  if (!fechaInicio || !fechaFin || isNaN(fechaInicio.getTime()) || isNaN(fechaFin.getTime())) return 0;
  
  let contador = 0;
  let fechaActual = new Date(fechaInicio.getTime());
  
  // Se empieza a contar a partir del día siguiente a la fecha del hecho
  fechaActual.setDate(fechaActual.getDate() + 1);
  
  while (fechaActual <= fechaFin) {
    if (esDiaHabilColombia(fechaActual)) {
      contador++;
    }
    fechaActual.setDate(fechaActual.getDate() + 1);
  }
  return contador;
}

function esDiaHabilColombia(fecha) {
  if (!fecha) return false;
  const diaSemana = fecha.getDay();
  // Sábado (6) y Domingo (0) no son hábiles
  if (diaSemana === 0 || diaSemana === 6) return false;
  
  const anio = fecha.getFullYear();
  const isoStr = formatearFechaISO(fecha);
  
  return !obtenerFestivosColombia(anio).includes(isoStr);
}

function obtenerFestivosColombia(year) {
  let festivos = [
    `${year}-01-01`, `${year}-05-01`, `${year}-07-20`, 
    `${year}-08-07`, `${year}-12-08`, `${year}-12-25`
  ];
  
  const pascua = calcularPascua(year);
  
  // Ley Emiliani y festivos religiosos móviles basados en pascua
  const leyEmiliani = [
    moverALunesDate(new Date(year, 0, 6)),   // Reyes Magos
    moverALunesDate(new Date(year, 2, 19)),  // San José
    moverALunesDate(new Date(year, 5, 29)),  // San Pedro y San Pablo
    moverALunesDate(new Date(year, 7, 15)),  // Asunción de la Virgen
    moverALunesDate(new Date(year, 9, 12)),  // Día de la Raza
    moverALunesDate(new Date(year, 10, 1)),  // Todos los Santos
    moverALunesDate(new Date(year, 10, 11))  // Independencia de Cartagena
  ];

  const festivosMoviles = [
    formatearFechaISO(sumarDias(pascua, -3)), // Jueves Santo
    formatearFechaISO(sumarDias(pascua, -2)), // Viernes Santo
    formatearFechaISO(moverALunesDate(sumarDias(pascua, 39))), // Ascensión del Señor
    formatearFechaISO(moverALunesDate(sumarDias(pascua, 60))), // Corpus Christi
    formatearFechaISO(moverALunesDate(sumarDias(pascua, 68)))  // Sagrado Corazón
  ];

  let todosFestivos = festivos.concat(festivosMoviles);
  leyEmiliani.forEach(f => todosFestivos.push(formatearFechaISO(f)));
  
  return todosFestivos;
}

function moverALunesDate(fecha) {
  const d = new Date(fecha.getTime());
  const dia = d.getDay();
  if (dia !== 1) {
    const diasHastaLunes = (8 - dia) % 7;
    d.setDate(d.getDate() + diasHastaLunes);
  }
  return d;
}

function sumarDias(fecha, dias) {
  const d = new Date(fecha.getTime());
  d.setDate(d.getDate() + dias);
  return d;
}

function calcularPascua(year) {
  const a = year % 19, b = Math.floor(year / 100), c = year % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, L = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * L) / 451);
  const mes = Math.floor((h + L - 7 * m + 114) / 31);
  const dia = ((h + L - 7 * m + 114) % 31) + 1;
  return new Date(year, mes - 1, dia);
}

function parseFechaLocal(fechaStr) {
  if (!fechaStr || !fechaStr.includes("-")) return new Date();
  const p = fechaStr.split("-");
  return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
}

function formatearFechaISO(fecha) {
  const y = fecha.getFullYear();
  const m = String(fecha.getMonth() + 1).padStart(2, "0");
  const d = String(fecha.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatearFechaTexto(fechaStr) {
  if (!fechaStr || !fechaStr.includes("-")) return fechaStr || "";
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const p = fechaStr.split("-");
  return `${parseInt(p[2], 10)} de ${meses[parseInt(p[1], 10) - 1]} del ${p[0]}`;
}
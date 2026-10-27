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
      "Fecha Hecho", "Días Hábiles", "¿Dentro Plazo?", "Párrafo RESUELVE"
    ]);
    sheet.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#d9ead3");
  }

  let textoResuelve = "";
  let diasHabilesContados = 0;
  let esDentroPlazo = true;

  if (datos.tipoAccion === "EXTINCION") {
    const fHechoTxt = formatearFechaTexto(datos.fechaHecho);
    textoResuelve = `EXTINGUIR. Ordenar la extinción del reconocimiento de la Bonificación para la Asistencia Familiar correspondiente al porcentaje del treinta por ciento (30 %) otorgado por razón de la unión marital de hecho entre la señora ${datos.grado} ${datos.nombreUniformado}, identificada con cédula de ciudadanía nro. ${datos.cedulaUniformado}, y el señor ${datos.nombreConyugeExtincion}, identificado con cédula de ciudadanía nro. ${datos.cedulaConyugeExtincion}, a partir del ${fHechoTxt}.`;
  } else if (datos.tipoAccion === "CONTINUIDADES") {
    const fInicioCont = formatearFechaTexto(datos.fechaInicioContinuidad);
    const fFinCont = formatearFechaTexto(datos.fechaFinContinuidad);
    const fCertCont = formatearFechaTexto(datos.fechaCertificadoContinuidad);
    const fRadicacionTxt = formatearFechaTexto(datos.fechaRadicacion);
    textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía N°. ${datos.cedulaUniformado}. Continuidad en el reconocimiento de la bonificación desde el ${fInicioCont} hasta el ${fFinCont}, equivalente al ${datos.porcentajeContinuidad} de la asignación básica del grado correspondiente, por su hijo ${datos.nombreHijoContinuidad}, identificado con cédula de ciudadanía N°. ${datos.cedulaHijoContinuidad}.`;
  } else {
    const fechaHecho = parseFechaLocal(datos.fechaHecho);
    const fechaRadicacion = parseFechaLocal(datos.fechaRadicacion);
    diasHabilesContados = contarDiasHabilesColombia(fechaHecho, fechaRadicacion);
    esDentroPlazo = diasHabilesContados <= 90;

    const fHechoTxt = formatearFechaTexto(datos.fechaHecho);
    const fRadicacionTxt = formatearFechaTexto(datos.fechaRadicacion);

    let textoPlazo = esDentroPlazo 
      ? `Al haberse presentado dentro del término legal (90) días hábiles, los efectos fiscales se reconocerán a partir de la fecha del hecho que lo motiva, esto es, desde el ${fHechoTxt}.`
      : `teniendo en cuenta que la solicitud fue presentada por fuera de los (90) días hábiles. Los efectos fiscales se reconocerán a partir de la fecha de su radicación, esto es, desde el ${fRadicacionTxt}, de conformidad con la norma aplicable.`;

    if (datos.tipoPorcentaje === "30") {
      textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del 30% de la asignación básica correspondiente a su grado, por haber constituido unión marital de hecho con la señora ${datos.nombreConyuge}, identificada con cédula de ciudadanía nro. ${datos.cedulaConyuge}. ${textoPlazo}`;
    } else {
      const pctTexto = datos.tipoPorcentaje === "3" ? "3%" : "2%";
      textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del ${pctTexto} de la asignación básica correspondiente a su grado, por concepto de su hija (${datos.nombreHijo}). ${textoPlazo}`;
    }
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
    textoResuelve
  ]);
}

function contarDiasHabilesColombia(fechaInicio, fechaFin) {
  if (!fechaInicio || !fechaFin || isNaN(fechaInicio.getTime()) || isNaN(fechaFin.getTime())) return 0;
  let contador = 0;
  let fechaActual = new Date(fechaInicio.getTime());
  fechaActual.setDate(fechaActual.getDate() + 1);
  while (fechaActual <= fechaFin) {
    if (esDiaHabilColombia(fechaActual)) contador++;
    fechaActual.setDate(fechaActual.getDate() + 1);
  }
  return contador;
}

function esDiaHabilColombia(fecha) {
  if (!fecha) return false;
  if (typeof fecha === 'string') fecha = parseFechaLocal(fecha);
  if (!fecha || typeof fecha.getFullYear !== 'function') return false;
  
  const diaSemana = fecha.getDay();
  if (diaSemana === 0 || diaSemana === 6) return false;
  return !obtenerFestivosColombia(fecha.getFullYear()).includes(formatearFechaISO(fecha));
}

function obtenerFestivosColombia(year) {
  let festivos = [`${year}-01-01`, `${year}-05-01`, `${year}-07-20`, `${year}-08-07`, `${year}-12-08`, `${year}-12-25`];
  const pascua = calcularPascua(year);
  const leyEmiliani = [
    moverALunesDate(new Date(year, 0, 6)), moverALunesDate(new Date(year, 2, 19)),
    moverALunesDate(new Date(year, 5, 29)), moverALunesDate(new Date(year, 7, 15)),
    moverALunesDate(new Date(year, 9, 12)), moverALunesDate(new Date(year, 10, 1)),
    moverALunesDate(new Date(year, 10, 11))
  ];
  return festivos.concat(leyEmiliani, [
    formatearFechaISO(sumarDias(pascua, -3)), 
    formatearFechaISO(sumarDias(pascua, -2)),
    formatearFechaISO(moverALunesDate(sumarDias(pascua, 39))),
    formatearFechaISO(moverALunesDate(sumarDias(pascua, 60))),
    formatearFechaISO(moverALunesDate(sumarDias(pascua, 68)))
  ]);
}

function moverALunesDate(fecha) {
  if (!fecha || typeof fecha.getDay !== 'function') return new Date();
  const dia = fecha.getDay();
  if (dia !== 1) fecha.setDate(fecha.getDate() + ((8 - dia) % 7));
  return fecha;
}

function sumarDias(fecha, dias) {
  if (!fecha || typeof fecha.getTime !== 'function') return new Date();
  const res = new Date(fecha.getTime());
  res.setDate(res.getDate() + dias);
  return res;
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
  if (!fechaStr || typeof fechaStr !== 'string' || !fechaStr.includes("-")) return new Date();
  const p = fechaStr.split("-");
  return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
}

function formatearFechaISO(fecha) {
  if (!fecha) return "";
  if (typeof fecha === 'string') return fecha;
  if (typeof fecha.getFullYear !== 'function') return "";
  const y = fecha.getFullYear();
  const m = String(fecha.getMonth() + 1).padStart(2, "0");
  const d = String(fecha.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatearFechaTexto(fechaStr) {
  if (!fechaStr || typeof fechaStr !== 'string' || !fechaStr.includes("-")) return fechaStr || "";
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const p = fechaStr.split("-");
  return `${parseInt(p[2], 10)} de ${meses[parseInt(p[1], 10) - 1]} del ${p[0]}`;
}
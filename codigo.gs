/**
 * SISTEMA DE GENERACIÓN DE RESOLUCIONES Y REGISTRO EN EXCEL
 * Policía Nacional de Colombia - Jefatura Nacional del Servicio de Policía
 */

const SPREADSHEET_ID = "1BvoAChx88VMGn2jRytV5XHNzGAvHvq6SgDefamxn5Wk";
const SHEET_NAME = "PROYECTO_RESOLUCIONES";
const FOLDER_ID = "1CTN0RhlbAdJbvy7C0CasXjnRWDJLXVFi";

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("Sin datos").setMimeType(ContentService.MimeType.TEXT);
    }

    const data = JSON.parse(e.postData.contents);

    if (data.accion === "GENERAR_DOCUMENTO_COMPLETO") {
      const resultDoc = crearDocumentoResolucionCompleto(data);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", downloadUrl: resultDoc.downloadUrl }))
                 .setMimeType(ContentService.MimeType.JSON);
    }

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

/**
 * 1. REGISTRAR CADA SOLICITUD EN LA PESTAÑA 'PROYECTO_RESOLUCIONES'
 */
function procesarYGuardarRegistro(datos) {
  if (!datos) return;

  const fechaHecho = parseFechaLocal(datos.fechaHecho);
  const fechaRadicacion = parseFechaLocal(datos.fechaRadicacion);
  
  const diasHabilesContados = contarDiasHabilesColombia(fechaHecho, fechaRadicacion);
  const esDentroPlazo = diasHabilesContados <= 90;

  const fHechoTxt = formatearFechaTexto(datos.fechaHecho);
  const fRadicacionTxt = formatearFechaTexto(datos.fechaRadicacion);

  let textoPlazo = esDentroPlazo 
    ? `Al haberse presentado dentro del término legal (90) días hábiles, los efectos fiscales se reconocerán a partir de la fecha del hecho que lo motiva, esto es, desde el ${fHechoTxt}.`
    : `teniendo en cuenta que la solicitud fue presentada por fuera de los (90) días hábiles. Los efectos fiscales se reconocerán a partir de la fecha de su radicación, esto es, desde el ${fRadicacionTxt}, de conformidad con la norma aplicable.`;

  let textoResuelve = "";
  if (datos.tipoPorcentaje === "30") {
    textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del 30% de la asignación básica correspondiente a su grado, por haber constituido unión marital de hecho con la señora ${datos.nombreConyuge}, identificada con cédula de ciudadanía nro. ${datos.cedulaConyuge}, conforme al ${datos.tipoDocumento} ${datos.numeroDocumento} del ${datos.enteEmisor} del ${fHechoTxt}. La solicitud fue radicada mediante el comunicado oficial ${datos.comunicadoRadicacion} del ${fRadicacionTxt}. ${textoPlazo}`;
  } else {
    const pctTexto = datos.tipoPorcentaje === "3" ? "3%" : "2%";
    textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del ${pctTexto} de la asignación básica correspondiente a su grado, por concepto de su hija (${datos.nombreHijo}), conforme al registro civil de nacimiento nro. ${datos.registroCivil}, que acredita su nacimiento el ${fHechoTxt}. La solicitud fue radicada mediante el comunicado oficial ${datos.comunicadoRadicacion} del ${fRadicacionTxt} ${textoPlazo}`;
  }

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

  sheet.appendRow([
    new Date(), 
    datos.tipoAccion, 
    datos.tipoResolucion, 
    datos.tipoPorcentaje + "%",
    datos.grado, 
    datos.nombreUniformado, 
    datos.cedulaUniformado, 
    datos.comunicadoRadicacion,
    datos.fechaRadicacion, 
    datos.fechaHecho, 
    diasHabilesContados, 
    esDentroPlazo ? "SÍ" : "NO", 
    textoResuelve
  ]);
}

/**
 * 2. CREAR DOCUMENTO EN GOOGLE DOCS Y MOVER A CARPETA DESTINO
 */
function crearDocumentoResolucionCompleto(payload) {
  const data = payload || {};
  const tipoResolucion = data.tipoResolucion || "PP";
  const listaResuelve = Array.isArray(data.listaResuelve) ? data.listaResuelve : [];

  const nombreDoc = `PROYECTO_RESOLUCION_${tipoResolucion}_${new Date().toISOString().substring(0,10)}`;
  const doc = DocumentApp.create(nombreDoc);
  const fileId = doc.getId();

  try {
    const file = DriveApp.getFileById(fileId);
    const targetFolder = DriveApp.getFolderById(FOLDER_ID);
    file.moveTo(targetFolder);
  } catch (e) {
    Logger.log("Aviso Drive: " + e.toString());
  }

  const body = doc.getBody();
  body.setMarginTop(72); body.setMarginBottom(72); body.setMarginLeft(72); body.setMarginRight(72);

  const pCenterBold = (txt) => {
    const p = body.appendParagraph(txt);
    p.setAlignment(DocumentApp.TextAlignment.CENTER);
    p.setFontFamily("Arial"); p.setFontSize(10); p.setBold(true);
    return p;
  };

  const pJustified = (txt) => {
    const p = body.appendParagraph(txt);
    p.setAlignment(DocumentApp.TextAlignment.JUSTIFY);
    p.setFontFamily("Arial"); p.setFontSize(10); p.setBold(false);
    return p;
  };

  pCenterBold("MINISTERIO DE DEFENSA NACIONAL\nPOLICÍA NACIONAL\nJEFATURA NACIONAL DEL SERVICIO DE POLICÍA\n");
  pCenterBold(`RESOLUCIÓN NÚMERO      ${data.numeroResolucion || "0000"}      DEL ${data.fechaResolucion || "DD/MM/AAAA"}\n`);

  const epigrafe = tipoResolucion === "PP"
    ? "“Por la cual se ordena el reconocimiento y actualización de la Bonificación para la Asistencia Familiar a un personal de Patrulleros de Policía en servicio activo de la Policía Nacional”"
    : "“Por la cual se ordena el reconocimiento y actualización de la Bonificación para la Asistencia Familiar a un personal del nivel ejecutivo en servicio activo de la Policía Nacional”";
  
  pCenterBold(epigrafe + "\n");
  pCenterBold("EL JEFE NACIONAL DEL SERVICIO DE POLICÍA\n");
  pJustified("En uso de las facultades conferidas por delegación, mediante resolución nro. 01182 del 04 de mayo de 2022, y\n");
  pCenterBold("CONSIDERANDO:\n");

  pJustified("Que el artículo 132 de la Ley 2179 del 30 de diciembre de 2021 “Por la cual se crea la categoría de Patrulleros de Policía, se establecen normas relacionadas con el régimen especial de carrera del personal uniformado de la Policía Nacional, se fortalece la profesionalización para el servicio público de policía y se dictan otras disposiciones.”, establece:");
  pJustified("“(…)\nARTÍCULO 132. BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. El personal del Nivel Ejecutivo y Patrulleros de Policía en servicio activo, tendrán derecho a partir del año siguiente a la entrada en vigencia de la presente ley, al reconocimiento y pago cada dos meses de una bonificación denominada \"para la asistencia familiar\"...\n(…)”.\n");
  pJustified("Que el personal que se relaciona en la parte resolutiva del presente acto administrativo, solicitó el reconocimiento de la Bonificación para la Asistencia Familiar a que consideran tener derecho.\n");
  pJustified("Que la unión marital de hecho entre compañeros permanentes, se declara por los mecanismos previstos en el artículo 2 de la Ley 979 de 2005.\n");
  pJustified("Que, mediante comunicaciones oficiales nro. GS-2026-000206-OFPLA del 13/01/2026 y GS-2026-006102-DITAH del 18/01/2026 se brindan los lineamientos para el trámite de Pago Pasivos Exigibles Vigencias Expiradas – PPEVE.\n");
  
  pJustified("En mérito de lo expuesto,\n");
  pCenterBold("RESUELVE:\n");

  if (listaResuelve.length > 0) {
    listaResuelve.forEach((texto, idx) => {
      pJustified(`ARTÍCULO ${idx + 1}°. ${texto}\n`);
    });
  }

  pCenterBold("NOTIFÍQUESE Y CÚMPLASE,\n\n\n\n");
  pCenterBold("Mayor general RICARDO AUGUSTO ALARCÓN CAMPOS\nJefe Nacional del Servicio de Policía\n\n");

  const table = body.appendTable();
  const row1 = table.appendTableRow();
  row1.appendTableCell("Elaboró: SI. Fabian Sanguinetti\nJESEP-GUTAH");
  row1.appendTableCell("Revisó: MY. Emma Paola Tique Bonilla\nJESEP-GUTAH");

  const row2 = table.appendTableRow();
  row2.appendTableCell("Revisó: MY. Jamer Alexander Beltran Ladino\nJESEP-ASJUR (E)");
  row2.appendTableCell("Revisó: CR. Hernando Alfredo Calderón Vega\nJESEP-SUSEP");

  pJustified("\nFecha de elaboración: 14-10-2026");
  pJustified("Ubicación: \\\\srvfilesponal3\\disec\\GUTAH\\DISEC - JOHN ALEXANDER BRAVO JIMENEZ\\2026\\4. NOMINA\\3. BONIFICACIONES\\10. OCTUBRE");
  pJustified("Carrera 59 26-21 CAN, Bogotá - Teléfonos 515 9382\njesep.gutah@policia.gov.co - www.policia.gov.co");
  pCenterBold("\nINFORMACIÓN PÚBLICA");

  doc.saveAndClose();
  return { downloadUrl: doc.getUrl() };
}

// FUNCIONES DE DÍAS HÁBILES Y FECHAS BLINDADAS CONTRA ERRORES
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
  if (!fecha || typeof fecha.getFullYear !== 'function') return false;
  const diaSemana = fecha.getDay();
  if (diaSemana === 0 || diaSemana === 6) return false;
  return !obtenerFestivosColombia(fecha.getFullYear()).includes(formatearFechaISO(fecha));
}

function obtenerFestivosColombia(year) {
  let festivos = [`${year}-01-01`, `${year}-05-01`, `${year}-07-20`, `${year}-08-07`, `${year}-12-08`, `${year}-12-25`];
  const pascua = calcularPascua(year);
  const leyEmiliani = [
    moverALunes(new Date(year, 0, 6)), moverALunes(new Date(year, 2, 19)),
    moverALunes(new Date(year, 5, 29)), moverALunes(new Date(year, 7, 15)),
    moverALunes(new Date(year, 9, 12)), moverALunes(new Date(year, 10, 1)),
    moverALunes(new Date(year, 10, 11))
  ];
  return festivos.concat(leyEmiliani, [
    formatearFechaISO(sumarDias(pascua, -3)), formatearFechaISO(sumarDias(pascua, -2)),
    formatearFechaISO(moverALunes(sumarDias(pascua, 39))),
    formatearFechaISO(moverALunes(sumarDias(pascua, 60))),
    formatearFechaISO(moverALunes(sumarDias(pascua, 68)))
  ]);
}

function moverALunes(fecha) {
  const dia = fecha.getDay();
  if (dia !== 1) fecha.setDate(fecha.getDate() + ((8 - dia) % 7));
  return formatearFechaISO(fecha);
}

function sumarDias(fecha, dias) {
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
  if (!fecha || typeof fecha.getFullYear !== 'function') return "";
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
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

    // Si la acción es generar el documento completo de Word
    if (data.accion === "GENERAR_DOCUMENTO_COMPLETO") {
      const resultDoc = crearDocumentoResolucionCompleto(data);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", downloadUrl: resultDoc.downloadUrl }))
                 .setMimeType(ContentService.MimeType.JSON);
    }

    // Si es una solicitud normal para guardar el registro en Google Sheets
    procesarYGuardarRegistro(data);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
               .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("Error en doPost: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
               .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 1. REGISTRAR CADA SOLICITUD EN LA HOJA DE CÁLCULO
 */
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
    const fRadicacionTxt = formatearFechaTexto(datos.fechaRadicacion);
    const fResAntTxt = formatearFechaTexto(datos.fechaResolucionAnterior);
    textoResuelve = `Dejar sin efectos la resolución nro. ${datos.resolucionAnterior} del ${fResAntTxt}, y ordenar la extinción, actualización y descuento de la Bonificación para la Asistencia Familiar reconocida al señor ${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, en lo correspondiente al porcentaje otorgado por razón de su ex compañera permanente, la señora ${datos.nombreConyugeExtincion}, identificada con cédula nro. ${datos.cedulaConyugeExtincion}, con efectos a partir del ${fHechoTxt}. En virtud de lo anterior, disponer el descuento de $ ${datos.valorDescuento} (${datos.textoValorDescuento}) valores reconocidos sin causar derecho a partir del ${fHechoTxt} al ${fRadicacionTxt}.`;
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
      textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del 30% de la asignación básica correspondiente a su grado, por haber constituido unión marital de hecho con la señora ${datos.nombreConyuge}, identificada con cédula de ciudadanía nro. ${datos.cedulaConyuge}, conforme al ${datos.tipoDocumento} ${datos.numeroDocumento} del ${datos.enteEmisor} del ${fHechoTxt}. La solicitud fue radicada mediante el comunicado oficial ${datos.comunicadoRadicacion} del ${fRadicacionTxt}. ${textoPlazo}`;
    } else {
      const pctTexto = datos.tipoPorcentaje === "3" ? "3%" : "2%";
      textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del ${pctTexto} de la asignación básica correspondiente a su grado, por concepto de su hija (${datos.nombreHijo}), conforme al registro civil de nacimiento nro. ${datos.registroCivil}, que acredita su nacimiento el ${fHechoTxt}. La solicitud fue radicada mediante el comunicado oficial ${datos.comunicadoRadicacion} del ${fRadicacionTxt} ${textoPlazo}`;
    }
  }

  sheet.appendRow([
    new Date(), 
    datos.tipoAccion, 
    datos.tipoResolucion, 
    datos.tipoPorcentaje || "N/A",
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
   * 1. REGISTRAR CADA SOLICITUD EN LA HOJA DE CÁLCULO
   */
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
      const fRadicacionTxt = formatearFechaTexto(datos.fechaRadicacion);
      const fResAntTxt = formatearFechaTexto(datos.fechaResolucionAnterior);
      textoResuelve = `Dejar sin efectos la resolución nro. ${datos.resolucionAnterior} del ${fResAntTxt}, y ordenar la extinción, actualización y descuento de la Bonificación para la Asistencia Familiar reconocida al señor ${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, en lo correspondiente al porcentaje otorgado por razón de su ex compañera permanente, la señora ${datos.nombreConyugeExtincion}, identificada con cédula nro. ${datos.cedulaConyugeExtincion}, con efectos a partir del ${fHechoTxt}. En virtud de lo anterior, disponer el descuento de $ ${datos.valorDescuento} (${datos.textoValorDescuento}) valores reconocidos sin causar derecho a partir del ${fHechoTxt} al ${fRadicacionTxt}.`;
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
        textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del 30% de la asignación básica correspondiente a su grado, por haber constituido unión marital de hecho con la señora ${datos.nombreConyuge}, identificada con cédula de ciudadanía nro. ${datos.cedulaConyuge}, conforme al ${datos.tipoDocumento} ${datos.numeroDocumento} del ${datos.enteEmisor} del ${fHechoTxt}. La solicitud fue radicada mediante el comunicado oficial ${datos.comunicadoRadicacion} del ${fRadicacionTxt}. ${textoPlazo}`;
      } else {
        const pctTexto = datos.tipoPorcentaje === "3" ? "3%" : "2%";
        textoResuelve = `${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del ${pctTexto} de la asignación básica correspondiente a su grado, por concepto de su hija (${datos.nombreHijo}), conforme al registro civil de nacimiento nro. ${datos.registroCivil}, que acredita su nacimiento el ${fHechoTxt}. La solicitud fue radicada mediante el comunicado oficial ${datos.comunicadoRadicacion} del ${fRadicacionTxt} ${textoPlazo}`;
      }
    }

    sheet.appendRow([
      new Date(), 
      datos.tipoAccion, 
      datos.tipoResolucion, 
      datos.tipoPorcentaje || "N/A",
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
   * 2. CREAR DOCUMENTO EN GOOGLE DOCS CON MAPEO EXACTO Y COMPLETO
   */
  function crearDocumentoResolucionCompleto(payload) {
    const data = payload || {};
    const tipoAccion = data.tipoAccion || "RECONOCIMIENTO";
    const tipoResolucion = data.tipoResolucion || "PP";
    const listaResuelve = Array.isArray(data.listaResuelve) ? data.listaResuelve : [];

    const sufijoDoc = tipoAccion === "EXTINCION" ? "EXTINCION" : tipoResolucion;
    const nombreDoc = `PROYECTO_RESOLUCION_${sufijoDoc}_${new Date().toISOString().substring(0,10)}`;
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

    let epigrafe = "";
    if (tipoAccion === "EXTINCION") {
      epigrafe = `“Por la cual se ordena la extinción, actualización y descuento de la Bonificación para la Asistencia Familiar a un señor ${data.grado || "patrullero"} en servicio activo de la Policía Nacional”`;
    } else {
      epigrafe = tipoResolucion === "PP"
        ? "“Por la cual se ordena el reconocimiento y actualización de la Bonificación para la Asistencia Familiar a un personal de Patrulleros de Policía en servicio activo de la Policía Nacional”"
        : "“Por la cual se ordena el reconocimiento y actualización de la Bonificación para la Asistencia Familiar a un personal del nivel ejecutivo en servicio activo de la Policía Nacional”";
    }
    
    pCenterBold(epigrafe + "\n");
    pCenterBold("EL JEFE NACIONAL DEL SERVICIO DE POLICÍA\n");
    
    if (tipoAccion === "EXTINCION") {
      const gradoTxt = data.grado || "Patrullero";
      const nombreTxt = datos.nombreUniformado|| "";
      const cedulaTxt = datos.cedulaUniformado|| "";
      const exConyugeTxt = data.nombreConyugeExtincion || "";
      const cedulaExConyugeTxt = data.cedulaConyugeExtincion || "";
      const resAntTxt = data.resolucionAnterior || "";
      const fResAntTxt = formatearFechaTexto(data.fechaResolucionAnterior);
      const comRadTxt = data.comunicadoRadicacion || "";
      const fRadicacionTxt = formatearFechaTexto(data.fechaRadicacion);
      const fHechoTxt = formatearFechaTexto(data.fechaHecho);
      const lugarSuscripcion = data.lugarSuscripcion || "";
      const nroActa = data.numeroActaSoporte || "";

      pJustified("En uso de las facultades conferidas por delegación, mediante Resolución nro. 0102 del 30 de octubre de 2024, y\n");
      pCenterBold("C O N S I D E R A N D O:\n");
      
      pJustified("Que el artículo 132 de la Ley 2179 del 30 de diciembre de 2021 “Por la cual se crea la categoría de Patrulleros de Policía, se establecen normas relacionadas con el régimen especial de carrera del personal uniformado de la Policía Nacional, se fortalece la profesionalización para el servicio público de policía y se dictan otras disposiciones.”, establece:\n");
      pJustified("“(…)\n\nARTÍCULO 132. BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. El personal del Nivel Ejecutivo y Patrulleros de Policía en servicio activo, tendrán derecho a partir del año siguiente a la entrada en vigencia de la presente ley, al reconocimiento y pago cada dos meses de una bonificación denominada \"para la asistencia familiar\" la cual se liquidará sobre la asignación básica del uniformado de la siguiente manera:\n\na) Un 30% por estado civil de casado o unión marital de hecho. No obstante, de presentarse la disolución de la unión marital de hecho, divorcio y/o cesación de efectos civiles se mantendrá dicho porcentaje de la asistencia cuando exista dependencia legal del hijo o hijos frutos de la unión. Se extenderá el beneficio de este literal, a los viudos con hijo o hijos habidos de la unión que generó el derecho.\n\nb) Por un primer hijo el 3% y 2% por el segundo, sin sobrepasar el 5%.\n\nEl Gobierno reglamentará dentro de los seis meses siguientes contados a partir de la entrada en vigencia de la presente ley lo referente al reconocimiento, periodicidad y extinción del emolumento.\n\nPARÁGRAFO. El reconocimiento y pago de la bonificación establecida en el presente artículo, es incompatible con el subsidio familiar establecido en el Decreto 1091 de 1995, establecido para el Nivel Ejecutivo.\n\n(…)”.\n");
      
      pJustified(`Que el ${gradoTxt.toLowerCase()}, relacionado en la parte resolutiva del presente acto administrativo, solicitó la extinción de la Bonificación para la Asistencia Familiar por las causales establecidas en el artículo 10 del Decreto 669 del 30 de abril de 2022 que adiciona el artículo 21C al Capítulo II del Título 1 del Decreto 1091 de 1995, el cual refiere:\n`);
      pJustified("“(…)\n\nARTÍCULO 21 C. EXTINCIÓN DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. La bonificación para la asistencia familiar se extingue por razón del cónyuge o compañero (a) permanente siempre que no hubiere hijos a cargo habidos dentro de la unión por los que exista el derecho a percibirlo, en los siguientes casos:\n\na) Por muerte del cónyuge o compañero (a) permanente.\nb) Por cesación de la vida conyugal en los siguientes casos:\n1. Por declaración judicial de nulidad o inexistencia del matrimonio.\n2. Por sentencia judicial de divorcio, válida en Colombia.\n3. Por separación judicial de cuerpos.\nc) Por terminación de la unión marital de hecho.\n\n(…)”.\n");
      
      pJustified("Que, el funcionario que se relaciona en la parte resolutiva del presente acto administrativo informó mediante comunicado oficial sobre la disolución de la unión marital de hecho, divorcio y/o cesación de efectos civiles.\n");
      
      pJustified("Que el artículo 11 del Decreto 669 del 30 de abril de 2022 que adiciona el artículo 21D al Capítulo II del Título 1 del Decreto 1091 de 1995, refiere lo siguiente:\n");
      pJustified("“(…)\n\nARTÍCULO 21 D. DESCUENTO DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. La extinción y la disminución de la bonificación para la asistencia familiar tendrán efectos a partir de la ocurrencia de cualquiera de las situaciones o hechos previstos en los Artículos 21B y 21C del presente Decreto.\n\nCuando ocurra cualquiera de los eventos que dan lugar a la extinción o disminución de la bonificación para la asistencia familiar, el integrante del Nivel Ejecutivo está en la obligación de dar el aviso correspondiente dentro de los noventa (90) días siguientes, si no lo hiciere, la Dirección General de la Policía Nacional ordenará el descuento de una suma igual a la que hubiere recibido en exceso.\n\n(…)”.\n");
      
      pJustified("Que la sociedad patrimonial entre compañeros permanentes se disuelve por los siguientes hechos previstos en el artículo 5 de la Ley 979 de 2005, “Por medio de la cual se modifica parcialmente la Ley 54 de 1990 y se establecen unos mecanismos ágiles para demostrar la unión marital de hecho y sus efectos patrimoniales entre compañeros permanentes”, así:\n");
      pJustified("“(…)\n\nArtículo 3º. El artículo 5º de la Ley 54 de 1990, quedará así:\n\nArtículo 5º. La sociedad patrimonial entre compañeros permanentes se disuelve por los siguientes hechos:\n\n1. Por mutuo consentimiento de los compañeros permanentes elevado a Escritura Pública ante Notario.\n2. De común acuerdo entre compañeros permanentes, mediante acta suscrita ante un Centro de Conciliación legalmente reconocido.\n3. Por Sentencia Judicial.\n4. Por la muerte de uno o ambos compañeros.\n\n(…)”.\n");
      
      // Párrafos con los datos inyectados de manera limpia y obligatoria
      pJustified(`Que al señor ${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${"datos.cedulaUniformado"}, le fue reconocida la bonificación de asistencia familiar mediante resolución nro. ${resAntTxt} del ${fResAntTxt}, por la unión marital de hecho demostrada con la señora ${exConyugeTxt}, identificada con cédula nro. ${cedulaExConyugeTxt}, en atención a la comunicación oficial nro. ${comRadTxt} del ${fRadicacionTxt}.\n`);
      
      pJustified(`Que el señor ${datos.grado} ${datos.nombreUniformado}, identificado con cédula de ciudadanía nro. ${datos.cedulaUniformado}, mediante comunicación oficial nro. ${comRadTxt} del ${fRadicacionTxt}, solicitó la extinción de la Bonificación para la Asistencia Familiar, por la terminación de la unión marital de hecho que sostenía con la señora ${exConyugeTxt}, identificada con cédula nro. ${cedulaExConyugeTxt}.\n`);
      
      pJustified(`Dicha disolución fue acreditada mediante acta de conciliación nro. ${nroActa} del ${fHechoTxt}, suscrita ante el Centro de Conciliación de la Policía Nacional sede ${lugarSuscripcion}, en aplicación de lo dispuesto en el numeral 3, literal c) del artículo 21C del Decreto 1091 de 1995, adicionado por el artículo 10 del Decreto 669 del 30 de abril de 2022.\n`);
      
      pJustified("Que revisada la documentación correspondiente por parte del Grupo de Talento Humano de la Jefatura Nacional del Servicio de Policía, se verificó que el servidor del Nivel Ejecutivo de la Policía Nacional, identificado en la parte resolutiva del presente acto administrativo, cumple con los requisitos exigidos para la extinción parcial y la actualización del reconocimiento de la Bonificación para la Asistencia Familiar.\n");

    } else {
      pJustified("En uso de las facultades conferidas por delegación, mediante resolución nro. 01182 del 04 de mayo de 2022, y\n");
      pCenterBold("CONSIDERANDO:\n");
      pJustified("Que el artículo 132 de la Ley 2179 del 30 de diciembre de 2021 “Por la cual se crea la categoría de Patrulleros de Policía, se establecen normas relacionadas con el régimen especial de carrera del personal uniformado de la Policía Nacional, se fortalece la profesionalización para el servicio público de policía y se dictan otras disposiciones.”, establece:");
      pJustified("“(…)\nARTÍCULO 132. BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. El personal del Nivel Ejecutivo y Patrulleros de Policía en servicio activo, tendrán derecho a partir del año siguiente a la entrada en vigencia de la presente ley, al reconocimiento y pago cada dos meses de una bonificación denominada \"para la asistencia familiar\"...\n(…)”.\n");
      pJustified("Que el personal que se relaciona en la parte resolutiva del presente acto administrativo, solicitó el reconocimiento de la Bonificación para la Asistencia Familiar a que consideran tener derecho.\n");
      pJustified("Que la unión marital de hecho entre compañeros permanentes, se declara por los mecanismos previstos en el artículo 2 de la Ley 979 de 2005.\n");
      pJustified("Que, mediante comunicaciones oficiales nro. GS-2026-000206-OFPLA del 13/01/2026 y GS-2026-006102-DITAH del 18/01/2026 se brindan los lineamientos para el trámite de Pago Pasivos Exigibles Vigencias Expiradas – PPEVE.\n");
    }
    
    pJustified("En mérito de lo expuesto,\n");
    pCenterBold("RESUELVE:\n");

    if (listaResuelve.length > 0) {
      if (tipoAccion === "EXTINCION") {
        pJustified(`ARTÍCULO 1°. ${listaResuelve[0]}\n`);
        pJustified(`ARTÍCULO 2°. Ordenar que los descuentos por concepto de Bonificación para la Asistencia del personal que se relacionan en el artículo nro.1° del presente acto administrativo se realicen en nómina del funcionario en cuotas de acuerdo a la capacidad de endeudamiento y sean reintegrados al presupuesto.\n`);
        pJustified(`ARTÍCULO 3°. Adelántese la notificación del presente acto administrativo de conformidad con los procedimientos previstos en la Ley 1437 de 2011 “Por la cual se expide el Código de Procedimiento Administrativo y de lo Contencioso Administrativo”.\n`);
        pJustified(`ARTÍCULO 4°. Contra la presente resolución proceden los recursos de reposición o apelación, los cuales deberán interponerse por escrito ante el señor Jefe Nacional del Servicio de Policía y Subdirector General de la Policía Nacional, respectivamente, en la diligencia de notificación personal, o dentro de los diez (10) días siguientes a ella, o a la notificación por aviso, o al vencimiento del término de publicación, según el caso, de conformidad con lo establecido en el artículo 76 de la Ley 1437 de 2011.\n`);
        pJustified(`ARTÍCULO 5°. La presente resolución rige a partir de la fecha de su expedición.\n`);
      } else {
        listaResuelve.forEach((texto, idx) => {
          pJustified(`ARTÍCULO ${idx + 1}°. ${texto}\n`);
        });
      }
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

  // UTILITARIOS Y DÍAS HÁBILES
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
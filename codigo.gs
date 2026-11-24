/**
 * SISTEMA DE GENERACIÓN DE RESOLUCIONES Y REGISTRO EN EXCEL
 * Policía Nacional de Colombia - Jefatura Nacional del Servicio de Policía
 */

const SPREADSHEET_ID = "1BvoAChx88VMGn2jRytV5XHNzGAvHvq6SgDefamxn5Wk";
const SHEET_NAME = "PROYECTO_RESOLUCIONES";
const FOLDER_ID = "1CTN0RhlbAdJbvy7C0CasXjnRWDJLXVFi";

const SIGLAS_GRADO_REC = {
  "Patrullero de Policía (PP)": "PP.",
  "Patrullero": "PT.",
  "Subintendente": "SI.",
  "Intendente": "IT.",
  "Intendente Jefe": "IJ.",
  "Subcomisario": "SC.",
  "Comisario": "CM."
};

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

    if (data.tipoAccion === "EXTINCION") {
      const reg = registrarExtincion(data);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", articulos: reg.articulos }))
                 .setMimeType(ContentService.MimeType.JSON);
    }

    const textoResuelve = procesarYGuardarRegistro(data);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", textoResuelve: textoResuelve }))
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
 * Construye el párrafo individual de cada funcionario en el ARTÍCULO 1° del RESUELVE
 * con cálculo exacto de 90 días hábiles Colombia, abreviatura de grado (PP o PT/SI/IT/IJ/SC/CM)
 * y concordancia de género.
 */
function construirParrafoResuelveReconocimiento(datos) {
  const fechaHecho = parseFechaLocal(datos.fechaHecho);
  const fechaRadicacion = parseFechaLocal(datos.fechaRadicacion);
  const diasHabilesContados = contarDiasHabilesColombia(fechaHecho, fechaRadicacion);
  const esDentroPlazo = diasHabilesContados <= 90;

  const fHechoTxt = formatearFechaTexto(datos.fechaHecho);
  const fRadicacionTxt = formatearFechaTexto(datos.fechaRadicacion);

  const esMujer = datos.generoUniformado === "Señora";
  const idUnif = esMujer ? "identificada" : "identificado";
  const sigla = datos.tipoResolucion === "PP"
    ? "PP."
    : (SIGLAS_GRADO_REC[datos.grado] || datos.grado);

  const textoPlazo = esDentroPlazo
    ? `Al haberse presentado dentro del término legal de noventa (90) días hábiles, los efectos fiscales se reconocerán a partir de la fecha del hecho que lo motiva, esto es, desde el ${fHechoTxt}.`
    : `Teniendo en cuenta que la solicitud fue presentada por fuera de los noventa (90) días hábiles, los efectos fiscales se reconocerán a partir de la fecha de su radicación, esto es, desde el ${fRadicacionTxt}, de conformidad con la norma aplicable.`;

  let textoResuelve = `${sigla} ${datos.nombreUniformado}, ${idUnif} con cédula de ciudadanía nro. ${datos.cedulaUniformado}, causa derecho al reconocimiento del `;

  if (datos.tipoPorcentaje === "30") {
    const esConyugeMujer = datos.generoConyuge === "Señora";
    const tratConyuge = esConyugeMujer ? "la señora" : "el señor";
    const idConyuge = esConyugeMujer ? "identificada" : "identificado";
    const motivo = datos.motivo30 || "haber constituido unión marital de hecho";
    const enteTxt = datos.enteEmisor ? ` del ${datos.enteEmisor}` : "";
    textoResuelve += `30 % de la asignación básica correspondiente a su grado, por ${motivo} con ${tratConyuge} ${datos.nombreConyuge}, ${idConyuge} con cédula de ciudadanía nro. ${datos.cedulaConyuge}, conforme al ${datos.tipoDocumento} nro. ${datos.numeroDocumento}${enteTxt} del ${fHechoTxt}. `;
  } else {
    const pctTexto = datos.tipoPorcentaje === "3" ? "3 %" : "2 %";
    const parentesco = datos.parentescoHijo || "su hija";
    textoResuelve += `${pctTexto} de la asignación básica correspondiente a su grado, por concepto de ${parentesco} ${datos.nombreHijo}, conforme al registro civil de nacimiento nro. ${datos.registroCivil}, que acredita su nacimiento el ${fHechoTxt}. `;
  }

  textoResuelve += `La solicitud fue radicada mediante el comunicado oficial ${datos.comunicadoRadicacion} del ${fRadicacionTxt}. ${textoPlazo}`;
  return { textoResuelve: limpiarRedaccion(textoResuelve), diasHabilesContados, esDentroPlazo };
}

/**
 * 1. REGISTRAR CADA SOLICITUD EN LA PESTAÑA 'PROYECTO_RESOLUCIONES'
 */
function procesarYGuardarRegistro(datos) {
  if (!datos) return "";

  const calc = construirParrafoResuelveReconocimiento(datos);

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
    calc.diasHabilesContados, 
    calc.esDentroPlazo ? "SÍ" : "NO", 
    calc.textoResuelve
  ]);

  return calc.textoResuelve;
}

/**
 * Calcula las concordancias gramaticales (singular / plural / género / grado / categoría)
 * para el epígrafe, considerandos y artículos cuando hay 1 o varios funcionarios.
 */
function obtenerConcordanciaLote(tipoResolucion, registros) {
  const esPP = tipoResolucion === "PP";
  const n = registros.length;

  if (n <= 1) {
    const r = registros[0] || {};
    const esMujer = r.generoUniformado === "Señora";
    const gradoRaw = String(r.grado || "Patrullero").trim();
    const gradoMin = esMujer
      ? gradoRaw.toLowerCase().replace(/^patrullero/, "patrullera")
      : gradoRaw.toLowerCase();

    if (esPP) {
      const cargoPP = esMujer ? "patrullera de policía" : "patrullero de policía";
      return {
        epigrafeSujeto: esMujer
          ? "a una señora patrullera de policía en servicio activo de la Policía Nacional"
          : "a un señor patrullero de policía en servicio activo de la Policía Nacional",
        considerandoSujeto: esMujer
          ? "Que la señora patrullera de policía de la Policía Nacional, que se relaciona en la parte resolutiva del presente acto administrativo, solicitó el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tiene derecho a ella."
          : "Que el señor patrullero de policía de la Policía Nacional, que se relaciona en la parte resolutiva del presente acto administrativo, solicitó el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tiene derecho a ella.",
        verificoSujeto: esMujer
          ? "la señora patrullera de policía de la Policía Nacional relacionada"
          : "el señor patrullero de policía de la Policía Nacional relacionado",
        art1Sujeto: esMujer ? "de la señora patrullera de policía" : "del señor patrullero de policía",
        art4Sujeto: esMujer ? "la señora patrullera de policía está" : "el señor patrullero de policía está"
      };
    } else {
      // Nivel Ejecutivo con el grado seleccionado
      return {
        epigrafeSujeto: esMujer
          ? `a una señora ${gradoMin} del Nivel Ejecutivo en servicio activo de la Policía Nacional`
          : `a un señor ${gradoMin} del Nivel Ejecutivo en servicio activo de la Policía Nacional`,
        considerandoSujeto: esMujer
          ? `Que la señora ${gradoMin} de la Policía Nacional, que se relaciona en la parte resolutiva del presente acto administrativo, solicitó el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tiene derecho a ella.`
          : `Que el señor ${gradoMin} de la Policía Nacional, que se relaciona en la parte resolutiva del presente acto administrativo, solicitó el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tiene derecho a ella.`,
        verificoSujeto: esMujer
          ? `la señora ${gradoMin} de la Policía Nacional relacionada`
          : `el señor ${gradoMin} de la Policía Nacional relacionado`,
        art1Sujeto: esMujer ? `de la señora ${gradoMin}` : `del señor ${gradoMin}`,
        art4Sujeto: esMujer ? "la integrante del Nivel Ejecutivo está" : "el integrante del Nivel Ejecutivo está"
      };
    }
  }

  // PLURAL (2 o más funcionarios en el lote)
  const todasMujeres = registros.every(r => r.generoUniformado === "Señora");
  const gradosUnicos = [...new Set(registros.map(r => String(r.grado || "").trim()))];

  const pluralizarGrado = (g, fem) => {
    const mapM = {
      "Patrullero": "patrulleros",
      "Subintendente": "subintendentes",
      "Intendente": "intendentes",
      "Intendente Jefe": "intendentes jefes",
      "Subcomisario": "subcomisarios",
      "Comisario": "comisarios"
    };
    const mapF = {
      "Patrullero": "patrulleras",
      "Subintendente": "subintendentes",
      "Intendente": "intendentes",
      "Intendente Jefe": "intendentes jefes",
      "Subcomisario": "subcomisarios",
      "Comisario": "comisarios"
    };
    return (fem ? mapF[g] : mapM[g]) || g.toLowerCase();
  };

  if (esPP) {
    return {
      epigrafeSujeto: todasMujeres
        ? "a unas señoras patrulleras de policía en servicio activo de la Policía Nacional"
        : "a un personal de Patrulleros de Policía en servicio activo de la Policía Nacional",
      considerandoSujeto: todasMujeres
        ? "Que las señoras patrulleras de policía de la Policía Nacional, que se relacionan en la parte resolutiva del presente acto administrativo, solicitaron el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tienen derecho a ella."
        : "Que los señores patrulleros de policía de la Policía Nacional, que se relacionan en la parte resolutiva del presente acto administrativo, solicitaron el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tienen derecho a ella.",
      verificoSujeto: todasMujeres
        ? "las señoras patrulleras de policía de la Policía Nacional relacionadas"
        : "los señores patrulleros de policía de la Policía Nacional relacionados",
      art1Sujeto: todasMujeres ? "de las señoras patrulleras de policía" : "de los señores patrulleros de policía",
      art4Sujeto: todasMujeres ? "las señoras patrulleras de policía están" : "los señores patrulleros de policía están"
    };
  } else {
    const descGrado = gradosUnicos.length === 1
      ? (todasMujeres ? `señoras ${pluralizarGrado(gradosUnicos[0], true)}` : `señores ${pluralizarGrado(gradosUnicos[0], false)}`)
      : (todasMujeres ? "señoras integrantes del Nivel Ejecutivo" : "señores integrantes del Nivel Ejecutivo");

    return {
      epigrafeSujeto: gradosUnicos.length === 1
        ? `a ${todasMujeres ? "unas" : "unos"} ${descGrado} del Nivel Ejecutivo en servicio activo de la Policía Nacional`
        : "a un personal del Nivel Ejecutivo en servicio activo de la Policía Nacional",
      considerandoSujeto: todasMujeres
        ? `Que las ${descGrado} de la Policía Nacional, que se relacionan en la parte resolutiva del presente acto administrativo, solicitaron el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tienen derecho a ella.`
        : `Que los ${descGrado} de la Policía Nacional, que se relacionan en la parte resolutiva del presente acto administrativo, solicitaron el reconocimiento de la Bonificación para la Asistencia Familiar, por considerar que tienen derecho a ella.`,
      verificoSujeto: todasMujeres
        ? `las ${descGrado} de la Policía Nacional relacionadas`
        : `los ${descGrado} de la Policía Nacional relacionados`,
      art1Sujeto: todasMujeres ? `de las ${descGrado}` : `de los ${descGrado}`,
      art4Sujeto: todasMujeres ? "las integrantes del Nivel Ejecutivo están" : "los integrantes del Nivel Ejecutivo están"
    };
  }
}

/**
 * 2. CREAR DOCUMENTO EN GOOGLE DOCS Y MOVER A CARPETA DESTINO
 *    (Soporta Decreto 668 de 2022 para PP y Decreto 669 de 2022 para Nivel Ejecutivo,
 *     con 1 o varios funcionarios en el ARTÍCULO 1.)
 */
function crearDocumentoResolucionCompleto(payload) {
  const data = payload || {};
  if (data.tipoAccion === "EXTINCION") return crearDocumentoExtincion(data);

  const tipoResolucion = data.tipoResolucion || "PP";
  const esPP = tipoResolucion === "PP";
  const listaResuelve = Array.isArray(data.listaResuelve) ? data.listaResuelve : [];
  const listaRegistros = Array.isArray(data.listaRegistros) ? data.listaRegistros : [];

  const conc = obtenerConcordanciaLote(tipoResolucion, listaRegistros);

  const nombreEtiqueta = esPP ? "Patrulleros_de_Policia" : "Nivel_Ejecutivo";
  const nombreDoc = `PROYECTO_RESOLUCION_${nombreEtiqueta}_${new Date().toISOString().substring(0, 10)}`;
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
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p.setFontFamily("Arial"); p.setFontSize(10); p.setBold(true); p.setSpacingAfter(8);
    return p;
  };

  const pJustified = (txt) => {
    const p = body.appendParagraph(txt);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setFontFamily("Arial"); p.setFontSize(10); p.setBold(false); p.setSpacingAfter(8);
    return p;
  };

  const pArticulo = (tituloNegrita, cuerpoTexto) => {
    const p = pJustified(tituloNegrita + cuerpoTexto);
    p.editAsText().setBold(0, tituloNegrita.length - 1, true);
    return p;
  };

  // Encabezado Institucional
  pCenterBold("MINISTERIO DE DEFENSA NACIONAL\nPOLICÍA NACIONAL\nJEFATURA NACIONAL DEL SERVICIO DE POLICÍA");
  pCenterBold(`RESOLUCIÓN NÚMERO          ${data.numeroResolucion || "0000"}          DEL ${data.fechaResolucion || "DD/MAAAA"}`);

  // Epígrafe dinámico según PP o Nivel Ejecutivo + grado + singular/plural
  const epigrafe = `“Por la cual se ordena el reconocimiento y actualización de la Bonificación para la Asistencia Familiar ${conc.epigrafeSujeto}”`;
  pCenterBold(epigrafe);

  pCenterBold("EL JEFE NACIONAL DEL SERVICIO DE POLICÍA");
  pJustified("En uso de las facultades conferidas por delegación, mediante Resolución nro. 01182 del 04 de mayo de 2022, y");
  pCenterBold("CONSIDERANDO:");

  // Ley 2179 de 2021 - Artículo 132
  pJustified("Que el artículo 132 de la Ley 2179 del 30 de diciembre de 2021 “Por la cual se crea la categoría de Patrulleros de Policía, se establecen normas relacionadas con el régimen especial de carrera del personal uniformado de la Policía Nacional, se fortalece la profesionalización para el servicio público de policía y se dictan otras disposiciones.”, establece:");
  pJustified(
    "“(…)\n" +
    "ARTÍCULO 132. BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. El personal del Nivel Ejecutivo y Patrulleros de Policía en servicio activo, tendrán derecho a partir del año siguiente a la entrada en vigencia de la presente ley, al reconocimiento y pago cada dos meses de una bonificación denominada \"para la asistencia familiar\" la cual se liquidará sobre la asignación básica del uniformado de la siguiente manera:\n" +
    "a) Un 30% por estado civil de casado o unión marital de hecho. No obstante, de presentarse la disolución de la unión marital de hecho, divorcio y/o cesación de efectos civiles se mantendrá dicho porcentaje de la asistencia cuando exista dependencia legal del hijo o hijos frutos de la unión. Se extenderá el beneficio de este literal, a los viudos con hijo o hijos habidos de la unión que generó el derecho.\n" +
    "b) Por un primer hijo el 3% y 2% por el segundo, sin sobrepasar el 5%.\n" +
    "El Gobierno reglamentará dentro de los seis meses siguientes contados a partir de la entrada en vigencia de la presente ley lo referente al reconocimiento, periodicidad y extinción del emolumento.\n" +
    "PARAGRAFO. El reconocimiento y pago de la bonificación establecida en el presente artículo, es incompatible con el subsidio familiar establecido en el Decreto 1091 de 1995, establecido para el Nivel Ejecutivo.\n" +
    "(…)”."
  );

  // Considerando de solicitud (dinámico según PP o Nivel Ejecutivo + grado + género/plural)
  pJustified(conc.considerandoSujeto);

  // Ley 979 de 2005 - Artículo 2
  pJustified("Que la unión marital de hecho entre compañeros permanentes, se declara por los mecanismos previstos en el artículo 2 de la Ley 979 de 2005, “Por medio de la cual se modifica parcialmente la Ley 54 de 1990 y se establecen unos mecanismos ágiles para demostrar la unión marital de hecho y sus efectos patrimoniales entre compañeros permanentes”, así:");
  pJustified(
    "“(…)\n" +
    "ARTÍCULO 2o. El artículo 4o. de la Ley 54 de 1990, quedará así:\n" +
    "Artículo 4o. La existencia de la unión marital de hecho entre compañeros permanentes, se declarará por cualquiera de los siguientes mecanismos:\n" +
    "Por escritura pública ante Notario por mutuo consentimiento de los compañeros permanentes.\n" +
    "Por Acta de Conciliación suscrita por los compañeros permanentes, en centro legalmente constituido.\n" +
    "Por sentencia judicial, mediante los medios ordinarios de prueba consagrados en el Código de Procedimiento Civil, con conocimiento de los Jueces de Familia de Primera Instancia.\n" +
    "(…)”."
  );

  // Bloque normativo: Decreto 668 de 2022 (PP) vs Decreto 669 de 2022 (Nivel Ejecutivo)
  if (esPP) {
    pJustified("Que, los artículos 21, 22, 23, 24, 25 y 26 del Decreto 668 del 30 de abril de 2022 “Por el cual se fijan los regímenes especiales en materia salarial, prestacional, pensional y de asignación de retiro para el personal de Patrulleros de Policía de la Policía Nacional y se dictan otras disposiciones”, establece:");
    pJustified(
      "“(…)\n" +
      "ARTÍCULO 21. Bonificación para la Asistencia Familiar. El Patrullero de Policía de la Policía Nacional, en servicio activo, tendrá derecho al reconocimiento y pago cada dos meses de una bonificación denominada \"para la asistencia familiar”, la cual se liquidará sobre la asignación básica del uniformado, así:\n" +
      "a) Casado o con unión marital de hecho el treinta por ciento (30%), más los porcentajes a que se tenga derecho conforme al literal c) de este Artículo.\n" +
      "b) Viudos, con hijos habidos dentro del matrimonio o la unión marital de hecho por los que exista el derecho a devengarlo, el treinta por ciento (30%), más los porcentajes de que trata el literal c) del presente Artículo.\n" +
      "c) Por el primer hijo el tres por ciento (3%) y un dos por ciento (2%) por el segundo hijo, sin que se sobrepase por este concepto del cinco por ciento (5%).\n" +
      "PARÁGRAFO 1. La solicitud de reconocimiento o aumento de la bonificación para la asistencia familiar deberá hacerse dentro de los noventa (90) días siguientes al hecho que la motive; las que se eleven con posterioridad al plazo antes fijado, tendrán efectos fiscales a partir de la fecha de su presentación.\n" +
      "PARÁGRAFO 2. Esta bonificación no será computable para efectos salariales, prestacionales, pensionales o de asignación de retiro.\n" +
      "ARTÍCULO 22. Disminución de la Bonificación para la Asistencia Familiar. Disminuye por razón de los hijos, así:\n" +
      "a) Por muerte.\n" +
      "b) Por matrimonio o constitución de la unión marital.\n" +
      "c) Por independencia económica.\n" +
      "d) Por haber llegado a la edad de dieciocho (18) años.\n" +
      "PARÁGRAFO. Se exceptúan de lo contemplado en el literal d) los hijos estudiantes menores de veinticuatro (24) años o los hijos inválidos, cuando dependan económicamente del Patrullero de Policía y mientras subsistan las condiciones de invalidez.\n" +
      "ARTÍCULO 23. Extinción de la bonificación para la asistencia familiar. La bonificación para la asistencia familiar se extingue por razón del cónyuge o compañero (a) permanente siempre que no hubiere hijos a cargo habidos dentro de la unión por los que exista el derecho a percibirlo, en los siguientes casos:\n" +
      "a) Por muerte del cónyuge o compañero (a) permanente.\n" +
      "b) Por cesación de la vida conyugal en los siguientes casos:\n" +
      "1. Por declaración judicial de nulidad o inexistencia del matrimonio.\n" +
      "2. Por sentencia judicial de divorcio, válida en Colombia.\n" +
      "3. Por separación judicial de cuerpos.\n" +
      "c) Por terminación de la unión marital de hecho.\n" +
      "ARTÍCULO 24. Descuento de la Bonificación para la Asistencia Familiar. La disminución o extinción de la bonificación para la asistencia familiar, tendrán efectos a partir de la ocurrencia de cualquiera de las situaciones o hechos previstos en los Artículos 22 y 23 del presente Decreto.\n" +
      "Cuando ocurra cualquiera de los eventos que dan lugar a la disminución o extinción de la bonificación para la asistencia familiar, el Patrullero de Policía está en la obligación de dar el aviso correspondiente dentro de los noventa (90) días siguientes; si no lo hiciere, la Dirección General de la Policía Nacional ordenará el descuento de una suma igual a la que hubiere recibido en exceso.\n" +
      "ARTÍCULO 25. Incompatibilidad de la Bonificación para la Asistencia Familiar con el Subsidio Familiar. La bonificación para la asistencia familiar consagrada en el presente Decreto, es incompatible con el subsidio familiar establecido para el personal de Oficiales, Nivel Ejecutivo, Suboficiales, Agentes, personal No Uniformado de la Policía Nacional y personal civil del Ministerio de Defensa Nacional. En tal sentido, en ningún caso habrá lugar al reconocimiento del subsidio familiar y la bonificación para la asistencia familiar por un mismo núcleo familiar.\n" +
      "Cuando el cónyuge o compañero (a) permanente del Patrullero de Policía ostente la calidad de Oficial, Nivel Ejecutivo, Suboficial, Agente, No Uniformado de la Policía Nacional o personal civil del Ministerio de Defensa Nacional, el reconocimiento de la bonificación para la asistencia familiar se hará siempre y cuando el Patrullero de Policía:\n" +
      "a) Reciba mayor asignación básica mensual.\n" +
      "b) Si la asignación básica fuere igual, se le reconocerá si acredita mayor tiempo de servicio.\n" +
      "PARÁGRAFO. El Patrullero de Policía de la Policía Nacional cuyo cónyuge o compañero (a) permanente preste sus servicios en el Ministerio de Defensa Nacional, para obtener derecho a la bonificación para la asistencia familiar, deberá acreditar que su cónyuge o compañero (a) permanente ha renunciado al subsidio familiar en la entidad en donde trabaja, mediante certificación expedida por esta última.\n" +
      "ARTÍCULO 26. Prohibición Pago Doble Bonificación para la Asistencia Familiar. En ningún caso habrá lugar al reconocimiento doble de la bonificación para la asistencia familiar. Cuando el cónyuge o compañero (a) permanente del Patrullero de Policía preste sus servicios en la Policía Nacional como integrante del Nivel Ejecutivo o de la categoría de Patrulleros de Policía, la bonificación para la asistencia familiar se reconocerá:\n" +
      "a) Al que reciba mayor asignación básica mensual.\n" +
      "b) Si la asignación básica fuere igual, recibirá el emolumento quien acredite mayor tiempo de servicio.\n" +
      "(…)”."
    );
  } else {
    // Decreto 669 del 30 de abril de 2022 - Nivel Ejecutivo (artículos 8 a 13 que adicionan los arts. 21A a 21F al Decreto 1091 de 1995)
    pJustified("Que, los artículos 8, 9, 10, 11, 12 y 13 del Decreto 669 del 30 de abril de 2022 “Por el cual se modifica el Decreto 1091 de 1995, que expide el Régimen de Asignaciones y Prestaciones para el personal del Nivel Ejecutivo de la Policía Nacional”, adicionaron los artículos 21A, 21B, 21C, 21D, 21E y 21F al Capítulo II del Título 1 del Decreto 1091 de 1995, estableciendo:");
    pJustified(
      "“(…)\n" +
      "ARTÍCULO 21A. BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. El integrante del Nivel Ejecutivo de la Policía Nacional, en servicio activo, tendrá derecho al reconocimiento y pago cada dos meses de una bonificación denominada \"para la asistencia familiar\", la cual se liquidará sobre la asignación básica del uniformado, así:\n" +
      "a) Casado o con unión marital de hecho el treinta por ciento (30%), más los porcentajes a que se tenga derecho conforme al literal c) de este Artículo.\n" +
      "b) Viudos, con hijos habidos dentro del matrimonio o la unión marital de hecho por los que exista el derecho a devengarlo, el treinta por ciento (30%), más los porcentajes de que trata el literal c) del presente Artículo.\n" +
      "c) Por el primer hijo el tres por ciento (3%) y un dos por ciento (2%) por el segundo hijo, sin que se sobrepase por este concepto del cinco por ciento (5%).\n" +
      "PARÁGRAFO 1. La solicitud de reconocimiento o aumento de la bonificación para la asistencia familiar deberá hacerse dentro de los noventa (90) días siguientes al hecho que la motive; las que se eleven con posterioridad al plazo antes fijado, tendrán efectos fiscales a partir de la fecha de su presentación.\n" +
      "PARÁGRAFO 2. Esta bonificación no será computable para efectos salariales, prestacionales, pensionales o de asignación de retiro.\n" +
      "ARTÍCULO 21B. DISMINUCIÓN DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. Disminuye por razón de los hijos, así:\n" +
      "a) Por muerte.\n" +
      "b) Por matrimonio o constitución de la unión marital.\n" +
      "c) Por independencia económica.\n" +
      "d) Por haber llegado a la edad de dieciocho (18) años.\n" +
      "PARÁGRAFO. Se exceptúan de lo contemplado en el literal d) los hijos estudiantes menores de veinticuatro (24) años o los hijos inválidos, cuando dependan económicamente del integrante del Nivel Ejecutivo y mientras subsistan las condiciones de invalidez.\n" +
      "ARTÍCULO 21C. EXTINCIÓN DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. La bonificación para la asistencia familiar se extingue por razón del cónyuge o compañero (a) permanente siempre que no hubiere hijos a cargo habidos dentro de la unión por los que exista el derecho a percibirlo, en los siguientes casos:\n" +
      "a) Por muerte del cónyuge o compañero (a) permanente.\n" +
      "b) Por cesación de la vida conyugal en los siguientes casos:\n" +
      "1. Por declaración judicial de nulidad o inexistencia del matrimonio.\n" +
      "2. Por sentencia judicial de divorcio, válida en Colombia.\n" +
      "3. Por separación judicial de cuerpos.\n" +
      "c) Por terminación de la unión marital de hecho.\n" +
      "ARTÍCULO 21D. DESCUENTO DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. La extinción y la disminución de la bonificación para la asistencia familiar tendrán efectos a partir de la ocurrencia de cualquiera de las situaciones o hechos previstos en los Artículos 21B y 21C del presente Decreto.\n" +
      "Cuando ocurra cualquiera de los eventos que dan lugar a la extinción o disminución de la bonificación para la asistencia familiar, el integrante del Nivel Ejecutivo está en la obligación de dar el aviso correspondiente dentro de los noventa (90) días siguientes, si no lo hiciere, la Dirección General de la Policía Nacional ordenará el descuento de una suma igual a la que hubiere recibido en exceso.\n" +
      "ARTÍCULO 21E. INCOMPATIBILIDAD DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR CON EL SUBSIDIO FAMILIAR. La bonificación para la asistencia familiar consagrada en el presente Decreto, es incompatible con el subsidio familiar establecido para el personal de Oficiales, Nivel Ejecutivo, Suboficiales, Agentes, personal No Uniformado de la Policía Nacional y personal civil del Ministerio de Defensa Nacional. En tal sentido, en ningún caso habrá lugar al reconocimiento del subsidio familiar y la bonificación para la asistencia familiar por un mismo núcleo familiar.\n" +
      "Cuando el cónyuge o compañero (a) permanente del integrante del Nivel Ejecutivo ostente la calidad de Oficial, Nivel Ejecutivo, Suboficial, Agente, No Uniformado de la Policía Nacional o personal civil del Ministerio de Defensa Nacional, el reconocimiento de la bonificación para la asistencia familiar se hará siempre y cuando el integrante del Nivel Ejecutivo:\n" +
      "a) Reciba mayor asignación básica mensual.\n" +
      "b) Si la asignación básica fuere igual, se le reconocerá si acredita mayor tiempo de servicio.\n" +
      "PARÁGRAFO. El integrante del Nivel Ejecutivo de la Policía Nacional cuyo cónyuge o compañero (a) permanente preste sus servicios en el Ministerio de Defensa Nacional, para obtener derecho a la bonificación para la asistencia familiar, deberá acreditar que su cónyuge o compañero (a) permanente ha renunciado al subsidio familiar en la entidad en donde trabaja, mediante certificación expedida por esta última.\n" +
      "ARTÍCULO 21F. PROHIBICIÓN PAGO DOBLE BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. En ningún caso habrá lugar al reconocimiento doble de la bonificación para la asistencia familiar. Cuando el cónyuge o compañero (a) permanente del integrante del Nivel Ejecutivo preste sus servicios en la Policía Nacional como integrante del Nivel Ejecutivo o de la categoría de Patrulleros de Policía, la bonificación para la asistencia familiar se reconocerá:\n" +
      "a) Al que reciba mayor asignación básica mensual.\n" +
      "b) Si la asignación básica fuere igual, recibirá el emolumento quien acredite mayor tiempo de servicio.\n" +
      "(…)”."
    );
  }

  // Párrafo de revisión documental (dinámico según categoría, grado, género y singular/plural)
  const verbCumple = listaRegistros.length > 1 ? "cumplen" : "cumple";
  pJustified(`Que, revisada la documentación correspondiente, por parte del Grupo de Talento Humano de la Jefatura Nacional del Servicio de Policía, se verificó que ${conc.verificoSujeto} en la parte resolutiva del presente acto administrativo, ${verbCumple} con los requisitos exigidos para el reconocimiento de la Bonificación para la Asistencia Familiar.`);

  pJustified("En mérito de lo expuesto,");
  pCenterBold("RESUELVE:");

  // ARTÍCULO 1 con cita normativa precisa según PP (Decreto 668) o Nivel Ejecutivo (Decreto 669)
  const citaNormaArt1 = esPP
    ? "el artículo 132 de la Ley 2179 del 30 de diciembre de 2021 y el artículo 21 del Decreto 668 del 30 de abril de 2022."
    : "el artículo 132 de la Ley 2179 del 30 de diciembre de 2021 y el artículo 8 del Decreto 669 del 30 de abril de 2022, que adicionó el artículo 21A al Decreto 1091 de 1995.";

  pArticulo(
    "ARTÍCULO 1. ",
    `Disponer el reconocimiento de la Bonificación por Asistencia Familiar a favor ${conc.art1Sujeto}, a partir del hecho que motive el reconocimiento; por concepto de hijos(as), compañeros(as) permanentes o cónyuges, de conformidad con lo previsto en ${citaNormaArt1}`
  );

  // Párrafos de cada funcionario cargado en el lote
  if (listaRegistros.length > 0) {
    listaRegistros.forEach((reg) => {
      const calc = construirParrafoResuelveReconocimiento(reg);
      pJustified(calc.textoResuelve);
    });
  } else if (listaResuelve.length > 0) {
    listaResuelve.forEach((texto) => {
      pJustified(texto);
    });
  }

  // Artículos 2° al 6° exactamente como el formato oficial
  pArticulo(
    "ARTÍCULO 2°. ",
    "Disponer que, en virtud del reconocimiento de la bonificación para la asistencia familiar, los valores liquidados sean cancelados con cargo a la vigencia actual para cada caso en concreto."
  );
  pArticulo(
    "ARTÍCULO 3°. ",
    "Adelántese la notificación del presente acto administrativo de conformidad con los procedimientos previstos en la Ley 1437 de 2011 “Por la cual se expide el Código de Procedimiento Administrativo y de lo Contencioso Administrativo”."
  );
  pArticulo(
    "ARTÍCULO 4°. ",
    `Cuando ocurra cualquiera de los eventos que dan lugar a la extinción o disminución de la bonificación para la asistencia familiar, ${conc.art4Sujeto} en la obligación de dar el aviso correspondiente dentro de los noventa (90) días siguientes mediante escrito, si no lo hiciere, la Dirección General de la Policía Nacional a través de la unidad nominadora ordenará el descuento de una suma igual a la que hubiere recibido en exceso.`
  );
  pArticulo(
    "ARTÍCULO 5°. ",
    "Contra la presente resolución procede el recurso de reposición y/o apelación que deberá ser interpuesto dentro de los diez (10) días hábiles siguientes a la fecha de su notificación, en los términos de los artículos 76 y 77 del Código de Procedimiento Administrativo y de lo Contencioso Administrativo."
  );
  pArticulo(
    "ARTÍCULO 6°. ",
    "La presente resolución rige a partir de la fecha de su expedición."
  );

  // Firmas y pie de página institucional
  pCenterBold("NOTIFÍQUESE Y CÚMPLASE:");
  pJustified("Dada en Bogotá D.C., a los");
  pCenterBold("\n\n\nMayor general RICARDO AUGUSTO ALARCÓN CAMPOS\nJefe Nacional del Servicio de Policía\n");

  const table = body.appendTable();
  const row1 = table.appendTableRow();
  row1.appendTableCell("Elaboró: SI. Fabian Sanguinetti\nJESEP-GUTAH.");
  row1.appendTableCell("Revisó: MY. Emma Paola Tique Bonilla\nJESEP-GUTAH");

  const row2 = table.appendTableRow();
  row2.appendTableCell("Revisó: MY. Beltran Ladino Jamer Alexander\nJESEP-ASJUR (E)");
  row2.appendTableCell("Revisó: CR. Hernando Alfredo Calderón Vega\nJESEP-SUSEP");

  const hoyTxt = Utilities.formatDate(new Date(), "America/Bogota", "dd-MM-yyyy");
  pJustified("\nFecha de elaboración: " + hoyTxt);
  pJustified("Ubicación: \\\\srvfilesponal3\\disec\\GUTAH\\DISEC - FABIAN SANGUINETTI\\2026\\4. NOMINA\\3. BONIFICACIONES\\10. OCTUBRE");
  pJustified("Carrera 59 26-21 CAN, Bogotá\nTeléfonos 515 9382\njesep.gutah@policia.gov.co\nwww.policia.gov.co");
  pCenterBold("\nINFORMACIÓN PÚBLICA");

  doc.saveAndClose();
  return { downloadUrl: doc.getUrl() };
}

// ============================================================
// 3. EXTINCIÓN DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR
//    (formato: formato_Extinción_Bonificación_GUIA.docx)
// ============================================================
const SHEET_EXTINCIONES = "PROYECTO_EXTINCIONES";

// Decreto que se cita según la categoría seleccionada en el formulario
const DECRETO_POR_CATEGORIA = {
  PP: "Decreto 668 del 30 de abril de 2022",
  EJECUTIVO: "Decreto 669 del 30 de abril de 2022"
};

// Documento con el que se acredita la disolución (Ley 979 de 2005, art. 5)
const TIPOS_DOCUMENTO_DISOLUCION = {
  "Acta de conciliación": { texto: "acta de conciliación", verbo: "suscrita ante" },
  "Escritura pública": { texto: "escritura pública", verbo: "otorgada ante" },
  "Sentencia judicial": { texto: "sentencia judicial", verbo: "proferida por" }
};

const EXT_ART132_INTRO = "Que el artículo 132 de la Ley 2179 del 30 de diciembre de 2021 “Por la cual se crea la categoría de Patrulleros de Policía, se establecen normas relacionadas con el régimen especial de carrera del personal uniformado de la Policía Nacional, se fortalece la profesionalización para el servicio público de policía y se dictan otras disposiciones.”, establece:";

const EXT_ART132 = [
  "“(…)",
  "ARTÍCULO 132. BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. El personal del Nivel Ejecutivo y Patrulleros de Policía en servicio activo, tendrán derecho a partir del año siguiente a la entrada en vigencia de la presente ley, al reconocimiento y pago cada dos meses de una bonificación denominada \"para la asistencia familiar\" la cual se liquidará sobre la asignación básica del uniformado de la siguiente manera:",
  "a) Un 30% por estado civil de casado o unión marital de hecho. No obstante, de presentarse la disolución de la unión marital de hecho, divorcio y/o cesación de efectos civiles se mantendrá dicho porcentaje de la asistencia cuando exista dependencia legal del hijo o hijos frutos de la unión. Se extenderá el beneficio de este literal, a los viudos con hijo o hijos habidos de la unión que generó el derecho.",
  "b) Por un primer hijo el 3% y 2% por el segundo, sin sobrepasar el 5%.",
  "El Gobierno reglamentará dentro de los seis meses siguientes contados a partir de la entrada en vigencia de la presente ley lo referente al reconocimiento, periodicidad y extinción del emolumento.",
  "PARAGRAFO. El reconocimiento y pago de la bonificación establecida en el presente artículo, es incompatible con el subsidio familiar establecido en el Decreto 1091 de 1995, establecido para el Nivel Ejecutivo.",
  "(…)”."
];

const EXT_ART21C = [
  "“(…)",
  "ARTÍCULO 21 C. EXTINCIÓN DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. La bonificación para la asistencia familiar se extingue por razón del cónyuge o compañero (a) permanente siempre que no hubiere hijos a cargo habidos dentro de la unión por los que exista el derecho a percibirlo, en los siguientes casos:",
  "a) Por muerte del cónyuge o compañero (a) permanente.",
  "b) Por cesación de la vida conyugal en los siguientes casos:",
  "1. Por declaración judicial de nulidad o inexistencia del matrimonio.",
  "2. Por sentencia judicial de divorcio, válida en Colombia.",
  "3. Por separación judicial de cuerpos.",
  "c) Por terminación de la unión marital de hecho.",
  "(…)”."
];

const EXT_ART21D = [
  "“(…)",
  "ARTÍCULO 21 D. DESCUENTO DE LA BONIFICACIÓN PARA LA ASISTENCIA FAMILIAR. La extinción y la disminución de la bonificación para la asistencia familiar tendrán efectos a partir de la ocurrencia de cualquiera de las situaciones o hechos previstos en los Artículos 21B y 21C del presente Decreto.",
  "Cuando ocurra cualquiera de los eventos que dan lugar a la extinción o disminución de la bonificación para la asistencia familiar, el integrante del Nivel Ejecutivo está en la obligación de dar el aviso correspondiente dentro de los noventa (90) días siguientes, si no lo hiciere, la Dirección General de la Policía Nacional ordenará el descuento de una suma igual a la que hubiere recibido en exceso.",
  "(…)”."
];

const EXT_LEY979_INTRO = "Que la sociedad patrimonial entre compañeros permanentes se disuelve por los siguientes hechos previstos en el artículo 5 de la Ley 979 de 2005, “Por medio de la cual se modifica parcialmente la Ley 54 de 1990 y se establecen unos mecanismos ágiles para demostrar la unión marital de hecho y sus efectos patrimoniales entre compañeros permanentes”, así:";

const EXT_LEY979 = [
  "“(…)",
  "Artículo 3º. El artículo 5º de la Ley 54 de 1990, quedará así:",
  "Artículo 5º. La sociedad patrimonial entre compañeros permanentes se disuelve por los siguientes hechos:",
  "1. Por mutuo consentimiento de los compañeros permanentes elevado a Escritura Pública ante Notario.",
  "2. De común acuerdo entre compañeros permanentes, mediante acta suscrita ante un Centro de Conciliación legalmente reconocido.",
  "3. Por Sentencia Judicial.",
  "4. Por la muerte de uno o ambos compañeros.",
  "(…)”."
];

const EXT_ARTICULOS_FINALES = [
  "Adelántese la notificación del presente acto administrativo de conformidad con los procedimientos previstos en la Ley 1437 de 2011 “Por la cual se expide el Código de Procedimiento Administrativo y de lo Contencioso Administrativo”.",
  "Contra la presente resolución procede el recurso de reposición y/o apelación que deberá ser interpuesto dentro de los diez (10) días hábiles siguientes a la fecha de su notificación, en los términos de los artículos 76 y 77 del Código de Procedimiento Administrativo y de lo Contencioso Administrativo.",
  "La presente resolución rige a partir de la fecha de su expedición."
];

function tratamientoExtincion(d) {
  const f = d.genero === "Señora";
  const pp = d.tipoResolucion === "PP";
  let cargo = pp ? "Patrullero de Policía (PP)" : String(d.grado || "").trim();
  if (f) cargo = cargo.replace(/^Patrullero/, "Patrullera");
  const fe = d.generoExPareja === "Señora";
  return {
    pp: pp,
    f: f,
    el: f ? "la" : "el",
    senor: f ? "señora" : "señor",
    al: f ? "a la señora" : "al señor",
    o: f ? "a" : "o",
    cargo: cargo,
    cargoLargo: cargo,
    exSenor: fe ? "la señora" : "el señor",
    exO: fe ? "a" : "o",
    exCompanero: fe ? "ex compañera permanente" : "ex compañero permanente"
  };
}

function construirTextosExtincion(d) {
  const t = tratamientoExtincion(d);
  const decreto = DECRETO_POR_CATEGORIA[d.tipoResolucion];
  const docInfo = TIPOS_DOCUMENTO_DISOLUCION[d.tipoDocumento];
  const f = fechaOficial;

  const nombre = d.nombreUniformado;
  const idUnif = `identificad${t.o} con cédula de ciudadanía nro. ${d.cedulaUniformado}`;
  const idEx = `identificad${t.exO} con cédula nro. ${d.cedulaExPareja}`;

  const epigrafe = `“Por la cual se ordena la extinción del reconocimiento y descuento de la Bonificación para la Asistencia Familiar a ${t.f ? "una señora" : "un señor"} ${t.cargo} en servicio activo de la Policía Nacional”.`;

  // Artículo 10 del Decreto 669 adiciona el art. 21C al Decreto 1091 de 1995 (Nivel Ejecutivo);
  // Artículo 23 del Decreto 668 establece la extinción para PP.
  const pCausales = t.pp
    ? `Que ${t.el} ${t.cargoLargo}, relacionad${t.o} en la parte resolutiva del presente acto administrativo, solicitó la extinción de la Bonificación para la Asistencia Familiar por las causales establecidas en el artículo 23 del ${decreto}, el cual refiere:`
    : `Que ${t.el} ${t.cargoLargo}, relacionad${t.o} en la parte resolutiva del presente acto administrativo, solicitó la extinción de la Bonificación para la Asistencia Familiar por las causales establecidas en el artículo 10 del ${decreto} que adiciona el artículo 21C al Capítulo II del Título 1 del Decreto 1091 de 1995, el cual refiere:`;

  const pArt21D = t.pp
    ? `Que el artículo 24 del ${decreto}, refiere lo siguiente:`
    : `Que el artículo 11 del ${decreto} que adiciona el artículo 21D al Capítulo II del Título 1 del Decreto 1091 de 1995, refiere lo siguiente:`;

  const pReconocida = `Que ${t.al} ${t.cargo} ${nombre}, ${idUnif}, le fue reconocida la bonificación de asistencia familiar mediante resolución nro. ${d.resolucionNro} del ${f(d.resolucionFecha)}, por la unión marital de hecho demostrada con ${t.exSenor} ${d.nombreExPareja}, ${idEx}, en atención a la comunicación oficial nro. ${d.comRecNro} del ${f(d.comRecFecha)}.`;

  const pSolicitud = `Que ${t.el} ${t.senor} ${t.cargo} ${nombre}, ${idUnif}, mediante comunicación oficial nro. ${d.comunicadoRadicacion} del ${f(d.fechaRadicacion)}, solicitó la extinción de la Bonificación para la Asistencia Familiar, por la terminación de la unión marital de hecho que sostenía con ${t.exSenor} ${d.nombreExPareja}, ${idEx}.`;

  const normaDisolucion = t.pp
    ? `en el literal c) del artículo 23 del ${decreto}.`
    : `en el literal c) del artículo 21C del Decreto 1091 de 1995, adicionado por el artículo 10 del ${decreto}.`;

  const pDisolucion = `Dicha disolución fue acreditada mediante ${docInfo.texto} nro. ${d.numeroDocumento} del ${f(d.fechaDocumento)}, ${docInfo.verbo} ${d.enteEmisor}, en aplicación de lo dispuesto ${normaDisolucion}`;

  const sujetoRevision = t.pp
    ? `${t.el} ${t.senor} ${t.cargoLargo}`
    : `${t.el} ${t.senor} ${t.cargoLargo} del Nivel Ejecutivo de la Policía Nacional`;
  const pRevisada = `Que revisada la documentación correspondiente por parte del Grupo de Talento Humano de la Jefatura Nacional del Servicio de Policía, se verificó que ${sujetoRevision}, identificad${t.o} en la parte resolutiva del presente acto administrativo, cumple con los requisitos exigidos para la extinción y la actualización del reconocimiento de la Bonificación para la Asistencia Familiar.`;

  const art1 = `ORDENAR la extinción de la Bonificación para la Asistencia Familiar reconocida ${t.al} ${t.cargo} ${nombre}, ${idUnif}, en lo correspondiente al porcentaje otorgado por razón de su ${t.exCompanero}, ${t.exSenor} ${d.nombreExPareja}, ${idEx}, con efectos a partir del ${f(d.fechaHecho)}.`;

  const art2 = `DISPONER el descuento de la suma de ${valorEnLetras(d.valorDescuento)} ($ ${formatearPesos(d.valorDescuento)}) M/CTE. valores reconocidos sin causar derecho a partir del ${f(d.descuentoDesde)} al ${f(d.descuentoHasta)}.`;

  const art3 = `ORDENAR al Grupo de Talento Humano de la Jefatura Nacional del Servicio de Policía adelantar las gestiones administrativas para el descuento periódico por nómina del valor pagado en exceso por concepto de Bonificación para la Asistencia Familiar ${t.al} ${t.cargo} ${nombre}, identificad${t.o} con cédula de ciudadanía No. ${d.cedulaUniformado}. Los descuentos se realizarán en cuotas mensuales ajustadas a su capacidad legal de pago conforme a la normatividad vigente, debiendo reintegrarse dichos recursos al rubro presupuestal de la unidad.`;

  const L = limpiarRedaccion;
  return {
    epigrafe: L(epigrafe),
    pCausales: L(pCausales),
    pArt21D: L(pArt21D),
    pReconocida: L(pReconocida),
    pSolicitud: L(pSolicitud),
    pDisolucion: L(pDisolucion),
    pRevisada: L(pRevisada),
    articulos: [L(art1), L(art2), L(art3)]
  };
}

function validarExtincion(d) {
  if (!d) throw new Error("No llegaron los datos de la extinción.");
  const requeridos = {
    tipoResolucion: "Categoría", genero: "Tratamiento del uniformado (señor/señora)", grado: "Grado",
    nombreUniformado: "Nombre del uniformado", cedulaUniformado: "Cédula del uniformado",
    resolucionNro: "Nro. de resolución de reconocimiento", resolucionFecha: "Fecha de resolución de reconocimiento",
    comRecNro: "Comunicación de la solicitud de reconocimiento", comRecFecha: "Fecha de la solicitud de reconocimiento",
    generoExPareja: "Tratamiento de la ex pareja", nombreExPareja: "Nombre de la ex pareja", cedulaExPareja: "Cédula de la ex pareja",
    comunicadoRadicacion: "Comunicado de la solicitud de extinción", fechaRadicacion: "Fecha de la solicitud de extinción",
    tipoDocumento: "Tipo de documento de disolución", numeroDocumento: "Nro. del documento de disolución",
    fechaDocumento: "Fecha del documento de disolución", enteEmisor: "Ente emisor",
    fechaHecho: "Fecha de efectos", valorDescuento: "Valor del descuento",
    descuentoDesde: "Descuento desde", descuentoHasta: "Descuento hasta"
  };
  const faltan = Object.keys(requeridos).filter(k => d[k] === undefined || d[k] === null || String(d[k]).trim() === "");
  if (faltan.length) throw new Error("Faltan datos de la extinción: " + faltan.map(k => requeridos[k]).join(", "));
  if (!DECRETO_POR_CATEGORIA[d.tipoResolucion]) throw new Error("Categoría no válida: " + d.tipoResolucion);
  if (!TIPOS_DOCUMENTO_DISOLUCION[d.tipoDocumento]) throw new Error("Tipo de documento no válido: " + d.tipoDocumento);
  if (!(Number(d.valorDescuento) > 0)) throw new Error("El valor del descuento debe ser mayor a cero.");
}

function registrarExtincion(datos) {
  validarExtincion(datos);
  const textos = construirTextosExtincion(datos);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_EXTINCIONES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EXTINCIONES);
    sheet.appendRow([
      "Fecha Registro", "Categoría", "Grado", "Nombre Uniformado", "Cédula",
      "Resolución Reconocimiento", "Fecha Resolución Reconocimiento", "Comunicado Solicitud Extinción", "Fecha Solicitud Extinción",
      "Ex Pareja", "Cédula Ex Pareja", "Documento Disolución", "Nro. Documento", "Fecha Documento", "Ente Emisor",
      "Efectos a partir de", "Valor Descuento", "Descuento Desde", "Descuento Hasta", "Artículos RESUELVE (1-3)"
    ]);
    sheet.getRange(1, 1, 1, 20).setFontWeight("bold").setBackground("#f4cccc");
  }

  sheet.appendRow([
    new Date(), datos.tipoResolucion, datos.grado, datos.nombreUniformado, datos.cedulaUniformado,
    datos.resolucionNro, datos.resolucionFecha, datos.comunicadoRadicacion, datos.fechaRadicacion,
    datos.nombreExPareja, datos.cedulaExPareja, datos.tipoDocumento, datos.numeroDocumento, datos.fechaDocumento, datos.enteEmisor,
    datos.fechaHecho, Number(datos.valorDescuento), datos.descuentoDesde, datos.descuentoHasta,
    textos.articulos.map((a, i) => `ARTÍCULO ${i + 1}°. ${a}`).join("\n\n")
  ]);

  return { articulos: textos.articulos };
}

function crearDocumentoExtincion(payload) {
  const d = payload.extincion;
  validarExtincion(d);
  const t = construirTextosExtincion(d);

  const hoyISO = formatearFechaISO(new Date());
  const doc = DocumentApp.create(`PROYECTO_RESOLUCION_EXTINCION_${d.tipoResolucion}_${hoyISO}`);

  try {
    DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(FOLDER_ID));
  } catch (e) {
    Logger.log("Aviso Drive: " + e.toString());
  }

  const body = doc.getBody();
  body.setMarginTop(72); body.setMarginBottom(72); body.setMarginLeft(72); body.setMarginRight(72);

  const pCenterBold = (txt) => {
    const p = body.appendParagraph(txt);
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p.setFontFamily("Arial"); p.setFontSize(10); p.setBold(true); p.setSpacingAfter(10);
    return p;
  };
  const pJustified = (txt) => {
    const p = body.appendParagraph(txt);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setFontFamily("Arial"); p.setFontSize(10); p.setBold(false); p.setSpacingAfter(10);
    return p;
  };
  const bloque = (lineas) => lineas.forEach(l => pJustified(l));

  pCenterBold("MINISTERIO DE DEFENSA NACIONAL\nPOLICÍA NACIONAL\nJEFATURA NACIONAL DEL SERVICIO DE POLICÍA");
  pCenterBold(`RESOLUCIÓN NÚMERO      ${payload.numeroResolucion || "0000"}      DEL ${payload.fechaResolucion || "DD/MM/AAAA"}`);
  pCenterBold(t.epigrafe);
  pCenterBold("EL JEFE NACIONAL DEL SERVICIO DE POLICÍA");
  pJustified("En uso de las facultades conferidas por delegación, mediante Resolución nro. 01182 del 04 de mayo de 2022, y");
  pCenterBold("C O N S I D E R A N D O:");

  pJustified(EXT_ART132_INTRO);
  bloque(EXT_ART132);

  pJustified(t.pCausales);
  bloque(EXT_ART21C);

  pJustified(t.pArt21D);
  bloque(EXT_ART21D);

  pJustified(EXT_LEY979_INTRO);
  bloque(EXT_LEY979);

  pJustified(t.pReconocida);
  pJustified(t.pSolicitud);
  pJustified(t.pDisolucion);
  pJustified(t.pRevisada);

  pJustified("En mérito de lo expuesto, el suscrito Jefe Nacional del Servicio de Policía en uso de sus facultades legales,");
  pCenterBold("R E S U E L V E:");

  t.articulos.forEach((texto, i) => {
    const prefijo = `ARTÍCULO ${i + 1}°. `;
    const p = pJustified(prefijo + texto);
    p.editAsText().setBold(0, prefijo.length - 1, true);
  });
  EXT_ARTICULOS_FINALES.forEach((texto, i) => {
    const prefijo = `ARTÍCULO ${t.articulos.length + i + 1}°. `;
    const p = pJustified(prefijo + texto);
    p.editAsText().setBold(0, prefijo.length - 1, true);
  });

  pCenterBold("NOTIFÍQUESE Y CÚMPLASE:");
  pJustified("Dada en Bogotá D.C., a los");
  pCenterBold("\n\n\nMayor general RICARDO AUGUSTO ALARCÓN CAMPOS\nJefe Nacional del Servicio de Policía");

  const table = body.appendTable();
  const row1 = table.appendTableRow();
  row1.appendTableCell("Elaboró: SI. Fabian Sanguinetti\nJESEP-GUTAH");
  row1.appendTableCell("Revisó: MY. Emma Paola Tique Bonilla\nJESEP-GUTAH");
  const row2 = table.appendTableRow();
  row2.appendTableCell("Revisó: MY. Jamer Alexander Beltran Ladino\nJESEP-ASJUR (E)");
  row2.appendTableCell("Revisó: CR. Hernando Alfredo Calderón Vega\nJESEP-SUSEP");

  const hoyTxt = Utilities.formatDate(new Date(), "America/Bogota", "dd-MM-yyyy");
  pJustified("\nFecha de elaboración: " + hoyTxt);
  pJustified("Ubicación: \\\\srvfilesponal3\\disec\\GUTAH\\DISEC - JOHN ALEXANDER BRAVO JIMENEZ\\2026\\4. NOMINA\\3. BONIFICACIONES\\3. EXTINCIÓN");
  pJustified("Carrera 59 26-21 CAN, Bogotá - Teléfonos 515 9382\njesep.gutah@policia.gov.co - www.policia.gov.co");
  pCenterBold("\nINFORMACIÓN PÚBLICA");

  doc.saveAndClose();
  return { downloadUrl: doc.getUrl() };
}

// ---- Utilidades de formato ----

function limpiarRedaccion(s) {
  s = String(s).replace(/\s+/g, " ").trim();
  s = s.replace(/\s+([,;:.)])/g, "$1").replace(/\(\s+/g, "(");
  s = s.replace(/([,;:])(?=[^\s\d])/g, "$1 ");
  s = s.replace(/,{2,}/g, ",").replace(/;{2,}/g, ";").replace(/\.{2,}(?!\.)/g, ".");
  s = s.replace(/[,;:]+\.$/, ".");
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function fechaOficial(iso) {
  if (!iso || !String(iso).includes("-")) return iso || "";
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const p = String(iso).split("-");
  return `${p[2].padStart(2, "0")} de ${meses[parseInt(p[1], 10) - 1]} de ${p[0]}`;
}

function formatearPesos(valor) {
  const total = Math.round(Number(valor) * 100);
  const pesos = Math.floor(total / 100);
  const cent = String(total % 100).padStart(2, "0");
  return String(pesos).replace(/\B(?=(\d{3})+(?!\d))/g, ".") + "," + cent;
}

function valorEnLetras(valor) {
  const total = Math.round(Number(valor) * 100);
  const pesos = Math.floor(total / 100);
  const cent = total % 100;
  let txt = numeroALetras(pesos).toUpperCase();
  if (pesos >= 1000000 && pesos % 1000000 === 0) txt += " DE PESOS";
  else txt += pesos === 1 ? " PESO" : " PESOS";
  if (cent) txt += " CON " + numeroALetras(cent).toUpperCase() + (cent === 1 ? " CENTAVO" : " CENTAVOS");
  return txt;
}

function numeroALetras(n) {
  const U = ["cero", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve", "veinte", "veintiuno", "veintidós", "veintitrés", "veinticuatro", "veinticinco", "veintiséis", "veintisiete", "veintiocho", "veintinueve"];
  const D = ["", "", "", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"];
  const C = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];
  const apocope = (s) => s.replace(/veintiuno$/, "veintiún").replace(/uno$/, "un");
  const menorMil = (x) => {
    if (x === 100) return "cien";
    const partes = [];
    const c = Math.floor(x / 100), r = x % 100;
    if (c) partes.push(C[c]);
    if (r) {
      if (r < 30) partes.push(U[r]);
      else partes.push(r % 10 ? `${D[Math.floor(r / 10)]} y ${U[r % 10]}` : D[Math.floor(r / 10)]);
    }
    return partes.join(" ");
  };

  n = Math.floor(Number(n));
  if (n === 0) return "cero";
  if (n >= 1000000000) return String(n);

  const millones = Math.floor(n / 1000000);
  const miles = Math.floor((n % 1000000) / 1000);
  const resto = n % 1000;
  const partes = [];
  if (millones) partes.push(millones === 1 ? "un millón" : apocope(menorMil(millones)) + " millones");
  if (miles) partes.push(miles === 1 ? "mil" : apocope(menorMil(miles)) + " mil");
  if (resto) partes.push(menorMil(resto));
  return apocope(partes.join(" "));
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
  return `${parseInt(p[2], 10)} de ${meses[parseInt(p[1], 10) - 1]} de ${p[0]}`;
}
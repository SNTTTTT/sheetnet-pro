/**
 * ============================================================================
 * SOLU · Capital Humano — Evaluación DISC CLEAVER
 * ============================================================================
 */

// ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
const CFG_APP = {
  SS_ID          : "17guNUaSg2bHQ9vd2HiS0f02YK92R-cl1rywFWFd1JN0",
  HOJA_TOKENS    : "Tokens_Candidatos",
  HOJA_DOMINIOS  : "Dominios_Permitidos",
  HOJA_RESPUESTAS: "Base_Cleaver_CANDIDATOS"
};

function _ss() { return SpreadsheetApp.openById(CFG_APP.SS_ID); }

// ── BLOQUES CLEAVER (24 × 4 adjetivos) ───────────────────────────────────────
const BLOQUES_CLEAVER = [
  {id:1,  adj:["PERSUASIVO","GENTIL","HUMILDE","ORIGINAL"]},
  {id:2,  adj:["AGRESIVO","ALMA DE LA FIESTA","COMODINO","TEMEROSO"]},
  {id:3,  adj:["AGRADABLE","TEMEROSO DE DIOS","TENAZ","ATRACTIVO"]},
  {id:4,  adj:["CAUTELOSO","DETERMINADO","CONVINCENTE","BONACHON"]},
  {id:5,  adj:["DOCIL","ATREVIDO","LEAL","ENCANTADOR"]},
  {id:6,  adj:["DISPUESTO","DESEOSO","CONSECUENTE","ENTUSIASTA"]},
  {id:7,  adj:["FUERZA DE VOLUNTAD","MENTE ABIERTA","COMPLACIENTE","ANIMOSO"]},
  {id:8,  adj:["CONFIADO","SIMPATIZADOR","TOLERANTE","AFIRMATIVO"]},
  {id:9,  adj:["ECUANIME","PRECISO","NERVIOSO","JOVIAL"]},
  {id:10, adj:["DISCIPLINADO","GENEROSO","ANIMOSO","PERSISTENTE"]},
  {id:11, adj:["COMPETITIVO","ALEGRE","CONSIDERADO","ARMONIOSO"]},
  {id:12, adj:["ADMIRABLE","BONDADOSO","RESIGNADO","CARACTER FIRME"]},
  {id:13, adj:["OBEDIENTE","QUISQUILLOSO","INCONQUISTABLE","JUGUETON"]},
  {id:14, adj:["RESPETUOSO","EMPRENDEDOR","OPTIMISTA","SERVICIAL"]},
  {id:15, adj:["VALIENTE","INSPIRADOR","SUMISO","TIMIDO"]},
  {id:16, adj:["ADAPTABLE","DISPUTADOR","INDIFERENTE","SANGRE LIVIANA"]},
  {id:17, adj:["AMIGUERO","PACIENTE","CONFIANZA EN SI MISMO","MESURADO PARA HABLAR"]},
  {id:18, adj:["CONFORME","CONFIABLE","PACIFICO","POSITIVO"]},
  {id:19, adj:["AVENTURERO","RECEPTIVO","CORDIAL","MODERADO"]},
  {id:20, adj:["INDULGENTE","ESTETA","VIGOROSO","SOCIABLE"]},
  {id:21, adj:["PARLANCHIN","CONTROLADO","CONVENCIONAL","DECISIVO"]},
  {id:22, adj:["COHIBIDO","EXACTO","FRANCO","BUEN COMPAÑERO"]},
  {id:23, adj:["DIPLOMATICO","AUDAZ","REFINADO","SATISFECHO"]},
  {id:24, adj:["INQUIETO","POPULAR","BUEN VECINO","DEVOTO"]}
];

// ── ENTRY POINT ───────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('SOLU · Evaluación DISC — CLEAVER')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1.0');
}

// ── AUTH: PLANTILLA ───────────────────────────────────────────────────────────
function validarCorreoCorporativo(correo) {
  try {
    correo = String(correo).trim().toLowerCase();
    if (!correo.includes('@')) return {ok: false, msg: "Correo inválido."};

    const dominio = correo.split('@')[1];
    const ss   = _ss();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_DOMINIOS);
    if (!hoja) return {ok: false, msg: "Hoja 'Dominios_Permitidos' no encontrada. Ejecute inicializarHojasApp()."};

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      // normalizar: quitar @ si está guardado como @dominio.com
      const stored = String(datos[i][0]).trim().toLowerCase().replace(/^@/, '');
      if (stored === dominio) {
        if (_yaRespondio(correo)) return {ok: false, msg: "Este correo ya completó la evaluación.", repetido: true};
        return {ok: true, tipo: "PLANTILLA", correo: correo, nombre: "", puesto: ""};
      }
    }
    return {ok: false, msg: "El dominio de este correo no está autorizado."};
  } catch(e) {
    Logger.log("Error validarCorreoCorporativo: " + e);
    return {ok: false, msg: "Error interno: " + e.toString()};
  }
}

// ── AUTH: CANDIDATO ───────────────────────────────────────────────────────────
function validarToken(token) {
  try {
    token = String(token).trim().toUpperCase();
    const ss   = _ss();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hoja) return {ok: false, msg: "Hoja 'Tokens_Candidatos' no encontrada. Ejecute inicializarHojasApp()."};

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toUpperCase() === token) {
        if (String(datos[i][4]).trim().toUpperCase() === "SI") {
          return {ok: false, msg: "Este token ya fue utilizado.", repetido: true};
        }
        return {
          ok    : true,
          tipo  : "CANDIDATO",
          correo: String(datos[i][1]).trim(),
          nombre: String(datos[i][2]).trim(),
          puesto: String(datos[i][3]).trim(),
          fila  : i + 1
        };
      }
    }
    return {ok: false, msg: "Token no encontrado. Verifique el código proporcionado."};
  } catch(e) {
    Logger.log("Error validarToken: " + e);
    return {ok: false, msg: "Error interno: " + e.toString()};
  }
}

// ── GUARDAR RESPUESTAS ────────────────────────────────────────────────────────
function guardarRespuestas(payload) {
  try {
    const ss   = _ss();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_RESPUESTAS);
    if (!hoja) return {ok: false, msg: "Hoja 'Base_Cleaver_CANDIDATOS' no encontrada."};

    if (!payload.correo || !payload.respuestas || payload.respuestas.length !== 24) {
      return {ok: false, msg: "Datos incompletos. No se guardó la evaluación."};
    }

    const fecha = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm:ss");
    let fila = [fecha, payload.correo, payload.nombre, payload.puesto];
    payload.respuestas.forEach(r => { fila.push(r.mas); fila.push(r.menos); });

    if (hoja.getLastRow() === 0) {
      let heads = ["Marca temporal", "Correo", "Nombre completo", "Área / Puesto"];
      for (let b = 1; b <= 24; b++) { heads.push("Bloque " + b + " MÁS"); heads.push("Bloque " + b + " MENOS"); }
      hoja.appendRow(heads);
    }

    hoja.appendRow(fila);
    SpreadsheetApp.flush();

    if (payload.tipo === "CANDIDATO" && payload.tokenFila) {
      const hTok = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
      if (hTok) {
        hTok.getRange(payload.tokenFila, 5).setValue("SI");
        hTok.getRange(payload.tokenFila, 6).setValue(fecha);
        SpreadsheetApp.flush();
      }
    }

    Logger.log("✅ Guardado: " + payload.correo + " (" + payload.tipo + ")");
    return {ok: true, msg: "Evaluación guardada correctamente."};

  } catch(e) {
    Logger.log("Error guardarRespuestas: " + e);
    return {ok: false, msg: "Error al guardar: " + e.toString()};
  }
}

// ── OBTENER BLOQUES ───────────────────────────────────────────────────────────
function obtenerBloques() {
  return BLOQUES_CLEAVER;
}

// ── UTIL: ya respondió ────────────────────────────────────────────────────────
function _yaRespondio(correo) {
  try {
    const ss   = _ss();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_RESPUESTAS);
    if (!hoja || hoja.getLastRow() < 2) return false;
    const datos = hoja.getRange(2, 2, hoja.getLastRow() - 1, 1).getValues();
    return datos.some(r => String(r[0]).trim().toLowerCase() === correo);
  } catch(e) { return false; }
}

// ── ADMIN: inicializar hojas (ejecutar UNA VEZ) ───────────────────────────────
function inicializarHojasApp() {
  const ss = _ss();

  if (!ss.getSheetByName(CFG_APP.HOJA_TOKENS)) {
    const h = ss.insertSheet(CFG_APP.HOJA_TOKENS);
    h.getRange(1,1,1,6).setValues([["Token","Correo","Nombre","Puesto","Usado","Fecha_Uso"]]);
    h.getRange(1,1,1,6).setBackground("#0d5a7a").setFontColor("white").setFontWeight("bold");
  }

  if (!ss.getSheetByName(CFG_APP.HOJA_DOMINIOS)) {
    const h = ss.insertSheet(CFG_APP.HOJA_DOMINIOS);
    h.getRange(1,1,1,2).setValues([["Dominio","Descripción"]]);
    h.getRange(1,1,1,2).setBackground("#0d5a7a").setFontColor("white").setFontWeight("bold");
  }

  if (!ss.getSheetByName(CFG_APP.HOJA_RESPUESTAS)) {
    const h = ss.insertSheet(CFG_APP.HOJA_RESPUESTAS);
    h.getRange(1,1,1,1).setValues([["(Se crea encabezado automático al guardar primera respuesta)"]]);
  }

  SpreadsheetApp.getUi().alert(
    "✅ Hojas creadas:\n" +
    "• Tokens_Candidatos\n" +
    "• Dominios_Permitidos\n" +
    "• Base_Cleaver_CANDIDATOS\n\n" +
    "Siguiente paso: agrega dominios en 'Dominios_Permitidos' y publica el webapp."
  );
}

// ── ADMIN: sincronizar tokens desde Nuevos_Ingresos ──────────────────────────
function sincronizarTokensDesdeNuevosIngresos() {
  try {
    const ss   = _ss();
    const hNI  = ss.getSheetByName("Nuevos_Ingresos");
    const hTok = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hNI || !hTok) { SpreadsheetApp.getUi().alert("Hojas no encontradas."); return; }

    const dataNI  = hNI.getDataRange().getValues();
    const dataTok = hTok.getDataRange().getValues();
    const existentes = dataTok.slice(1).map(r => String(r[0]).trim().toUpperCase());

    const heads     = dataNI[0].map(h => String(h).trim().toUpperCase());
    const colID     = heads.findIndex(h => h.includes("ID_PROCESO"));
    const colNombre = heads.findIndex(h => h === "NOMBRE");
    const colCorreo = heads.findIndex(h => h === "CORREO");
    const colPuesto = heads.findIndex(h => h.includes("PUESTO"));

    let nuevos = 0;
    for (let i = 1; i < dataNI.length; i++) {
      const id = String(dataNI[i][colID] || "").trim().toUpperCase();
      if (!id || existentes.includes(id)) continue;
      hTok.appendRow([
        id,
        colCorreo > -1 ? dataNI[i][colCorreo] : "",
        colNombre > -1 ? dataNI[i][colNombre] : "",
        colPuesto > -1 ? dataNI[i][colPuesto] : "",
        "NO", ""
      ]);
      nuevos++;
    }

    SpreadsheetApp.getUi().alert("✅ " + nuevos + " token(s) nuevos sincronizados.");
  } catch(e) {
    Logger.log("Error sincronizarTokens: " + e);
  }
}

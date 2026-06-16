/**
 * ============================================================================
 * CLEAVER_APLICACION.gs — MÓDULO DE APLICACIÓN DE PRUEBA DISC
 * ============================================================================
 * Sistema: SOLU · Capital Humano
 * Flujo PLANTILLA : correo corporativo → whitelist dominios → prueba → Sheets
 * Flujo CANDIDATO : token único (generado en Nuevos_Ingresos) → prueba → Sheets
 * Guardado directo en Base_Cleaver_CANDIDATOS (compatible con pipeline v3.1)
 * ============================================================================
 */

// ── HOJAS ADICIONALES (se suman al CONFIG existente) ─────────────────────────
const CFG_APP = {
  HOJA_TOKENS    : "Tokens_Candidatos",    // token | correo | nombre | puesto | usado | fecha_uso
  HOJA_DOMINIOS  : "Dominios_Permitidos",  // dominio (ej: empresa.com)
  HOJA_RESPUESTAS: "Base_Cleaver_CANDIDATOS", // misma hoja del pipeline
  URL_APP        : ScriptApp.getService().getUrl()
};

// ── BLOQUES CLEAVER (24 × 4 adjetivos, orden visual exacto) ──────────────────
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

// ============================================================================
// AUTH: PLANTILLA — valida correo corporativo contra dominios permitidos
// ============================================================================
function validarCorreoCorporativo(correo) {
  try {
    correo = String(correo).trim().toLowerCase();
    if (!correo.includes('@')) return {ok: false, msg: "Correo inválido."};

    const dominio = correo.split('@')[1];
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_DOMINIOS);
    if (!hoja) return {ok: false, msg: "Configuración de dominios no encontrada."};

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toLowerCase() === dominio) {
        const yaRespondio = _yaRespondio(correo);
        if (yaRespondio) return {ok: false, msg: "Este correo ya completó la evaluación.", repetido: true};
        return {ok: true, tipo: "PLANTILLA", correo: correo, nombre: "", puesto: ""};
      }
    }
    return {ok: false, msg: "El dominio de este correo no está autorizado."};
  } catch(e) {
    Logger.log("Error validarCorreoCorporativo: " + e);
    return {ok: false, msg: "Error interno. Contacte a Capital Humano."};
  }
}

// ============================================================================
// AUTH: CANDIDATO — valida token único desde Tokens_Candidatos
// ============================================================================
function validarToken(token) {
  try {
    token = String(token).trim().toUpperCase();
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hoja) return {ok: false, msg: "Sistema de tokens no configurado."};

    const datos = hoja.getDataRange().getValues();
    // Cols: 0=Token 1=Correo 2=Nombre 3=Puesto 4=Usado(SI/NO) 5=Fecha_Uso
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
          fila  : i + 1   // para marcar como usado al guardar
        };
      }
    }
    return {ok: false, msg: "Token no encontrado. Verifique el código proporcionado."};
  } catch(e) {
    Logger.log("Error validarToken: " + e);
    return {ok: false, msg: "Error interno. Contacte a Capital Humano."};
  }
}

// ============================================================================
// GUARDAR RESPUESTAS → Base_Cleaver_CANDIDATOS
// ============================================================================
function guardarRespuestas(payload) {
  /*
   * payload = {
   *   tipo     : "PLANTILLA" | "CANDIDATO",
   *   correo   : String,
   *   nombre   : String,
   *   puesto   : String,
   *   tokenFila: Number (solo CANDIDATO),
   *   respuestas: [ {bloque:1, mas:"PERSUASIVO", menos:"HUMILDE"}, ... ]  // 24 items
   * }
   */
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_RESPUESTAS);
    if (!hoja) return {ok: false, msg: "Hoja de respuestas no encontrada."};

    if (!payload.correo || !payload.respuestas || payload.respuestas.length !== 24) {
      return {ok: false, msg: "Datos incompletos. No se guardó la evaluación."};
    }

    const fecha = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm:ss");
    let fila = [fecha, payload.correo, payload.nombre, payload.puesto];
    payload.respuestas.forEach(r => { fila.push(r.mas); fila.push(r.menos); });

    if (hoja.getLastRow() === 0) {
      let heads = ["Marca temporal","Dirección de correo electrónico","Nombre completo","Área / Puesto"];
      for (let b = 1; b <= 24; b++) {
        heads.push(`Bloque ${b} MÁS`);
        heads.push(`Bloque ${b} MENOS`);
      }
      hoja.appendRow(heads);
    }

    hoja.appendRow(fila);
    SpreadsheetApp.flush();

    if (payload.tipo === "CANDIDATO" && payload.tokenFila) {
      const hTokens = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
      if (hTokens) {
        hTokens.getRange(payload.tokenFila, 5).setValue("SI");
        hTokens.getRange(payload.tokenFila, 6).setValue(fecha);
        SpreadsheetApp.flush();
      }
    }

    Logger.log("✅ Respuestas guardadas: " + payload.correo + " (" + payload.tipo + ")");
    return {ok: true, msg: "Evaluación guardada correctamente."};

  } catch(e) {
    Logger.log("Error guardarRespuestas: " + e);
    return {ok: false, msg: "Error al guardar. Contacte a Capital Humano: " + e.toString()};
  }
}

// ============================================================================
// UTIL: verificar si correo ya tiene respuesta registrada
// ============================================================================
function _yaRespondio(correo) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_RESPUESTAS);
    if (!hoja || hoja.getLastRow() < 2) return false;
    const datos = hoja.getRange(2, 2, hoja.getLastRow() - 1, 1).getValues();
    return datos.some(r => String(r[0]).trim().toLowerCase() === correo);
  } catch(e) { return false; }
}

// ============================================================================
// UTIL: devolver bloques al frontend
// ============================================================================
function obtenerBloques() {
  return BLOQUES_CLEAVER;
}

// ============================================================================
// ADMIN: Crear hojas auxiliares si no existen (ejecutar UNA VEZ)
// ============================================================================
function inicializarHojasApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(CFG_APP.HOJA_TOKENS)) {
    const h = ss.insertSheet(CFG_APP.HOJA_TOKENS);
    h.getRange(1,1,1,6).setValues([["Token","Correo","Nombre","Puesto","Usado","Fecha_Uso"]]);
    h.getRange(1,1,1,6).setBackground("#0d5a7a").setFontColor("white").setFontWeight("bold");
    Logger.log("✅ Hoja Tokens_Candidatos creada");
  }

  if (!ss.getSheetByName(CFG_APP.HOJA_DOMINIOS)) {
    const h = ss.insertSheet(CFG_APP.HOJA_DOMINIOS);
    h.getRange(1,1,1,2).setValues([["Dominio","Descripción"]]);
    h.getRange(1,1,1,2).setBackground("#0d5a7a").setFontColor("white").setFontWeight("bold");
    Logger.log("✅ Hoja Dominios_Permitidos creada — agrega tus dominios");
  }

  SpreadsheetApp.getUi().alert(
    "✅ Hojas auxiliares listas.\n\n" +
    "Próximos pasos:\n" +
    "1. Agrega dominios en 'Dominios_Permitidos'\n" +
    "2. Los tokens se generan desde 'Nuevos_Ingresos'\n" +
    "3. Publica el webapp como aplicación web"
  );
}

// ============================================================================
// ADMIN: Sincronizar tokens desde Nuevos_Ingresos → Tokens_Candidatos
// ============================================================================
function sincronizarTokensDesdeNuevosIngresos() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const hNI   = ss.getSheetByName("Nuevos_Ingresos");
    const hTok  = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hNI || !hTok) { Logger.log("Hojas no encontradas"); return; }

    const dataNI  = hNI.getDataRange().getValues();
    const dataTok = hTok.getDataRange().getValues();

    const tokensExistentes = dataTok.slice(1).map(r => String(r[0]).trim().toUpperCase());

    const heads     = dataNI[0].map(h => String(h).trim().toUpperCase());
    const colID     = heads.findIndex(h => h.includes("ID_PROCESO"));
    const colNombre = heads.findIndex(h => h === "NOMBRE");
    const colCorreo = heads.findIndex(h => h === "CORREO");
    const colPuesto = heads.findIndex(h => h.includes("PUESTO"));

    let nuevos = 0;
    for (let i = 1; i < dataNI.length; i++) {
      const id = String(dataNI[i][colID] || "").trim().toUpperCase();
      if (!id || id === "ID_PROCESO") continue;
      if (tokensExistentes.includes(id)) continue;

      hTok.appendRow([
        id,
        colCorreo > -1 ? dataNI[i][colCorreo] : "",
        colNombre > -1 ? dataNI[i][colNombre] : "",
        colPuesto > -1 ? dataNI[i][colPuesto] : "",
        "NO",
        ""
      ]);
      nuevos++;
    }

    Logger.log("✅ Tokens sincronizados: " + nuevos + " nuevos");
    SpreadsheetApp.getUi().alert(
      "✅ Sincronización completada.\n" + nuevos + " token(s) nuevos agregados a Tokens_Candidatos."
    );
  } catch(e) {
    Logger.log("Error sincronizarTokens: " + e);
  }
}

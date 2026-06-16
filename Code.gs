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
  HOJA_RESPUESTAS: "Base_Cleaver_CANDIDATOS",
  HOJA_RESULTADOS: "Resultados_DISC"
};

function _ss() { return SpreadsheetApp.openById(CFG_APP.SS_ID); }

// ── BLOQUES CLEAVER ───────────────────────────────────────────────────────────
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

// ── MAPEO ADJETIVO → FACTOR DISC ─────────────────────────────────────────────
const DISC_MAP = {
  // D — Dominancia
  'AGRESIVO':'D','ANIMOSO':'D','ATREVIDO':'D','AUDAZ':'D','AVENTURERO':'D',
  'CARACTER FIRME':'D','COMPETITIVO':'D','CONFIANZA EN SI MISMO':'D',
  'DECISIVO':'D','DETERMINADO':'D','DISPUTADOR':'D','EMPRENDEDOR':'D',
  'FRANCO':'D','FUERZA DE VOLUNTAD':'D','INCONQUISTABLE':'D','INQUIETO':'D',
  'ORIGINAL':'D','TENAZ':'D','VALIENTE':'D','VIGOROSO':'D','POSITIVO':'D',
  'DESEOSO':'D','PERSISTENTE':'D','CONFIADO':'D',
  // I — Influencia
  'ADMIRABLE':'I','ALEGRE':'I','ALMA DE LA FIESTA':'I','AMIGUERO':'I',
  'ATRACTIVO':'I','BUEN COMPAÑERO':'I','CONVINCENTE':'I','CORDIAL':'I',
  'ENCANTADOR':'I','ENTUSIASTA':'I','JOVIAL':'I','JUGUETON':'I',
  'MENTE ABIERTA':'I','NERVIOSO':'I','OPTIMISTA':'I','PARLANCHIN':'I',
  'PERSUASIVO':'I','POPULAR':'I','SANGRE LIVIANA':'I','SOCIABLE':'I',
  'GENEROSO':'I','INSPIRADOR':'I','SIMPATIZADOR':'I',
  // S — Estabilidad
  'ADAPTABLE':'S','AGRADABLE':'S','ARMONIOSO':'S','BUEN VECINO':'S',
  'COMODINO':'S','CONFORME':'S','CONSIDERADO':'S','DEVOTO':'S',
  'DISPUESTO':'S','DOCIL':'S','ECUANIME':'S','GENTIL':'S',
  'HUMILDE':'S','INDULGENTE':'S','LEAL':'S','PACIFICO':'S',
  'PACIENTE':'S','RECEPTIVO':'S','RESIGNADO':'S','RESPETUOSO':'S',
  'SATISFECHO':'S','SERVICIAL':'S','SUMISO':'S','TOLERANTE':'S',
  'BONDADOSO':'S','OBEDIENTE':'S','COMPLACIENTE':'S','CONVENCIONAL':'S',
  // C — Cumplimiento
  'AFIRMATIVO':'C','CAUTELOSO':'C','COHIBIDO':'C','CONSECUENTE':'C',
  'CONTROLADO':'C','CONFIABLE':'C','DIPLOMATICO':'C','DISCIPLINADO':'C',
  'EXACTO':'C','MESURADO PARA HABLAR':'C','MODERADO':'C','PRECISO':'C',
  'QUISQUILLOSO':'C','REFINADO':'C','TEMEROSO':'C','TEMEROSO DE DIOS':'C',
  'TIMIDO':'C','INDIFERENTE':'C','ESTETA':'C'
};

// ── PERFILES DISC (16 combinaciones) ─────────────────────────────────────────
const PERFILES_DISC = {
  'D' : { nombre:'Director',        desc:'Orientado a resultados con alta capacidad de decisión. Asume riesgos, ejerce liderazgo con autoridad y opera a un ritmo acelerado. Le motiva ganar y superar metas.' },
  'DI': { nombre:'Promotor',        desc:'Enérgico, persuasivo y orientado a resultados. Combina liderazgo con carisma para motivar y mover equipos hacia metas ambiciosas con entusiasmo.' },
  'DS': { nombre:'Desarrollador',   desc:'Decidido y orientado a las personas. Establece metas firmes y construye relaciones de largo plazo para alcanzarlas de forma sostenida y consistente.' },
  'DC': { nombre:'Perfeccionista',  desc:'Exigente consigo mismo y con los demás. Busca excelencia, control y precisión sobre los resultados. Analiza antes de actuar pero actúa con determinación.' },
  'I' : { nombre:'Inspirador',      desc:'Comunicador nato con energía positiva contagiosa. Motiva a otros, disfruta los entornos sociales y genera entusiasmo y creatividad en el equipo.' },
  'ID': { nombre:'Agente de Cambio',desc:'Dinámico e influyente. Combina energía social con capacidad de acción directa. Impulsa iniciativas y arrastra al equipo con visión y entusiasmo.' },
  'IS': { nombre:'Consejero',       desc:'Cálido, empático y profundamente orientado a las personas. Excelente construyendo relaciones de confianza y generando ambientes de trabajo colaborativos y positivos.' },
  'IC': { nombre:'Evaluador',       desc:'Analítico y sociable. Combina precisión y rigor con habilidad para comunicar ideas complejas de forma accesible y convincente para su audiencia.' },
  'S' : { nombre:'Especialista',    desc:'Estable, leal y confiable. Brinda soporte constante al equipo, trabaja mejor en entornos con relaciones sólidas, expectativas claras y ritmo predecible.' },
  'SD': { nombre:'Investigador',    desc:'Metódico y orientado a las personas. Analiza situaciones antes de actuar, pero cuando lo hace es firme, consistente y genera confianza en su entorno.' },
  'SI': { nombre:'Armonizador',     desc:'Cooperativo y comunicativo. Facilita el trabajo en equipo, construye acuerdos naturalmente y prefiere el consenso y la armonía sobre la confrontación.' },
  'SC': { nombre:'Coordinador',     desc:'Sistemático y estable. Excelente planificando y ejecutando procesos con consistencia, rigor y alto nivel de cumplimiento. Confiable en cada entrega.' },
  'C' : { nombre:'Analista',        desc:'Detallista, preciso y orientado a la calidad. Evalúa cuidadosamente antes de tomar decisiones y prefiere datos y hechos comprobados sobre la intuición.' },
  'CD': { nombre:'Observador',      desc:'Cauteloso y orientado a resultados. Evalúa riesgos antes de decidir y asegura que cada acción esté respaldada por análisis sólido y evidencia objetiva.' },
  'CI': { nombre:'Diplomático',     desc:'Preciso y sociable. Combina rigor analítico con habilidades de comunicación. Excelente mediador en situaciones complejas, negociador efectivo y empático.' },
  'CS': { nombre:'Defensor',        desc:'Meticuloso y empático. Asegura que los procesos sean correctos y que las personas estén bien atendidas. Une el rigor con el cuidado genuino por otros.' }
};

// ── DATOS POR FACTOR (fortalezas y áreas de desarrollo) ──────────────────────
const DISC_DATOS = {
  D: {
    fortalezas: [
      'Orientación clara a resultados y metas',
      'Toma de decisiones rápida y firme',
      'Liderazgo con autoridad y determinación',
      'Capacidad para asumir riesgos calculados',
      'Empuje ante obstáculos y adversidad'
    ],
    areas: [
      'Impaciencia con procesos lentos o detallados',
      'Puede descuidar aspectos relacionales del equipo',
      'Dificultad para escuchar activamente antes de decidir',
      'Puede percibirse como impositivo o insensible'
    ]
  },
  I: {
    fortalezas: [
      'Comunicación efectiva y persuasiva',
      'Entusiasmo contagioso que motiva al equipo',
      'Construcción natural de redes de contacto',
      'Creatividad e innovación constante en ideas',
      'Habilidad para vender visiones y conceptos'
    ],
    areas: [
      'Dificultad para mantener seguimiento sistemático',
      'Tendencia a prometer más de lo que puede cumplir',
      'Dispersión entre múltiples proyectos simultáneos',
      'Sensibilidad elevada a la crítica o el rechazo'
    ]
  },
  S: {
    fortalezas: [
      'Lealtad y confiabilidad sostenida en el tiempo',
      'Excelente trabajo en equipo y colaboración genuina',
      'Paciencia y consistencia en la ejecución',
      'Estabilidad emocional en situaciones de presión',
      'Apoyo genuino a compañeros sin buscar reconocimiento'
    ],
    areas: [
      'Resistencia al cambio rápido o no planificado',
      'Dificultad para establecer límites y decir no',
      'Tendencia a postergar conversaciones difíciles',
      'Exceso de complacencia que puede limitar su crecimiento'
    ]
  },
  C: {
    fortalezas: [
      'Precisión y atención al detalle',
      'Análisis profundo antes de tomar decisiones',
      'Altos estándares de calidad en cada entrega',
      'Seguimiento riguroso a procesos y normativas',
      'Pensamiento crítico y evaluación objetiva'
    ],
    areas: [
      'Parálisis por análisis excesivo antes de actuar',
      'Perfeccionismo que puede frenar la ejecución',
      'Rigidez ante procedimientos que cambian',
      'Comunicación percibida como fría o distante'
    ]
  }
};

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
    const hoja = _ss().getSheetByName(CFG_APP.HOJA_DOMINIOS);
    if (!hoja) return {ok: false, msg: "Hoja 'Dominios_Permitidos' no encontrada."};

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
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
    const hoja = _ss().getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hoja) return {ok: false, msg: "Sistema de tokens no configurado."};

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

// ── GUARDAR RESPUESTAS + CALCULAR DISC ───────────────────────────────────────
function guardarRespuestas(payload) {
  try {
    const ss   = _ss();
    const hoja = ss.getSheetByName(CFG_APP.HOJA_RESPUESTAS);
    if (!hoja) return {ok: false, msg: "Hoja de respuestas no encontrada."};

    if (!payload.correo || !payload.respuestas || payload.respuestas.length !== 24) {
      return {ok: false, msg: "Datos incompletos."};
    }

    const fecha = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm:ss");

    // ── Guardar respuestas crudas ─────────────────────────────────────────────
    let fila = [fecha, payload.correo, payload.nombre, payload.puesto];
    payload.respuestas.forEach(r => { fila.push(r.mas); fila.push(r.menos); });

    if (hoja.getLastRow() === 0) {
      let heads = ["Marca temporal","Correo","Nombre completo","Area / Puesto"];
      for (let b = 1; b <= 24; b++) { heads.push("Bloque " + b + " MAS"); heads.push("Bloque " + b + " MENOS"); }
      hoja.appendRow(heads);
    }
    hoja.appendRow(fila);
    SpreadsheetApp.flush();

    // ── Marcar token como usado ───────────────────────────────────────────────
    if (payload.tipo === "CANDIDATO" && payload.tokenFila) {
      const hTok = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
      if (hTok) {
        hTok.getRange(payload.tokenFila, 5).setValue("SI");
        hTok.getRange(payload.tokenFila, 6).setValue(fecha);
        SpreadsheetApp.flush();
      }
    }

    // ── Calcular perfil DISC ──────────────────────────────────────────────────
    const disc = calcularPerfilDISC(payload.respuestas);

    // ── Escribir en Resultados_DISC ───────────────────────────────────────────
    const hRes = ss.getSheetByName(CFG_APP.HOJA_RESULTADOS);
    if (hRes) {
      if (hRes.getLastRow() === 0) {
        const heads = ["Fecha","Nombre","Correo","Puesto","Tipo","D","I","S","C","Perfil","Descripcion"];
        hRes.appendRow(heads);
        hRes.getRange(1, 1, 1, heads.length)
            .setBackground("#0d5a7a").setFontColor("white").setFontWeight("bold");
      }
      hRes.appendRow([
        fecha, payload.nombre, payload.correo, payload.puesto, payload.tipo,
        disc.neto.D, disc.neto.I, disc.neto.S, disc.neto.C,
        disc.perfil.nombre, disc.perfil.desc
      ]);
      SpreadsheetApp.flush();
    }

    Logger.log("DISC guardado: " + payload.nombre + " — " + disc.perfil.nombre);
    return {ok: true, msg: "Evaluacion guardada.", disc: disc};

  } catch(e) {
    Logger.log("Error guardarRespuestas: " + e);
    return {ok: false, msg: "Error al guardar: " + e.toString()};
  }
}

// ── CALCULAR PERFIL DISC ──────────────────────────────────────────────────────
function calcularPerfilDISC(respuestas) {
  const sc = {D:{m:0,l:0}, I:{m:0,l:0}, S:{m:0,l:0}, C:{m:0,l:0}};

  respuestas.forEach(function(r) {
    const fm = DISC_MAP[String(r.mas  || '').trim().toUpperCase()];
    const fl = DISC_MAP[String(r.menos|| '').trim().toUpperCase()];
    if (fm && sc[fm]) sc[fm].m++;
    if (fl && sc[fl]) sc[fl].l++;
  });

  const neto = {
    D: sc.D.m - sc.D.l,
    I: sc.I.m - sc.I.l,
    S: sc.S.m - sc.S.l,
    C: sc.C.m - sc.C.l
  };

  const sorted = [['D',neto.D],['I',neto.I],['S',neto.S],['C',neto.C]]
                  .sort(function(a,b){ return b[1]-a[1]; });

  const top = sorted[0][0];
  const sec = sorted[1][1] > 0 ? sorted[1][0] : '';
  const key = top + sec;

  const perfil = PERFILES_DISC[key] || PERFILES_DISC[top] ||
    {nombre:'Equilibrado', desc:'Perfil balanceado con flexibilidad para adaptarse a distintos contextos y demandas.'};

  const topDatos = DISC_DATOS[top]   || {fortalezas:[], areas:[]};
  const secDatos = DISC_DATOS[sec]   || {fortalezas:[], areas:[]};

  const fortalezas = topDatos.fortalezas.slice(0,4).concat(secDatos.fortalezas.slice(0,2));
  const areas      = topDatos.areas.slice(0,3).concat(secDatos.areas.slice(0,2));

  return {
    sc   : sc,
    neto : neto,
    top  : top,
    sec  : sec,
    key  : key,
    perfil: perfil,
    fortalezas: fortalezas,
    areas: areas
  };
}

// ── OBTENER BLOQUES ───────────────────────────────────────────────────────────
function obtenerBloques() {
  return BLOQUES_CLEAVER;
}

// ── UTIL ──────────────────────────────────────────────────────────────────────
function _yaRespondio(correo) {
  try {
    const hoja = _ss().getSheetByName(CFG_APP.HOJA_RESPUESTAS);
    if (!hoja || hoja.getLastRow() < 2) return false;
    const datos = hoja.getRange(2, 2, hoja.getLastRow() - 1, 1).getValues();
    return datos.some(function(r){ return String(r[0]).trim().toLowerCase() === correo; });
  } catch(e) { return false; }
}

// ── ADMIN: inicializar hojas ──────────────────────────────────────────────────
function inicializarHojasApp() {
  const ss = _ss();
  const azdk = "#0d5a7a";

  const hojas = [
    { nombre: CFG_APP.HOJA_TOKENS,     heads: ["Token","Correo","Nombre","Puesto","Usado","Fecha_Uso"] },
    { nombre: CFG_APP.HOJA_DOMINIOS,   heads: ["Dominio","Descripcion"] },
    { nombre: CFG_APP.HOJA_RESULTADOS, heads: ["Fecha","Nombre","Correo","Puesto","Tipo","D","I","S","C","Perfil","Descripcion"] }
  ];

  hojas.forEach(function(def) {
    if (!ss.getSheetByName(def.nombre)) {
      const h = ss.insertSheet(def.nombre);
      h.appendRow(def.heads);
      h.getRange(1, 1, 1, def.heads.length).setBackground(azdk).setFontColor("white").setFontWeight("bold");
      Logger.log("Creada: " + def.nombre);
    }
  });

  if (!ss.getSheetByName(CFG_APP.HOJA_RESPUESTAS)) {
    ss.insertSheet(CFG_APP.HOJA_RESPUESTAS);
    Logger.log("Creada: " + CFG_APP.HOJA_RESPUESTAS);
  }

  SpreadsheetApp.getUi().alert("Hojas listas. Agrega dominios en Dominios_Permitidos y publica el webapp.");
}

// ── ADMIN: sincronizar tokens ─────────────────────────────────────────────────
function sincronizarTokensDesdeNuevosIngresos() {
  try {
    const ss   = _ss();
    const hNI  = ss.getSheetByName("Nuevos_Ingresos");
    const hTok = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hNI || !hTok) { SpreadsheetApp.getUi().alert("Hojas no encontradas."); return; }

    const dataNI  = hNI.getDataRange().getValues();
    const dataTok = hTok.getDataRange().getValues();
    const existentes = dataTok.slice(1).map(function(r){ return String(r[0]).trim().toUpperCase(); });

    const heads     = dataNI[0].map(function(h){ return String(h).trim().toUpperCase(); });
    const colID     = heads.findIndex(function(h){ return h.includes("ID_PROCESO"); });
    const colNombre = heads.findIndex(function(h){ return h === "NOMBRE"; });
    const colCorreo = heads.findIndex(function(h){ return h === "CORREO"; });
    const colPuesto = heads.findIndex(function(h){ return h.includes("PUESTO"); });

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
    SpreadsheetApp.getUi().alert(nuevos + " token(s) nuevos sincronizados.");
  } catch(e) {
    Logger.log("Error sincronizarTokens: " + e);
  }
}

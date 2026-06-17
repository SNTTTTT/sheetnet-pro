/**
 * SOLU · Capital Humano
 * Evaluación de Perfil Conductual — Método CLEAVER
 */

// ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
const CFG_APP = {
  SS_ID              : "17guNUaSg2bHQ9vd2HiS0f02YK92R-cl1rywFWFd1JN0",
  HOJA_TOKENS        : "Tokens_Candidatos",
  HOJA_DOMINIOS      : "Dominios_Permitidos",
  HOJA_USUARIOS_CH   : "USUARIOS_CH",
  HOJA_PLANTILLA     : "Base_Plantilla_DISC",
  HOJA_CANDIDATOS    : "Base_Candidatos_DISC",
  HOJA_CONCENTRADO   : "CONCENTRADO",
  HOJA_VALIDACION    : "VALIDACION",
  HOJA_CALCULOS      : "CALCULOS",
  HOJA_INTERPRETACION: "INTERPRETACION",
  HOJA_NI            : "Nuevos_Ingresos"
};

function _ss()          { return SpreadsheetApp.openById(CFG_APP.SS_ID); }
function _pct(v, total) { return total > 0 ? parseFloat((v / total * 100).toFixed(1)) : 0; }
function _hdr(h, n)     { h.getRange(1,1,1,n).setBackground('#0d5a7a').setFontColor('#ffffff').setFontWeight('bold'); }

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

// ── FACTOR POR ADJETIVO (DISC_MAP) ───────────────────────────────────────────
const DISC_MAP = {
  // D — Dominancia
  'AGRESIVO':'D','ATREVIDO':'D','AUDAZ':'D','AVENTURERO':'D',
  'CARACTER FIRME':'D','COMPETITIVO':'D','CONFIANZA EN SI MISMO':'D',
  'DECISIVO':'D','DETERMINADO':'D','DISPUTADOR':'D','EMPRENDEDOR':'D',
  'FRANCO':'D','FUERZA DE VOLUNTAD':'D','INCONQUISTABLE':'D','INQUIETO':'D',
  'ORIGINAL':'D','TENAZ':'D','VALIENTE':'D','VIGOROSO':'D','POSITIVO':'D',
  'PERSISTENTE':'D','CONFIADO':'D','AFIRMATIVO':'D','CONSECUENTE':'D','NERVIOSO':'D',
  // I — Influencia
  'ADMIRABLE':'I','ALEGRE':'I','ALMA DE LA FIESTA':'I','AMIGUERO':'I',
  'ATRACTIVO':'I','BUEN COMPAÑERO':'I','CONVINCENTE':'I','CORDIAL':'I',
  'ENCANTADOR':'I','ENTUSIASTA':'I','JOVIAL':'I','JUGUETON':'I',
  'MENTE ABIERTA':'I','OPTIMISTA':'I','PARLANCHIN':'I',
  'PERSUASIVO':'I','POPULAR':'I','SANGRE LIVIANA':'I','SOCIABLE':'I',
  'INSPIRADOR':'I','SIMPATIZADOR':'I','ANIMOSO':'I','DESEOSO':'I','REFINADO':'I',
  // S — Estabilidad
  'ADAPTABLE':'S','AGRADABLE':'S','ARMONIOSO':'S','BUEN VECINO':'S',
  'COMODINO':'S','CONFORME':'S','CONSIDERADO':'S',
  'DISPUESTO':'S','DOCIL':'S','ECUANIME':'S','GENTIL':'S',
  'INDULGENTE':'S','LEAL':'S','PACIFICO':'S',
  'PACIENTE':'S','RECEPTIVO':'S','RESIGNADO':'S',
  'SATISFECHO':'S','SERVICIAL':'S','SUMISO':'S','TOLERANTE':'S',
  'BONDADOSO':'S','COMPLACIENTE':'S','GENEROSO':'S','TIMIDO':'S','INDIFERENTE':'S','MODERADO':'S',
  // C — Cumplimiento
  'CAUTELOSO':'C','COHIBIDO':'C','CONTROLADO':'C','CONFIABLE':'C',
  'DIPLOMATICO':'C','DISCIPLINADO':'C',
  'EXACTO':'C','MESURADO PARA HABLAR':'C','PRECISO':'C',
  'QUISQUILLOSO':'C','TEMEROSO':'C','TEMEROSO DE DIOS':'C','ESTETA':'C',
  'HUMILDE':'C','DEVOTO':'C','OBEDIENTE':'C','RESPETUOSO':'C','CONVENCIONAL':'C'
};

// ── DATOS POR FACTOR ──────────────────────────────────────────────────────────
const DISC_DATOS = {
  D: {
    textoM: 'Se autopercibe como una persona decisiva, competitiva y orientada a resultados. En su mejor versión actúa con firmeza, toma la iniciativa y busca superar metas.',
    textoL: 'Su imagen adaptativa proyecta control, autoridad y determinación. Muestra poca flexibilidad ante la oposición y puede percibirse como directo o exigente.',
    textoT: 'Perfil total dominante: alta energía, orientación al logro. Prefiere ambientes retadores con autonomía y decisión propia.',
    textoGrafM: 'Alto D — Estilo directivo natural',
    textoGrafL: 'Alto D — Imagen controladora adaptativa',
    textoGrafT: 'Alto D — Enfoque orientado a resultados',
    fortalezas: ['Orientación clara a resultados y metas','Toma de decisiones rápida y firme','Liderazgo con autoridad y determinación','Capacidad para asumir riesgos calculados','Empuje constante ante obstáculos'],
    areas: ['Impaciencia con procesos lentos o detallados','Puede descuidar aspectos relacionales del equipo','Dificultad para escuchar activamente','Puede percibirse como impositivo'],
    motivadores: 'Control sobre decisiones, retos de alta complejidad, reconocimiento por resultados, autonomía operativa.',
    limitaciones: 'Impaciencia, dificultad para escuchar, tendencia al control excesivo, resistencia a colaborar en igualdad.',
    recomendaciones: 'Asignar proyectos con alta responsabilidad y autonomía. Proveer metas claras y métricas de éxito. Evitar microgestión.',
    alertas: 'Puede ignorar el proceso y saltarse pasos. Tiende a imponer decisiones sin consenso. Riesgo de conflicto con pares del mismo perfil.'
  },
  I: {
    textoM: 'Se autopercibe como persona sociable, entusiasta y comunicativa. En su mejor versión inspira y motiva a otros con energía positiva y facilidad de expresión.',
    textoL: 'Su imagen adaptativa proyecta carisma, optimismo y apertura. Puede percibirse como superficial si no hay sustancia detrás de la energía social.',
    textoT: 'Perfil total influyente: estilo comunicativo y relacional. Brilla en entornos colaborativos y de alta visibilidad.',
    textoGrafM: 'Alto I — Comunicador nato',
    textoGrafL: 'Alto I — Imagen carismática adaptativa',
    textoGrafT: 'Alto I — Enfoque relacional e influyente',
    fortalezas: ['Comunicación efectiva y persuasiva','Entusiasmo contagioso que motiva equipos','Red de contactos natural y amplia','Creatividad e innovación constante','Habilidad para vender ideas y visiones'],
    areas: ['Dificultad para mantener seguimiento sistemático','Tendencia a prometer más de lo que cumple','Dispersión entre múltiples proyectos','Sensibilidad elevada a la crítica'],
    motivadores: 'Reconocimiento público, libertad de expresión, trabajo en equipo dinámico, variedad y novedad.',
    limitaciones: 'Falta de seguimiento, promesas incumplidas, dificultad con tareas repetitivas o aisladas.',
    recomendaciones: 'Roles de cara al cliente o con visibilidad. Asignar un par analítico que dé seguimiento. Retroalimentación frecuente y positiva.',
    alertas: 'Puede prometer sin cumplir. Evita confrontaciones difíciles. Necesita supervisión en detalles y plazos.'
  },
  S: {
    textoM: 'Se autopercibe como persona leal, paciente y orientada al equipo. En su mejor versión es el sostén del grupo, constante y confiable bajo cualquier circunstancia.',
    textoL: 'Su imagen adaptativa proyecta calma, disponibilidad y apoyo. Puede percibirse como pasivo si no se le da espacio para contribuir.',
    textoT: 'Perfil total colaborativo y estable. Prefiere ambientes predecibles con relaciones sólidas y rutinas establecidas.',
    textoGrafM: 'Alto S — Soporte y lealtad',
    textoGrafL: 'Alto S — Imagen empática adaptativa',
    textoGrafT: 'Alto S — Enfoque colaborativo y estable',
    fortalezas: ['Lealtad y confiabilidad sostenida','Excelente trabajo en equipo y colaboración','Paciencia y consistencia en la ejecución','Estabilidad emocional bajo presión','Apoyo genuino sin buscar reconocimiento'],
    areas: ['Resistencia al cambio rápido','Dificultad para establecer límites y decir no','Tendencia a postergar conversaciones difíciles','Exceso de complacencia que limita el crecimiento'],
    motivadores: 'Seguridad laboral, relaciones de confianza, reconocimiento de lealtad, entorno armónico y predecible.',
    limitaciones: 'Resistencia al cambio, dificultad para confrontar, puede ser aprovechado, lentitud para adaptarse.',
    recomendaciones: 'Dar tiempo ante cambios. Reconocer explícitamente su contribución. Fomentar que exprese opiniones. Evitar sobrecarga silenciosa.',
    alertas: 'Puede acumular frustración sin expresarla. Riesgo de burnout por no poner límites. Necesita dirección clara para tomar iniciativa.'
  },
  C: {
    textoM: 'Se autopercibe como persona precisa, analítica y orientada a la calidad. Sus entregas son impecables y sus análisis son profundos y confiables.',
    textoL: 'Su imagen adaptativa proyecta rigor y alto estándar. Puede percibirse como rígido o crítico en entornos más informales.',
    textoT: 'Perfil total analítico y orientado al cumplimiento. Prefiere entornos donde los procesos están bien definidos y se respetan.',
    textoGrafM: 'Alto C — Analítico y preciso',
    textoGrafL: 'Alto C — Imagen rigurosa adaptativa',
    textoGrafT: 'Alto C — Enfoque de calidad y cumplimiento',
    fortalezas: ['Precisión y atención extrema al detalle','Análisis profundo antes de tomar decisiones','Altos estándares de calidad en cada entrega','Seguimiento riguroso a procesos y normativas','Pensamiento crítico y evaluación objetiva'],
    areas: ['Parálisis por análisis excesivo','Perfeccionismo que puede frenar la ejecución','Rigidez ante procedimientos que cambian','Comunicación percibida como fría o distante'],
    motivadores: 'Procesos claros, información completa, estándares de calidad, reconocimiento por precisión.',
    limitaciones: 'Lentitud por exceso de análisis, rigidez, dificultad para improvisar, puede ser muy crítico.',
    recomendaciones: 'Proporcionar contexto completo antes de asignar tareas. Respetar su necesidad de tiempo para analizar. Reconocer calidad sobre velocidad.',
    alertas: 'Puede bloquear el avance por buscar perfección. Conflicto con perfiles D que actúan rápido. Necesita claridad en criterios de éxito.'
  }
};

// ── PERFILES DISC ─────────────────────────────────────────────────────────────
const PERFILES_DISC = {
  'D':  { nombre:'Director',        segmento:'Dirección General / Alta Dirección / Emprendimiento',
          resumen:'Perfil de alto mando orientado a resultados. Toma decisiones firmes, opera a ritmo acelerado y lidera con autoridad.',
          desc:'Orientado a resultados con alta capacidad de decisión. Asume riesgos, ejerce liderazgo con autoridad y opera a un ritmo acelerado. Le motiva ganar y superar metas.',
          caract:'Directo, competitivo, independiente, orientado a resultados, alto ritmo, tolerante al riesgo.',
          motivadores:'Control, retos, autonomía, posiciones de poder, reconocimiento por logros.',
          limitaciones:'Impaciencia, ignora procesos, puede herir susceptibilidades, bajo trabajo en equipo.',
          recomendaciones:'Asignar proyectos de alta complejidad con autonomía. Dar metas claras. Evitar microgestión.',
          alertas:'Riesgo de conflicto con pares. Puede omitir procesos. Alta rotación si no tiene retos.' },
  'DI': { nombre:'Promotor',        segmento:'Ventas / Desarrollo de Negocios / Marketing',
          resumen:'Combina liderazgo directivo con carisma social. Enérgico, persuasivo y orientado a resultados.',
          desc:'Enérgico, persuasivo y orientado a resultados. Combina liderazgo con carisma para motivar y mover equipos hacia metas ambiciosas.',
          caract:'Enérgico, carismático, orientado a resultados, persuasivo, emprendedor, rápido en decisiones.',
          motivadores:'Reconocimiento público, retos de alta visibilidad, posiciones de influencia, metas ambiciosas.',
          limitaciones:'Puede prometer más de lo que cumple. Impaciencia con análisis. Poco enfoque en detalles.',
          recomendaciones:'Roles de liderazgo comercial o de proyectos. Asignar soporte analítico. Retroalimentación frecuente.',
          alertas:'Puede asumir compromisos sin respaldo. Conflicto si pierde visibilidad o reconocimiento.' },
  'DS': { nombre:'Desarrollador',   segmento:'Gestión de Proyectos / Producción / Operaciones',
          resumen:'Decidido y orientado a las personas. Establece metas firmes y construye relaciones de largo plazo.',
          desc:'Decidido y orientado a las personas. Establece metas firmes y construye relaciones de largo plazo para alcanzarlas de forma sostenida.',
          caract:'Orientado a resultados con sensibilidad social, construye equipos sólidos, consistente, leal a sus metas.',
          motivadores:'Logros concretos, equipo leal, estabilidad con reto, reconocimiento por impacto.',
          limitaciones:'Puede cargarse de responsabilidad. Dificultad para delegar. Poco tolerante con la mediocridad.',
          recomendaciones:'Proyectos de mediano plazo con equipo propio. Fomentar delegación. Reconocer logros del equipo.',
          alertas:'Puede sobrecargarse. Riesgo si el equipo no responde a su nivel de exigencia.' },
  'DC': { nombre:'Perfeccionista',  segmento:'Finanzas / Auditoría / Dirección de Calidad',
          resumen:'Exigente consigo mismo y con los demás. Busca excelencia, control y precisión sobre los resultados.',
          desc:'Exigente consigo mismo y con los demás. Busca excelencia, control y precisión. Analiza antes de actuar pero actúa con determinación.',
          caract:'Exigente, analítico, orientado a calidad y resultados, alto estándar, poca tolerancia al error.',
          motivadores:'Control sobre calidad, evidencia de resultados, reconocimiento por excelencia, poder con rigor.',
          limitaciones:'Puede ser percibido como inflexible. Dificultad para trabajar con perfiles más relajados.',
          recomendaciones:'Roles de control y dirección con impacto en calidad. Dar datos completos. Respetar su ritmo analítico.',
          alertas:'Puede bloquear equipos por estándares muy altos. Conflicto con perfiles I o S.' },
  'I':  { nombre:'Inspirador',      segmento:'Comunicación / Relaciones Públicas / Formación',
          resumen:'Comunicador nato con energía positiva contagiosa. Motiva a otros y genera entusiasmo y creatividad en el equipo.',
          desc:'Comunicador nato con energía positiva contagiosa. Motiva a otros, disfruta los entornos sociales y genera entusiasmo y creatividad.',
          caract:'Sociable, optimista, carismático, creativo, motivador, orientado a las personas.',
          motivadores:'Reconocimiento, libertad de expresión, interacción social, variedad y novedad.',
          limitaciones:'Falta de seguimiento, impuntualidad, dispersión, dificultad con tareas rutinarias.',
          recomendaciones:'Roles de comunicación interna, formación, ventas o relaciones públicas. Asignar estructura de soporte.',
          alertas:'Sin estructura puede dispersarse. Necesita seguimiento externo para compromisos formales.' },
  'ID': { nombre:'Agente de Cambio',segmento:'Liderazgo Comercial / Emprendimiento / Innovación',
          resumen:'Dinámico e influyente. Combina energía social con capacidad de acción directa. Impulsa iniciativas con visión.',
          desc:'Dinámico e influyente. Combina energía social con capacidad de acción directa. Impulsa iniciativas y arrastra al equipo con visión.',
          caract:'Innovador, enérgico, persuasivo, orientado a la acción, líder natural en entornos de cambio.',
          motivadores:'Visibilidad, impacto, libertad para ejecutar, equipo motivado, retos nuevos.',
          limitaciones:'Impaciencia con procesos, puede ignorar detalles, riesgo de sobrecomprometer.',
          recomendaciones:'Proyectos de transformación o innovación. Asignar soporte en ejecución y seguimiento.',
          alertas:'Alto riesgo de agotamiento con demasiados frentes abiertos. Necesita foco.' },
  'IS': { nombre:'Consejero',       segmento:'Recursos Humanos / Capacitación / Atención al Cliente',
          resumen:'Cálido, empático y profundamente orientado a las personas. Excelente construyendo relaciones de confianza.',
          desc:'Cálido, empático y profundamente orientado a las personas. Excelente construyendo relaciones de confianza y ambientes colaborativos.',
          caract:'Empático, sociable, cálido, orientado al bienestar del equipo, comunicador efectivo.',
          motivadores:'Relaciones genuinas, impacto en personas, reconocimiento de su contribución, armonía.',
          limitaciones:'Evita conflictos, puede sacrificar metas por relaciones, difícil con personas hostiles.',
          recomendaciones:'Roles de consultoría interna, RRHH, atención al cliente o capacitación. Fomentar asertividad.',
          alertas:'Puede retener información negativa para no herir. Necesita desarrollo en confrontación sana.' },
  'IC': { nombre:'Evaluador',       segmento:'Marketing / Consultoría Comunicacional / Análisis',
          resumen:'Analítico y sociable. Combina precisión y rigor con habilidad para comunicar ideas complejas.',
          desc:'Analítico y sociable. Combina precisión y rigor con habilidad para comunicar ideas complejas de forma accesible y convincente.',
          caract:'Sociable con fondo analítico, comunicador técnico, persuasivo con datos, empático e inteligente.',
          motivadores:'Proyectos que combinan análisis y comunicación, reconocimiento por rigor y claridad.',
          limitaciones:'Puede sobre-analizar antes de comunicar, lento en ejecutar si no tiene datos suficientes.',
          recomendaciones:'Roles de comunicación técnica, marketing de contenidos, consultoría. Dar datos y plataforma.',
          alertas:'Puede perderse en análisis y no ejecutar. Necesita plazos claros.' },
  'S':  { nombre:'Especialista',    segmento:'Administración / Soporte Operativo / Logística',
          resumen:'Estable, leal y confiable. Brinda soporte constante al equipo en entornos con relaciones sólidas y ritmo predecible.',
          desc:'Estable, leal y confiable. Brinda soporte constante al equipo, trabaja mejor en entornos con relaciones sólidas y ritmo predecible.',
          caract:'Leal, paciente, consistente, orientado al servicio, bajo perfil, soporte del equipo.',
          motivadores:'Seguridad, relaciones estables, reconocimiento de lealtad, entorno armónico.',
          limitaciones:'Resistencia al cambio, pasividad ante conflictos, puede ser aprovechado.',
          recomendaciones:'Roles de soporte o especialización técnica. Reconocer explícitamente su contribución. Dar tiempo ante cambios.',
          alertas:'Puede acumular frustración sin expresarla. Riesgo de estancamiento si no se le reta gradualmente.' },
  'SD': { nombre:'Investigador',    segmento:'Supervisión / Coordinación Operativa / Proyectos',
          resumen:'Metódico y orientado a las personas. Analiza antes de actuar, pero cuando lo hace es firme y consistente.',
          desc:'Metódico y orientado a las personas. Analiza situaciones antes de actuar, pero cuando lo hace es firme y consistente.',
          caract:'Estable con orientación a resultados, metódico, confiable, construye relaciones y entrega resultados.',
          motivadores:'Reconocimiento por resultados concretos, equipo leal, metas bien definidas, estabilidad.',
          limitaciones:'Lento para iniciar, puede ser demasiado cauteloso, necesita estructura para brillar.',
          recomendaciones:'Coordinación de procesos o equipos operativos. Dar dirección clara. Reconocer consistencia.',
          alertas:'Puede paralizarse ante ambigüedad. Necesita definición clara de expectativas.' },
  'SI': { nombre:'Armonizador',     segmento:'Trabajo Social / Psicología Organizacional / RRHH',
          resumen:'Cooperativo y comunicativo. Facilita el trabajo en equipo y construye acuerdos naturalmente.',
          desc:'Cooperativo y comunicativo. Facilita el trabajo en equipo, construye acuerdos naturalmente y prefiere el consenso sobre la confrontación.',
          caract:'Empático, comunicativo, orientado a la armonía del grupo, mediador natural, altamente colaborativo.',
          motivadores:'Trabajo en equipo, relaciones armónicas, impacto positivo en personas, reconocimiento grupal.',
          limitaciones:'Evita conflictos a toda costa, puede ceder ante presión, dificultad con personas dominantes.',
          recomendaciones:'Roles de mediación, trabajo social, atención interna. Fomentar asertividad.',
          alertas:'Puede comprometer sus propias necesidades por mantener la paz. Necesita apoyo para confrontar.' },
  'SC': { nombre:'Coordinador',     segmento:'Control de Calidad / Procesos / Administración',
          resumen:'Sistemático y estable. Excelente planificando y ejecutando procesos con consistencia y alto cumplimiento.',
          desc:'Sistemático y estable. Excelente planificando y ejecutando procesos con consistencia, rigor y alto nivel de cumplimiento.',
          caract:'Metódico, confiable, orientado a procesos, alta calidad en entrega, consistente.',
          motivadores:'Claridad en procesos, reconocimiento por calidad, estabilidad, entornos ordenados.',
          limitaciones:'Puede resistir cambios de proceso, poco flexible ante improvisación, lento ante urgencias no planificadas.',
          recomendaciones:'Coordinación de procesos, calidad o administración. Dar procedimientos claros. Avisar cambios con anticipación.',
          alertas:'Puede saturarse si los procesos cambian constantemente. Necesita predictibilidad para rendir al máximo.' },
  'C':  { nombre:'Analista',        segmento:'Auditoría / Finanzas / Sistemas / Control de Calidad',
          resumen:'Detallista, preciso y orientado a la calidad. Evalúa cuidadosamente antes de decidir.',
          desc:'Detallista, preciso y orientado a la calidad. Evalúa cuidadosamente antes de tomar decisiones y prefiere datos y hechos sobre la intuición.',
          caract:'Analítico, meticuloso, orientado a datos, alto estándar, sistemático, pensamiento crítico.',
          motivadores:'Datos completos, procesos definidos, reconocimiento por precisión, entornos de alta calidad.',
          limitaciones:'Parálisis por análisis, perfeccionismo, lento en ambientes de alta velocidad, comunicación fría.',
          recomendaciones:'Análisis, auditoría, finanzas, sistemas. Dar información completa. Respetar su necesidad de tiempo.',
          alertas:'Puede bloquear el avance por estándares muy altos. Necesita claridad en criterios de aceptación.' },
  'CD': { nombre:'Observador',      segmento:'Control de Gestión / Dirección Analítica / Consultoría',
          resumen:'Cauteloso y orientado a resultados. Evalúa riesgos antes de decidir con base en análisis sólido.',
          desc:'Cauteloso y orientado a resultados. Evalúa riesgos antes de decidir y asegura que cada acción esté respaldada por análisis sólido.',
          caract:'Analítico con enfoque a resultados, evalúa riesgos, decisivo cuando tiene información suficiente.',
          motivadores:'Datos, control, resultados medibles, reconocimiento por análisis riguroso.',
          limitaciones:'Puede ser lento si falta información, puede parecer frío o distante.',
          recomendaciones:'Roles de control estratégico o consultoría. Dar acceso a datos. Respetar su ritmo.',
          alertas:'Puede generar fricción al cuestionar decisiones sin datos. Necesita respaldo informacional sólido.' },
  'CI': { nombre:'Diplomático',     segmento:'Consultoría / Asesoría Técnica / Docencia',
          resumen:'Preciso y sociable. Combina rigor analítico con habilidades de comunicación. Excelente mediador.',
          desc:'Preciso y sociable. Combina rigor analítico con habilidades de comunicación. Excelente mediador en situaciones complejas y negociador efectivo.',
          caract:'Analítico y comunicativo, mediador técnico, orientado a calidad con sensibilidad social.',
          motivadores:'Proyectos que requieren análisis y comunicación, reconocimiento por precisión y trato.',
          limitaciones:'Puede sobre-analizar respuestas en conversaciones, lento para cerrar acuerdos.',
          recomendaciones:'Consultoría, formación, asesoría. Dar contexto completo. Fomentar decisiones oportunas.',
          alertas:'Puede perder oportunidades por buscar el momento perfecto. Necesita plazos que lo empujen.' },
  'CS': { nombre:'Defensor',        segmento:'Cumplimiento Normativo / Gestión de Riesgos / Calidad',
          resumen:'Meticuloso y empático. Asegura que los procesos sean correctos y que las personas estén bien atendidas.',
          desc:'Meticuloso y empático. Asegura que los procesos sean correctos y que las personas estén bien atendidas. Une el rigor con el cuidado genuino.',
          caract:'Meticuloso, empático, orientado a calidad y bienestar, confiable, coherente.',
          motivadores:'Procesos correctos, personas bien atendidas, reconocimiento por integridad y calidad.',
          limitaciones:'Puede sacrificar velocidad por calidad, dificultad para decir no, resistencia al cambio.',
          recomendaciones:'Cumplimiento, gestión de riesgos, calidad con impacto humano. Dar procesos y tiempo.',
          alertas:'Puede sobrecargarse al cuidar a todos. Necesita límites claros de responsabilidad.' }
};

// ── ENCABEZADOS ───────────────────────────────────────────────────────────────
function _headsBloques() {
  var h = [];
  for (var b = 1; b <= 24; b++) { h.push('Bloque ' + b + ' – MÁS'); h.push('Bloque ' + b + ' – MENOS'); }
  return h;
}

const HEADS_PLANTILLA     = ['Marca temporal','Dirección de correo electrónico','Nombre completo','Área / Linea de Negocio'].concat(_headsBloques());
const HEADS_CANDIDATOS    = ['Marca temporal','ID de proceso','Nombre completo','Correo personal','Teléfono','Puesto al que aplica'].concat(_headsBloques());
const HEADS_CONCENTRADO   = ['TIPO'].concat(HEADS_PLANTILLA);
const HEADS_VALIDACION    = ['Marca temporal','Correo','Nombre','Tipo','ESTADO','DETALLE'];
const HEADS_CALCULOS      = ['TIPO','Fecha','Correo','Nombre','Área','ID_GENERADO','D_MAS','I_MAS','S_MAS','C_MAS','D_MENOS','I_MENOS','S_MENOS','C_MENOS','D','I','S','C','Dom','D_M%','I_M%','S_M%','C_M%','D_L%','I_L%','S_L%','C_L%','D_T%','I_T%','S_T%','C_T%'];
const HEADS_INTERPRETACION = ['TIPO','Correo','NOMBRE','Perfil DISC','Combinación','Segmento Funcional','Resumen Ejecutivo','D_TOTAL','I_TOTAL','S_TOTAL','C_TOTAL','Características Completas','Motivadores','Limitaciones','Recomendaciones RH','Alertas','Texto_T','Texto_M','Texto_L','D_T%','I_T%','S_T%','C_T%','D_M%','I_M%','S_M%','C_M%','D_L%','I_L%','S_L%','C_L%','Texto_Grafica_T','Texto_Grafica_M','Texto_Grafica_L','Comb1_Nombre','Comb1_Caract','Comb1_Limit','Comb1_Deseos','Comb2_Nombre','Comb2_Caract','Comb2_Limit','Comb2_Deseos','Comb3_Nombre','Comb3_Caract','Comb3_Limit','Comb3_Deseos'];

// ── ENTRY POINT ───────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('SOLU · Evaluación DISC — CLEAVER')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1.0');
}

// ── AUTH: PERSONAL INTERNO ────────────────────────────────────────────────────
function validarCorreoCorporativo(correo) {
  try {
    correo = String(correo).trim().toLowerCase();
    if (!correo.includes('@')) return {ok: false, msg: "Correo inválido."};

    const dominio = correo.split('@')[1];
    const hoja    = _ss().getSheetByName(CFG_APP.HOJA_DOMINIOS);
    if (!hoja) return {ok: false, msg: "Configuración de dominios no encontrada."};

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const stored = String(datos[i][0]).trim().toLowerCase().replace(/^@/, '');
      if (stored !== dominio) continue;

      // CH puede volver a ingresar en cualquier momento; duplicados solo aplican a usuarios regulares
      const esAdmin = _esUsuarioCH(correo);
      if (!esAdmin && _yaRespondio(correo))
        return {ok: false, msg: "Este correo ya completó la evaluación.", repetido: true};

      return {ok: true, tipo: "PLANTILLA", correo: correo, esAdmin: esAdmin};
    }
    return {ok: false, msg: "El dominio de este correo no está autorizado."};
  } catch(e) {
    Logger.log("Error validarCorreoCorporativo: " + e);
    return {ok: false, msg: "Error interno. Contacte a Capital Humano."};
  }
}

// ── AUTH: CANDIDATO EXTERNO ───────────────────────────────────────────────────
function validarToken(token) {
  try {
    token = String(token).trim().toUpperCase();
    const hoja = _ss().getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hoja) return {ok: false, msg: "Sistema de tokens no configurado."};

    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toUpperCase() !== token) continue;
      if (String(datos[i][4]).trim().toUpperCase() === "SI")
        return {ok: false, msg: "Este token ya fue utilizado.", repetido: true};
      return {
        ok: true, tipo: "CANDIDATO",
        correo: String(datos[i][1]).trim(),
        nombre: String(datos[i][2]).trim(),
        puesto: String(datos[i][3]).trim(),
        fila  : i + 1
      };
    }
    return {ok: false, msg: "Token no encontrado. Verifique el código proporcionado."};
  } catch(e) {
    Logger.log("Error validarToken: " + e);
    return {ok: false, msg: "Error interno. Contacte a Capital Humano."};
  }
}

// ── GUARDAR RESPUESTAS + CALCULAR DISC ───────────────────────────────────────
function guardarRespuestas(payload) {
  try {
    const ss    = _ss();
    const fecha = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm:ss");
    const tipo  = payload.tipo || 'PLANTILLA';
    const esCandidato = tipo === 'CANDIDATO';

    if (!payload.correo || !payload.respuestas || payload.respuestas.length !== 24)
      return {ok: false, msg: "Datos incompletos."};

    var filaBloques = [];
    payload.respuestas.forEach(function(r) { filaBloques.push(r.mas); filaBloques.push(r.menos); });

    const disc = calcularPerfilDISC(payload.respuestas);

    // Validación de estado
    var estado = 'OK', detalle = 'Evaluación válida y completa.';
    var colisiones = payload.respuestas.filter(function(r){ return r.mas === r.menos; });
    if (colisiones.length > 0) {
      estado  = 'INVÁLIDA';
      detalle = 'MÁS = MENOS en bloque(s): ' + colisiones.map(function(r){ return r.bloque; }).join(', ');
    }

    // Hoja de respuestas por tipo
    if (esCandidato) {
      const tokenId = payload.tokenFila ? 'CAND-' + String(payload.tokenFila).padStart(4,'0') : '';
      const h = _getOrCreateSheet(ss, CFG_APP.HOJA_CANDIDATOS, HEADS_CANDIDATOS);
      h.appendRow([fecha, tokenId, payload.nombre, payload.correo, '', payload.puesto].concat(filaBloques));
    } else {
      const h = _getOrCreateSheet(ss, CFG_APP.HOJA_PLANTILLA, HEADS_PLANTILLA);
      h.appendRow([fecha, payload.correo, payload.nombre, payload.puesto].concat(filaBloques));
    }

    // CONCENTRADO
    const hConc = _getOrCreateSheet(ss, CFG_APP.HOJA_CONCENTRADO, HEADS_CONCENTRADO);
    hConc.appendRow([tipo, fecha, payload.correo, payload.nombre, payload.puesto].concat(filaBloques));

    // VALIDACION con color de estado
    const hVal = _getOrCreateSheet(ss, CFG_APP.HOJA_VALIDACION, HEADS_VALIDACION);
    hVal.appendRow([fecha, payload.correo, payload.nombre, tipo, estado, detalle]);
    const colorEstado = estado === 'INVÁLIDA' ? '#FECACA' : estado === 'DUPLICADA' ? '#FEF08A' : '#BBF7D0';
    hVal.getRange(hVal.getLastRow(), 5).setBackground(colorEstado);

    // CALCULOS
    const idGen = esCandidato
      ? 'CAND-' + String(payload.tokenFila || 0).padStart(4,'0')
      : payload.correo;
    const hCalc = _getOrCreateSheet(ss, CFG_APP.HOJA_CALCULOS, HEADS_CALCULOS);
    hCalc.appendRow([
      tipo, fecha, payload.correo, payload.nombre, payload.puesto, idGen,
      disc.sc.D.m, disc.sc.I.m, disc.sc.S.m, disc.sc.C.m,
      disc.sc.D.l, disc.sc.I.l, disc.sc.S.l, disc.sc.C.l,
      disc.neto.D, disc.neto.I, disc.neto.S, disc.neto.C,
      disc.top,
      _pct(disc.sc.D.m,24), _pct(disc.sc.I.m,24), _pct(disc.sc.S.m,24), _pct(disc.sc.C.m,24),
      _pct(disc.sc.D.l,24), _pct(disc.sc.I.l,24), _pct(disc.sc.S.l,24), _pct(disc.sc.C.l,24),
      _pct(disc.sc.D.m+disc.sc.D.l,48), _pct(disc.sc.I.m+disc.sc.I.l,48),
      _pct(disc.sc.S.m+disc.sc.S.l,48), _pct(disc.sc.C.m+disc.sc.C.l,48)
    ]);

    // INTERPRETACION
    const hInt = _getOrCreateSheet(ss, CFG_APP.HOJA_INTERPRETACION, HEADS_INTERPRETACION);
    const c1 = disc.combs[0]||{}, c2 = disc.combs[1]||{}, c3 = disc.combs[2]||{};
    hInt.appendRow([
      tipo, payload.correo, payload.nombre,
      disc.perfil.nombre, disc.key, disc.perfil.segmento, disc.perfil.resumen,
      disc.sc.D.m+disc.sc.D.l, disc.sc.I.m+disc.sc.I.l, disc.sc.S.m+disc.sc.S.l, disc.sc.C.m+disc.sc.C.l,
      disc.perfil.caract, disc.perfil.motivadores, disc.perfil.limitaciones,
      disc.perfil.recomendaciones, disc.perfil.alertas,
      disc.textoT, disc.textoM, disc.textoL,
      _pct(disc.neto.D+12,24), _pct(disc.neto.I+12,24), _pct(disc.neto.S+12,24), _pct(disc.neto.C+12,24),
      _pct(disc.sc.D.m,24), _pct(disc.sc.I.m,24), _pct(disc.sc.S.m,24), _pct(disc.sc.C.m,24),
      _pct(disc.sc.D.l,24), _pct(disc.sc.I.l,24), _pct(disc.sc.S.l,24), _pct(disc.sc.C.l,24),
      disc.textoGrafT, disc.textoGrafM, disc.textoGrafL,
      c1.nombre||'', c1.caract||'', c1.limit||'', c1.deseos||'',
      c2.nombre||'', c2.caract||'', c2.limit||'', c2.deseos||'',
      c3.nombre||'', c3.caract||'', c3.limit||'', c3.deseos||''
    ]);

    // Marcar token como usado
    if (esCandidato && payload.tokenFila) {
      const hTok = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
      if (hTok) {
        hTok.getRange(payload.tokenFila, 5).setValue("SI");
        hTok.getRange(payload.tokenFila, 6).setValue(fecha);
      }
    }

    SpreadsheetApp.flush();
    Logger.log("Guardado: " + payload.nombre + " — " + disc.perfil.nombre + " [" + tipo + "]");
    return {ok: true, msg: "Evaluación guardada.", disc: disc};

  } catch(e) {
    Logger.log("Error guardarRespuestas: " + e);
    return {ok: false, msg: "Error al guardar. Contacte a Capital Humano."};
  }
}

// ── HELPER: obtener o crear hoja ──────────────────────────────────────────────
function _getOrCreateSheet(ss, nombre, heads) {
  var h = ss.getSheetByName(nombre);
  if (!h) { h = ss.insertSheet(nombre); h.appendRow(heads); _hdr(h, heads.length); }
  else if (h.getLastRow() === 0) { h.appendRow(heads); _hdr(h, heads.length); }
  return h;
}

// ── CALCULAR PERFIL DISC ──────────────────────────────────────────────────────
function calcularPerfilDISC(respuestas) {
  const sc = {D:{m:0,l:0}, I:{m:0,l:0}, S:{m:0,l:0}, C:{m:0,l:0}};

  respuestas.forEach(function(r) {
    const fm = DISC_MAP[String(r.mas  ||'').trim().toUpperCase()];
    const fl = DISC_MAP[String(r.menos||'').trim().toUpperCase()];
    if (fm && sc[fm]) sc[fm].m++;
    if (fl && sc[fl]) sc[fl].l++;
  });

  const neto   = {D:sc.D.m-sc.D.l, I:sc.I.m-sc.I.l, S:sc.S.m-sc.S.l, C:sc.C.m-sc.C.l};
  const sorted = [['D',neto.D],['I',neto.I],['S',neto.S],['C',neto.C]].sort(function(a,b){ return b[1]-a[1]; });

  const top = sorted[0][0];
  const sec = sorted[1][1] > 0 ? sorted[1][0] : '';
  const key = top + sec;

  const perfil = PERFILES_DISC[key] || PERFILES_DISC[top] || {
    nombre:'Equilibrado', segmento:'Multifuncional',
    resumen:'Perfil balanceado con alta adaptabilidad a distintos contextos.',
    desc:'Perfil balanceado con flexibilidad para adaptarse a distintos contextos y demandas.',
    caract:'Adaptable, flexible, multifuncional.',
    motivadores:'Variedad y equilibrio en las tareas.',
    limitaciones:'Puede carecer de estilo propio definido.',
    recomendaciones:'Roles versátiles con variedad de actividades.',
    alertas:'Puede ser percibido como inconsistente en estilo de trabajo.'
  };

  const topD = DISC_DATOS[top] || {};
  const secD = DISC_DATOS[sec] || {};

  const fortalezas = (topD.fortalezas||[]).slice(0,4).concat((secD.fortalezas||[]).slice(0,2));
  const areas      = (topD.areas||[]).slice(0,3).concat((secD.areas||[]).slice(0,2));

  function _comb(f1, f2) {
    if (!f1 || !f2) return {};
    const pf = PERFILES_DISC[f1+f2] || PERFILES_DISC[f1] || {};
    return {nombre:pf.nombre||'', caract:pf.caract||'', limit:pf.limitaciones||'', deseos:pf.motivadores||''};
  }

  return {
    sc, neto, top, sec, key, perfil, fortalezas, areas,
    textoM    : topD.textoM     || '',
    textoL    : topD.textoL     || '',
    textoT    : topD.textoT     || '',
    textoGrafM: topD.textoGrafM || '',
    textoGrafL: topD.textoGrafL || '',
    textoGrafT: topD.textoGrafT || '',
    combs: [
      _comb(sorted[0][0], sorted[1][0]),
      _comb(sorted[0][0], sorted[2][0]),
      _comb(sorted[1][0], sorted[2][0])
    ]
  };
}

// ── OBTENER BLOQUES ───────────────────────────────────────────────────────────
function obtenerBloques() { return BLOQUES_CLEAVER; }

// ── VERIFICAR SI CORREO YA RESPONDIÓ ─────────────────────────────────────────
function _yaRespondio(correo) {
  try {
    const ss = _ss();
    for (const nombre of [CFG_APP.HOJA_PLANTILLA, CFG_APP.HOJA_CANDIDATOS]) {
      const h = ss.getSheetByName(nombre);
      if (!h || h.getLastRow() < 2) continue;
      const col  = nombre === CFG_APP.HOJA_CANDIDATOS ? 4 : 2;
      const datos = h.getRange(2, col, h.getLastRow()-1, 1).getValues();
      if (datos.some(function(r){ return String(r[0]).trim().toLowerCase() === correo; })) return true;
    }
    return false;
  } catch(e) { return false; }
}

// ── VERIFICAR SI ES USUARIO DE CAPITAL HUMANO ────────────────────────────────
function _esUsuarioCH(correo) {
  try {
    const hoja = _ss().getSheetByName(CFG_APP.HOJA_USUARIOS_CH);
    if (!hoja || hoja.getLastRow() < 2) return false;
    const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 2).getValues();
    return datos.some(function(r) {
      const email  = String(r[0]).trim().toLowerCase();
      const activo = String(r[1]).trim().toUpperCase();
      return email === correo && activo !== 'NO';
    });
  } catch(e) { return false; }
}

// ── ADMIN: INICIALIZAR HOJAS ──────────────────────────────────────────────────
function inicializarHojasApp() {
  const ss = _ss();
  const defs = [
    { nombre: CFG_APP.HOJA_TOKENS,         heads: ["Token","Correo","Nombre","Puesto","Usado","Fecha_Uso"] },
    { nombre: CFG_APP.HOJA_DOMINIOS,       heads: ["Dominio","Descripcion"] },
    { nombre: CFG_APP.HOJA_USUARIOS_CH,    heads: ["Correo","Activo","Nombre","Rol"] },
    { nombre: CFG_APP.HOJA_PLANTILLA,      heads: HEADS_PLANTILLA },
    { nombre: CFG_APP.HOJA_CANDIDATOS,     heads: HEADS_CANDIDATOS },
    { nombre: CFG_APP.HOJA_CONCENTRADO,    heads: HEADS_CONCENTRADO },
    { nombre: CFG_APP.HOJA_VALIDACION,     heads: HEADS_VALIDACION },
    { nombre: CFG_APP.HOJA_CALCULOS,       heads: HEADS_CALCULOS },
    { nombre: CFG_APP.HOJA_INTERPRETACION, heads: HEADS_INTERPRETACION }
  ];

  defs.forEach(function(def) {
    if (!ss.getSheetByName(def.nombre)) {
      const h = ss.insertSheet(def.nombre);
      h.appendRow(def.heads);
      _hdr(h, def.heads.length);
    }
  });

  SpreadsheetApp.getUi().alert(
    "Listo. " + defs.length + " hojas verificadas.\n\n" +
    "Siguiente paso:\n" +
    "1. Agrega tu correo en '" + CFG_APP.HOJA_USUARIOS_CH + "' (Activo = SI)\n" +
    "2. Agrega dominios en '" + CFG_APP.HOJA_DOMINIOS + "'\n" +
    "3. Publica el webapp (Nueva versión)"
  );
}

// ── ADMIN: SINCRONIZAR TOKENS DESDE NUEVOS_INGRESOS ──────────────────────────
function sincronizarTokensDesdeNuevosIngresos() {
  try {
    const ss   = _ss();
    const hNI  = ss.getSheetByName(CFG_APP.HOJA_NI);
    const hTok = ss.getSheetByName(CFG_APP.HOJA_TOKENS);
    if (!hNI || !hTok) { SpreadsheetApp.getUi().alert("Hojas no encontradas."); return; }

    const dataNI  = hNI.getDataRange().getValues();
    const dataTok = hTok.getDataRange().getValues();
    const existentes = dataTok.slice(1).map(function(r){ return String(r[0]).trim().toUpperCase(); });

    const heads  = dataNI[0].map(function(h){ return String(h).trim().toUpperCase(); });
    const colID  = heads.findIndex(function(h){ return h.includes("ID_PROCESO") || h.includes("ID PROCESO"); });
    const colNom = heads.findIndex(function(h){ return h === "NOMBRE" || h.includes("NOMBRE COMPLETO"); });
    const colCor = heads.findIndex(function(h){ return h === "CORREO" || h.includes("CORREO"); });
    const colPue = heads.findIndex(function(h){ return h.includes("PUESTO"); });

    let nuevos = 0;
    for (let i = 1; i < dataNI.length; i++) {
      const id = String(dataNI[i][colID >= 0 ? colID : 2] || '').trim().toUpperCase();
      if (!id || existentes.includes(id)) continue;
      hTok.appendRow([
        id,
        colCor >= 0 ? dataNI[i][colCor] : '',
        colNom >= 0 ? dataNI[i][colNom] : '',
        colPue >= 0 ? dataNI[i][colPue] : '',
        "NO", ""
      ]);
      nuevos++;
    }
    SpreadsheetApp.getUi().alert(nuevos + " token(s) agregados a " + CFG_APP.HOJA_TOKENS + ".");
  } catch(e) {
    Logger.log("Error sincronizarTokens: " + e);
    SpreadsheetApp.getUi().alert("Error: " + e.toString());
  }
}

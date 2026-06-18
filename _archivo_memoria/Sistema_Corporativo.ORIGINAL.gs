/**
 * ╔═══════════════════════════════════════════════════════════════════╗
 * ║  DISEÑO Y SOLUCIÓN INTEGRAL SANTILLÁN RAMÍREZ                     ║
 * ║  Sistema Corporativo Unificado · V4.2 (Limpieza y Optimizacion)   ║
 * ╚═══════════════════════════════════════════════════════════════════╝
 */

const SS_MASTER_ID    = "1rdfAXSEml3cDamy48r02FMrrZ19G0C1-LyRaK2oRQ4w";
const SH_USUARIOS     = "DATA_USUARIOS";
const SH_EMPLEADOS    = "DATA_EMPLEADOS";
const SH_COMPRAS      = "DB_COMPRAS";
const SH_ACTAS        = "DB_ACTAS";
const SH_SOPORTE      = "DB_SOPORTE";
const SH_EXPEDIENTES  = "DB_EXPEDIENTES";
const SH_FOTOS        = "DATA_FOTOS";
const SH_AUDITORIA    = "DB_AUDITORIA";
const DRIVE_EXP_ID    = "1r5_jIBIFHVCnGGkmy36BS9F34IbH8BJn";

/* ═══════════════════════════════════════════════════════════════════════════
   📋 ESTRUCTURA REQUERIDA EN DATA_USUARIOS — Columna 26 (índice 25)
   ═══════════════════════════════════════════════════════════════════════════
   Encabezado col 26: equipoNotificacion
   
   Valores válidos (separados por coma si aplica más de uno):
     SUPER_ADMIN          → Acceso total al sistema (reemplaza SUPER_ADMINS[])
     COMPRAS              → Recibe correos de nuevas solicitudes de compra
     COMPRAS_APROBADOR    → Autoriza compras (reemplaza EMAIL_OMAR / EMAIL_ISELA)
     SOPORTE              → Recibe tickets de soporte TI
     INFRA                → Equipo de infraestructura (reemplaza INFRA_TEAM[])
     MANT                 → Recibe alertas de mantenimiento
     CH                   → Capital Humano, recibe actas (reemplaza EMAIL_AUX_CH)
     RESIDUOS             → Recibe alertas de residuos peligrosos
     LIMPIEZA             → Recibe solicitudes de limpieza
   
   Ejemplo fila:
     Op001 | isaac@ssolu.com.mx | Isaac G. | DIRECCIÓN | Director | ADM | SI
           | SI | SI | DIRECCIÓN | SI | SI | SI | SI | SI | SI | SI | SI | SI
           | SI | SI | SI | SI | SI | SI | SUPER_ADMIN,COMPRAS
   
   NOTA: Un usuario puede tener múltiples roles: "SUPER_ADMIN,MANT,CH"
   Para agregar o quitar de un equipo: editar solo la hoja, sin tocar código.
═══════════════════════════════════════════════════════════════════════════ */

// ══════════════════════════════════════════════════════════════════════════════
// ✅ EMAILS Y EQUIPOS — 100% dinámicos desde DATA_USUARIOS (col 26: equipoNotif)
//    Sin hardcode. Agregar/quitar personas: solo editar la hoja, no el código.
//
//    COLUMNA 26 (índice 25) de DATA_USUARIOS → equipoNotificacion
//    Valores posibles (separados por coma para múltiples equipos):
//      SUPER_ADMIN | COMPRAS | SOPORTE | INFRA | MANT | CH | RESIDUOS | LIMPIEZA
//
//    EJEMPLO fila DATA_USUARIOS:
//    Op001 | isaac@ssolu.com.mx | Isaac G. | DIRECCIÓN | Director | ADMINISTRATIVO | SI
//          | SI(compras) | SI(comprasAdmin) | ... | SUPER_ADMIN,COMPRAS
// ══════════════════════════════════════════════════════════════════════════════

/**
 * _cargarUsuariosBD()
 * Lee DATA_USUARIOS completo (cols 1-26) con cache de script 10 min.
 * Retorna array de filas. Evita N lecturas repetidas en la misma ejecución.
 */
function _cargarUsuariosBD() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('_usBD_v1');
  if (cached) return JSON.parse(cached);
  const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 26).getValues();
  try { cache.put('_usBD_v1', JSON.stringify(rows), 600); } catch(e) {}
  return rows;
}

/**
 * _getSuperAdmins()
 * Retorna array de emails que tienen SUPER_ADMIN en col 26 (equipoNotif).
 * Reemplaza la constante SUPER_ADMINS hardcodeada.
 */
function _getSuperAdmins() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('_sa_v1');
  if (cached) return JSON.parse(cached);
  const rows = _cargarUsuariosBD();
  const admins = [];
  rows.forEach(r => {
    const equipo = String(r[25] || '').toUpperCase();
    const correo = String(r[1] || '').toLowerCase().trim();
    const activo = normalizarBooleano(r[6]);
    if (activo && correo.includes('@') && equipo.includes('SUPER_ADMIN')) {
      admins.push(correo);
    }
  });
  try { cache.put('_sa_v1', JSON.stringify(admins), 600); } catch(e) {}
  return admins;
}

/**
 * _getEquipoNotif(modulo)
 * Retorna array de emails del equipo de un módulo específico.
 * modulo: 'COMPRAS' | 'SOPORTE' | 'INFRA' | 'MANT' | 'CH' | 'RESIDUOS' | 'LIMPIEZA'
 * Reemplaza INFRA_TEAM, SOPORTE_TEAM y todos los EMAIL_XXX de destino masivo.
 */
function _getEquipoNotif(modulo) {
  const rows = _cargarUsuariosBD();
  const equipo = [];
  const moduloUp = String(modulo || '').toUpperCase().trim();
  rows.forEach(r => {
    const notif = String(r[25] || '').toUpperCase();
    const correo = String(r[1] || '').toLowerCase().trim();
    const activo = normalizarBooleano(r[6]);
    if (activo && correo.includes('@') && notif.includes(moduloUp)) {
      equipo.push(correo);
    }
  });
  return [...new Set(equipo)]; // sin duplicados
}

/**
 * _getReplyTo(modulo)
 * Retorna el primer correo del equipo para usarlo como replyTo en MailApp.
 * Si no hay nadie en el equipo, retorna cadena vacía.
 */
function _getReplyTo(modulo) {
  const equipo = _getEquipoNotif(modulo);
  return equipo.length > 0 ? equipo[0] : '';
}

// Compatibilidad: estas funciones sintetizan lo que antes hacían las constantes
// EMAIL_OMAR y EMAIL_ISELA (aprobadores de compras desde hoja, buscados por rol)
function _getAprobadoresCompras() {
  // Se buscan usuarios con flag COMPRAS_APROBADOR en col 26
  return _getEquipoNotif('COMPRAS_APROBADOR');
}

function getWebAppUrl() {
  try { return ScriptApp.getService().getUrl(); } catch(e) { return ""; }
}

/* ═══════════════════════════════════════════════════════════════════
   📧 SISTEMA DE EMAILS PREMIUM — Plantilla unificada corporativa
   Todos los correos usan _emailShell para diseño consistente
═══════════════════════════════════════════════════════════════════ */
function _emailShell(cfg) {
  var headerGrad = cfg.headerGrad || "linear-gradient(135deg,#1F3A4D 0%,#2C6C84 100%)";
  var accentColor = cfg.accent || "#54A9C8";
  var title = cfg.title || "";
  var subtitle = cfg.subtitle || "";
  var folioLabel = cfg.folioLabel || "FOLIO";
  var folioValue = cfg.folioValue || "";
  var body = cfg.body || "";
  var footerNote = cfg.footerNote || "";

  return '<div style="font-family:\'Segoe UI\',Arial,sans-serif;background:#F0F2F5;padding:32px 16px;line-height:1.5">' +
    '<div style="max-width:600px;margin:0 auto">' +
    
    // HEADER PREMIUM
    '<div style="background:' + headerGrad + ';border-radius:16px 16px 0 0;padding:0;position:relative;overflow:hidden">' +
      '<div style="position:absolute;inset:0;background-image:radial-gradient(circle,rgba(255,255,255,.04) 1px,transparent 1px);background-size:20px 20px;pointer-events:none"></div>' +
      '<div style="position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,' + accentColor + ',rgba(255,255,255,.3),' + accentColor + ')"></div>' +
      '<div style="padding:28px 32px;position:relative;z-index:1">' +
        '<table width="100%" cellpadding="0" cellspacing="0"><tr valign="middle">' +
          '<td>' +
            '<div style="font-family:monospace;font-size:9px;letter-spacing:4px;color:rgba(255,255,255,.5);font-weight:700;margin-bottom:8px;text-transform:uppercase">SANTILLÁN RAMÍREZ</div>' +
            '<div style="font-size:22px;font-weight:900;color:#FFFFFF;letter-spacing:-.5px">' + title + '</div>' +
            (subtitle ? '<div style="font-size:12px;color:rgba(255,255,255,.6);margin-top:6px">' + subtitle + '</div>' : '') +
          '</td>' +
          (folioValue ? '<td align="right" style="vertical-align:middle">' +
            '<div style="background:rgba(0,0,0,.2);border:1px solid rgba(255,255,255,.15);border-radius:12px;padding:12px 20px;text-align:center;min-width:80px">' +
              '<div style="font-family:monospace;font-size:8px;letter-spacing:2.5px;color:rgba(255,255,255,.55);margin-bottom:4px">' + folioLabel + '</div>' +
              '<div style="font-family:monospace;font-size:22px;font-weight:900;color:#FFFFFF;line-height:1">' + folioValue + '</div>' +
            '</div>' +
          '</td>' : '') +
        '</tr></table>' +
      '</div>' +
      '<div style="height:2px;background:linear-gradient(90deg,transparent,' + accentColor + ',transparent);opacity:.5"></div>' +
    '</div>' +
    
    // BODY
    '<div style="background:#FFFFFF;padding:28px 32px;border-left:1px solid #E8ECF0;border-right:1px solid #E8ECF0">' + body + '</div>' +
    
    // FOOTER
    '<div style="background:#FAFBFC;border:1px solid #E8ECF0;border-top:none;border-radius:0 0 16px 16px;padding:20px 32px;text-align:center">' +
      (footerNote ? '<div style="font-size:12px;color:#94A3B8;margin-bottom:10px">' + footerNote + '</div>' : '') +
      '<div style="display:inline-flex;align-items:center;gap:8px">' +
        '<div style="font-family:monospace;font-size:8px;letter-spacing:3px;color:#CBD5E1;text-transform:uppercase">Santillán Ramírez · Sistema Corporativo</div>' +
      '</div>' +
    '</div>' +
    
    '</div></div>';
}

function _emailInfoBlock(label, value, borderColor) {
  borderColor = borderColor || "#54A9C8";
  return '<div style="background:#F8FAFC;border-left:4px solid ' + borderColor + ';border-radius:0 10px 10px 0;padding:13px 16px;margin-bottom:10px">' +
    '<div style="font-family:monospace;font-size:8px;letter-spacing:2.5px;color:#94A3B8;margin-bottom:4px;text-transform:uppercase">' + label + '</div>' +
    '<div style="font-size:15px;font-weight:800;color:#0F172A">' + value + '</div>' +
  '</div>';
}

function _emailButton(href, text, color) {
  color = color || "#54A9C8";
  return '<a href="' + href + '" style="display:block;background:' + color + ';color:#FFFFFF;text-align:center;padding:14px 24px;border-radius:10px;text-decoration:none;font-weight:700;font-size:14px;letter-spacing:.5px;margin-top:16px">' + text + '</a>';
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) { return handleCallback(e.parameter); }
  return HtmlService.createTemplateFromFile("Index").evaluate()
    .setTitle("Sistema Corporativo | Santillán Ramírez")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ═══════════════════════════════════════════════════════════════
   AUDITORÍA — registra cada acción, sesión y movimiento
   ═══════════════════════════════════════════════════════════════ */
function registrarAuditoria(modulo, accion, detalle, referencia) {
  try {
    const ss  = SpreadsheetApp.openById(SS_MASTER_ID);
    let sh    = ss.getSheetByName(SH_AUDITORIA);
    if (!sh) {
      sh = ss.insertSheet(SH_AUDITORIA);
      const hdr = ["ID","FECHA","HORA","EMAIL","NOMBRE","AREA","MODULO","ACCION","DETALLE","REFERENCIA","IP_SESION"];
      sh.getRange(1,1,1,hdr.length).setValues([hdr]).setBackground("#1F3A4D").setFontColor("#FFFFFF").setFontWeight("bold");
      sh.setFrozenRows(1);
      sh.setColumnWidth(9,320);
    }
    let email = "";
    try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(se) { email = "sistema"; }
    const now   = new Date();
    const fecha = Utilities.formatDate(now,"GMT-6","dd/MM/yyyy");
    const hora  = Utilities.formatDate(now,"GMT-6","HH:mm:ss");
    const id    = "AUD-" + Utilities.formatDate(now,"GMT-6","yyyyMMdd-HHmmss");
    sh.appendRow([id, fecha, hora, email, email.split("@")[0].toUpperCase(), "", modulo || "SISTEMA", accion || "ACCIÓN", detalle || "", referencia || "", ""]);
  } catch(e) { Logger.log("AUDITORIA_ERR: " + e.message); }
}

function registrarSesion() {
  registrarAuditoria("SESIÓN", "INICIO_SESIÓN", "Usuario accedió al sistema", "—");
  return { ok: true };
}

function obtenerAuditoria(filtros) {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!_getSuperAdmins().includes(email)) return { error: "Sin permisos para: " + email };
    const ss = SpreadsheetApp.openById(SS_MASTER_ID);
    const sh = ss.getSheetByName(SH_AUDITORIA);
    if (!sh || sh.getLastRow() < 2) return { stats: { total:0, hoy:0, sesiones:0, solicitudes:0, cambios:0, usuarios:0 }, rows: [] };
    const numRows = sh.getLastRow() - 1;
    const cols    = Math.min(sh.getLastColumn(), 10);
    const raw     = sh.getRange(2, 1, numRows, cols).getDisplayValues();
    const hoy     = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy");
    const modulosSol = ["COMPRAS","ACTAS","MANTENIMIENTO","RESIDUOS","SOPORTE","LIMPIEZA"];
    let sesiones=0, solicitudes=0, cambios=0, hoyCount=0;
    const emails  = new Set();
    const rows    = [];
    for (let i = raw.length - 1; i >= 0; i--) {
      const r = raw[i];
      const fecha  = r[1] || "";
      const accion = r[7] || "";
      const modulo = r[6] || "";
      const em     = (r[3] || "").toLowerCase().trim();
      if (fecha === hoy) hoyCount++;
      if (accion.indexOf("SESIÓN") > -1 || accion.indexOf("SESION") > -1) sesiones++;
      if (modulosSol.indexOf(modulo) > -1 && accion.indexOf("NUEVA") > -1) solicitudes++;
      if (accion.indexOf("ESTATUS") > -1) cambios++;
      if (em) emails.add(em);
      if (rows.length < 500) {
        rows.push({ id:r[0]||"", fecha:fecha, hora:r[2]||"", email:r[3]||"", nombre:r[4]||"", area:r[5]||"", modulo:modulo, accion:accion, detalle:r[8]||"", referencia:r[9]||"" });
      }
    }
    return {
      stats: { total: raw.length, hoy: hoyCount, sesiones: sesiones, solicitudes: solicitudes, cambios: cambios, usuarios: emails.size },
      rows:  rows
    };
  } catch(e) { return { error: "GAS_ERROR: " + e.message }; }
}

function getAuditoriaStats() {
  try {
    const res = obtenerAuditoria({});
    if (res.error) return { ok: false, error: res.error };
    return { ok: true, stats: res.stats };
  } catch(e) { return { ok: false, error: e.message }; }
}

function _inicializarBase() {
  const cache = CacheService.getScriptCache();
  if (cache.get('_base_init')) return; 
  const ss = SpreadsheetApp.openById(SS_MASTER_ID);
  const esquemas = [
    { n: SH_USUARIOS, h: ["NE", "CORREO", "NOMBRE", "AREA", "PUESTO", "TIPO", "ACTIVO", "COMPRAS_ACCESO", "COMPRAS_ADMIN", "COMPRAS_AREA_RESPONSABLE", "ACTAS_ACCESO", "ACTAS_ADMIN", "ACTAS_PUEDE_SOLICITAR", "SOPORTE_ACCESO", "SOPORTE_ADMIN", "MANTENIMIENTO_ACCESO", "MANTENIMIENTO_ADMIN", "RESIDUOS_ACCESO", "RESIDUOS_ADMIN", "TRANSPORTE_ACCESO", "TRANSPORTE_ADMIN", "EXPEDIENTES_ACCESO", "EXPEDIENTES_ADMIN", "LIMPIEZA_ACCESO", "LIMPIEZA_ADMIN"] },
    { n: SH_EMPLEADOS, h: ["ID_EMPLEADO", "NOMBRE", "CORREO", "PUESTO", "AREA", "TIPO", "REGIMEN_FISCAL"] },
    { n: SH_COMPRAS, h: ["FOLIO_SOL", "FOLIO_LOTE", "AREA", "FECHA", "SOLICITANTE", "OP", "OP_ID", "AREA_OCUPA", "TIPO", "MATERIAL", "CARACTERISTICA", "ANCHO", "LARGO", "CALIBRE", "CANTIDAD", "UNIDAD", "DESCRIPCION", "FECHA_REQUERIDA", "CODIGO", "ESTATUS", "HORA", "MES", "AÑO", "OBSERVACION"] },
    { n: SH_ACTAS, h: ["FOLIO", "FECHA", "HORA", "SOLICITANTE_CORREO", "SOLICITANTE_NOMBRE", "ID_REPORTADO", "NOMBRE_REPORTADO", "PUESTO_REPORTADO", "AREA_REPORTADO", "TIPO_INCIDENCIA", "DESCRIPCION", "ESTATUS", "NOTIFICADO", "AREA_SOLICITANTE"] },
    { n: SH_SOPORTE, h: ["FOLIO", "FECHA", "HORA", "SOLICITANTE_CORREO", "SOLICITANTE_NOMBRE", "AREA_SOLICITANTE", "TIPO_SOLICITUD", "PRIORIDAD", "TITULO", "DESCRIPCION", "EQUIPO", "UBICACION", "ESTATUS", "FECHA_RESOLUCION"] },
    { n: "DB_MANTENIMIENTO", h: ["FOLIO", "FECHA_SOLICITUD", "HORA", "SOLICITANTE_CORREO", "SOLICITANTE_NOMBRE", "AREA_SOLICITANTE", "CATEGORIA", "ELEMENTO_A_MANTENER", "UBICACION_EXACTA", "TIPO_MANTENIMIENTO", "PRIORIDAD", "DESCRIPCION", "FECHA_SUGERIDA", "ESTATUS", "FECHA_RESOLUCION"] },
    { n: "DB_RESIDUOS", h: ["FOLIO", "PARTIDA", "FECHA", "HORA", "SOLICITANTE_CORREO", "SOLICITANTE_NOMBRE", "AREA", "TIPO_RESIDUO", "DESCRIPCION", "RECIPIENTE", "CANTIDAD", "PESO_TOTAL", "UNIDAD", "ESTATUS", "FECHA_RECOLECCION"] },
    { n: "DB_LIMPIEZA", h: ["FOLIO","FECHA","HORA","SOLICITANTE_CORREO","SOLICITANTE_NOMBRE","AREA_SOLICITANTE","NO_EQUIPO","NE_EQUIPO","USUARIO_EQUIPO","ARTICULOS","NOTAS","ESTATUS","FECHA_RESOLUCION","RESUELTO_POR","REEMPLAZADO","DETALLE_REEMPLAZO","OBSERVACIONES","FECHA_HORARIO"] },
    { n: SH_FOTOS, h: ["NE", "URL_FOTO"] },
    { n: SH_AUDITORIA, h: ["ID", "FECHA", "HORA", "EMAIL", "NOMBRE", "AREA", "MODULO", "ACCION", "DETALLE", "REFERENCIA", "IP_SESION"] }
  ];
  esquemas.forEach(e => { let sheet = ss.getSheetByName(e.n); if (!sheet) { sheet = ss.insertSheet(e.n); sheet.getRange(1, 1, 1, e.h.length).setValues([e.h]).setBackground("#5DC1E8").setFontColor("#FFFFFF").setFontWeight("bold"); sheet.setFrozenRows(1); } });
  cache.put('_base_init', '1', 1800);
}

function normalizarBooleano(valor) {
  if (valor === true || valor === 1) return true;
  const s = String(valor).toUpperCase().trim();
  return s === "SI" || s === "TRUE" || s === "YES" || s === "SÍ";
}

// Normaliza un área: MAYÚSCULAS + sin acentos → "LITOGRAFÍA" y "LITOGRAFIA" son iguales
function _normArea(s) {
  return String(s || "").toUpperCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function _getFotoUrl(ne) {
  try {
    if (!ne) return '';
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_FOTOS);
    if (!sh || sh.getLastRow() < 2) return '';
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const neStr = String(ne).trim();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === neStr) {
        const raw = String(data[i][1]).trim();
        if (!raw) return '';
        const m1 = raw.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
        const m2 = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
        const fileId = (m1 && m1[1]) || (m2 && m2[1]);
        if (fileId) return 'https://lh3.googleusercontent.com/d/' + fileId;
        return raw; 
      }
    }
    return '';
  } catch (e) { return ''; }
}

// 🚀 OPTIMIZACIÓN EXTREMA: Caché de Memoria RAM + Seguridad Estricta de Accesos
function obtenerPerfilCompleto(forceRefresh = false) {
  const email = Session.getActiveUser().getEmail().toLowerCase().trim();
  const cache = CacheService.getUserCache();
  const cacheKey = 'perfil_v4_' + email;

  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData); // Velocidad instantánea
  }

  _inicializarBase();
  const ss = SpreadsheetApp.openById(SS_MASTER_ID);
  const esSA = _getSuperAdmins().includes(email);
  const esEquipoSoporte = _getEquipoNotif("SOPORTE").includes(email);
  const dominiosOk = ["@ssolu.com.mx", "@bur.mx", "@bursatron.com.mx"];
  
  if (!(esSA || dominiosOk.some(d => email.endsWith(d)))) { 
    return { acceso: false, email, razon: "DOMINIO_NO_PERMITIDO" }; 
  }
  
  const shUsuarios = ss.getSheetByName(SH_USUARIOS);
  if (!shUsuarios || shUsuarios.getLastRow() < 2) { 
    return esSA ? crearPerfilSuperAdmin(email) : { acceso: false, email, razon: "BASE_DATOS_VACIA" }; 
  }
  
  const data = shUsuarios.getRange(2, 1, shUsuarios.getLastRow() - 1, 25).getValues();
  let u = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === email) {
      if (!normalizarBooleano(data[i][6])) return { acceso: false, email, razon: "USUARIO_INACTIVO" }; 
      u = { 
        numeroOperador: String(data[i][0]), correo: String(data[i][1]), nombre: String(data[i][2]), 
        area: String(data[i][3]).toUpperCase(), puesto: String(data[i][4]), tipo: String(data[i][5]).toUpperCase(), 
        activo: true, 
        compras_acceso: normalizarBooleano(data[i][7]), compras_admin: normalizarBooleano(data[i][8]), compras_area_responsable: String(data[i][9]).trim().toUpperCase(), 
        actas_acceso: normalizarBooleano(data[i][10]), actas_admin: normalizarBooleano(data[i][11]), actas_puede_solicitar: normalizarBooleano(data[i][12]), 
        soporte_acceso: normalizarBooleano(data[i][13]), soporte_admin_raw: String(data[i][14] || "").toUpperCase().trim(), 
        mantenimiento_acceso: normalizarBooleano(data[i][15]), mantenimiento_admin: normalizarBooleano(data[i][16]), 
        residuos_acceso: normalizarBooleano(data[i][17]), residuos_admin: normalizarBooleano(data[i][18]), 
        transporte_acceso: normalizarBooleano(data[i][19]), transporte_admin: normalizarBooleano(data[i][20]), 
        expedientes_acceso: normalizarBooleano(data[i][21]), expedientes_admin: normalizarBooleano(data[i][22]), 
        limpieza_acceso: normalizarBooleano(data[i][23]), limpieza_admin_raw: String(data[i][24] || "").toUpperCase().trim() 
      }; 
      break; 
    }
  }
  
  if (!u) { return esSA ? crearPerfilSuperAdmin(email) : { acceso: false, email, razon: "USUARIO_NO_REGISTRADO" }; }
  
  const esInfra = _getEquipoNotif("INFRA").includes(email);
  
  // REGLA DE HIERRO: Si acceso es falso, la interfaz gráfica bloquea/oculta la sección correspondiente.
    const _limpRaw = u.limpieza_admin_raw || "";
    const _limpEsAdmin = esSA || esInfra || (_limpRaw !== "" && _limpRaw !== "NO" && _limpRaw !== "FALSE");
    const _limpEsGlobal = esSA || esInfra || _limpRaw === "SI" || _limpRaw === "SÍ" || _limpRaw === "TRUE" || _limpRaw === "YES";
    const _limpAdminArea = (_limpEsAdmin && !_limpEsGlobal) ? _limpRaw : "";

    // Soporte: misma lógica que limpieza — SI=global, AREA_NAME=admin de esa área
    const _sopRaw = u.soporte_admin_raw || "";
    const _sopEsAdmin = esSA || esEquipoSoporte || (_sopRaw !== "" && _sopRaw !== "NO" && _sopRaw !== "FALSE");
    const _sopEsGlobal = esSA || esEquipoSoporte || _sopRaw === "SI" || _sopRaw === "SÍ" || _sopRaw === "TRUE" || _sopRaw === "YES";
    const _sopAdminArea = (_sopEsAdmin && !_sopEsGlobal) ? _sopRaw : "";

    const perfilFinal = { 
    acceso: true, fotoUrl: _getFotoUrl(u.numeroOperador), email: u.correo, numeroOperador: u.numeroOperador, nombre: u.nombre, 
    area: u.area, puesto: u.puesto, tipo: u.tipo, esSuperAdmin: esSA, esAdminUsuarios: esSA, 
    compras: { acceso: u.compras_acceso || esSA, admin: u.compras_admin || esSA, areaResponsable: u.compras_area_responsable, areaUsuario: u.area }, 
    actas: { acceso: u.actas_acceso || esSA, admin: u.actas_admin || esSA, puedeSolicitar: u.actas_puede_solicitar || esSA }, 
    soporte: { acceso: u.soporte_acceso || esSA || esEquipoSoporte, admin: _sopEsAdmin, adminArea: _sopAdminArea }, 
    mantenimiento: { acceso: u.mantenimiento_acceso || esSA, admin: u.mantenimiento_admin || esSA }, 
    residuos: { acceso: u.residuos_acceso || esSA, admin: u.residuos_admin || esSA }, 
    transporte: { acceso: u.transporte_acceso || esSA, admin: u.transporte_admin || esSA }, 
    expedientes: { acceso: u.expedientes_acceso || esSA, admin: u.expedientes_admin || esSA }, 
    limpieza: { acceso: true, admin: _limpEsAdmin, adminArea: _limpAdminArea } 
  };

  cache.put(cacheKey, JSON.stringify(perfilFinal), 300); // 5 minutos de memoria ultra rápida
  return perfilFinal;
}

function crearPerfilSuperAdmin(email) { return { acceso: true, fotoUrl: '', email: email, numeroOperador: "SA-999", nombre: email.split("@")[0].toUpperCase(), area: "DIRECCIÓN", puesto: "SUPER ADMIN", tipo: "ADMINISTRATIVO", esSuperAdmin: true, esAdminUsuarios: true, compras: { acceso: true, admin: true, areaResponsable: "", areaUsuario: "DIRECCIÓN" }, actas: { acceso: true, admin: true, puedeSolicitar: true }, soporte: { acceso: true, admin: true, adminArea: "" }, mantenimiento: { acceso: true, admin: true }, residuos: { acceso: true, admin: true }, transporte: { acceso: true, admin: true }, expedientes: { acceso: true, admin: true }, limpieza: { acceso: true, admin: true, adminArea: "" } }; }

function cargarAreasCompras() { try { const cache = CacheService.getScriptCache(); const cached = cache.get('areas_compras_v2'); if (cached) return JSON.parse(cached); const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS); if(!sh || sh.getLastRow()<2) return []; const rows = sh.getRange(2,1,sh.getLastRow()-1,4).getValues(); const areas = new Set(); rows.forEach(r => { const a = String(r[3]).trim().toUpperCase(); if(a && a !== "UNDEFINED") areas.add(a); }); const result = [...areas].sort(); try { cache.put('areas_compras_v2', JSON.stringify(result), 600); } catch(ce) {} return result; } catch(e) { return []; } }
function resolverJefeArea(area) { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS); const rows = sh.getDataRange().getValues(); for(let i=1; i<rows.length; i++) { if(String(rows[i][9]).trim().toUpperCase() === area.trim().toUpperCase()) { return { jefe: String(rows[i][2]), puesto: String(rows[i][4]), correo: String(rows[i][1]) }; } } return { error: true }; } catch(e) { return { error: true }; } }
function buscarEmpleadoPorId(id) { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_EMPLEADOS); const rows = sh.getDataRange().getValues(); for(let i=1; i<rows.length; i++) { if(String(rows[i][0]).trim() === String(id).trim()) { return { encontrado: true, nombre: String(rows[i][1]), puesto: String(rows[i][3]), area: String(rows[i][4]).toUpperCase(), tipo: String(rows[i][5]||"OPERADOR"), regimenFiscal: String(rows[i][6]||"NO FISCAL").toUpperCase().trim() }; } } return { encontrado: false }; } catch(e) { return { encontrado: false }; } }

/* MÓDULO 01 COMPRAS */
function calcularFolioCompra() { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_COMPRAS); if (!sh || sh.getLastRow() < 2) return { nextSol: 1, nextLote: 1 }; const data = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues(); let maxSol = 0, maxLote = 0; data.forEach(r => { const s = parseInt(String(r[0]).replace(/\D/g,"")); const l = parseInt(String(r[1]).replace(/\D/g,"")); if (!isNaN(s) && s > maxSol) maxSol = s; if (!isNaN(l) && l > maxLote) maxLote = l; }); return { nextSol: maxSol + 1, nextLote: maxLote + 1 }; } catch(e) { return { nextSol: 1, nextLote: 1 }; } }
function procesarSolicitudCompra(payload) { const lock = LockService.getScriptLock(); lock.tryLock(30000); try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_COMPRAS); const { nextSol, nextLote } = calcularFolioCompra(); const items = JSON.parse(payload.items); const now = new Date(); const rows = []; items.forEach((item, idx) => { rows.push([ nextSol, nextLote + idx, payload.cabecera.area, Utilities.formatDate(now,"GMT-6","dd/MM/yyyy"), payload.cabecera.solicitante, payload.cabecera.op||"N/A", payload.cabecera.opId||"0", payload.cabecera.areaOcupa||"", item.tipo, item.material, item.caracteristica||"", item.ancho||"", item.largo||"", item.calibre||"", item.cantidad, item.unidad, item.descripcion||"", item.fechaRequerida, item.codigoElemento||"", "SOLICITADA", Utilities.formatDate(now,"GMT-6","HH:mm:ss"), now.getMonth()+1, now.getFullYear() ]); }); sh.getRange(sh.getLastRow()+1, 1, rows.length, 23).setValues(rows); const correoFernanda = _getEquipoNotif("COMPRAS").filter(e => !_getSuperAdmins().includes(e))[0] || ""; 
        let jefeCorreoFinal = payload.cabecera.jefeCorreo || "";
        if (!jefeCorreoFinal && payload.cabecera.area) { try { const ji = resolverJefeArea(payload.cabecera.area); if (!ji.error && ji.correo) jefeCorreoFinal = ji.correo; } catch(je) {} }
        const destinos = [payload.emisorCorreo, jefeCorreoFinal, ..._getEquipoNotif("COMPRAS")].filter(e => e && e.includes("@")); let emailErrComp = ""; try { _mailCompra(destinos, nextSol, nextLote, items, payload.cabecera, correoFernanda); } catch(me) { emailErrComp = me.message; Logger.log("⚠️ Error email compra: " + me.message); } registrarAuditoria("COMPRAS","SOLICITUD_NUEVA","Area: "+payload.cabecera.area+", "+items.length+" partida(s), Material: "+items[0].material,"SOL-"+nextSol); return { success:true, folio:nextSol, loteInicio:nextLote, loteFin:nextLote+items.length-1, elementos:items.length, emailError: emailErrComp }; } catch(e) { return { success:false, msg:e.message }; } finally { lock.releaseLock(); } }
// [DEPRECATED] _getCorreoFernanda eliminada — usar _getEquipoNotif("COMPRAS") para obtener al equipo de compras.
function obtenerPedidosVisor(modo, areaUsuario, filtros) { try { const perfil = obtenerPerfilCompleto(); const esAdminReal = perfil.acceso && perfil.compras.admin; const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_COMPRAS); if (!sh || sh.getLastRow() < 2) return []; const anioActual = new Date().getFullYear(); let filas = sh.getDataRange().getDisplayValues().slice(1); if (!esAdminReal) { const a = String(perfil.compras.areaResponsable || perfil.area || "").trim().toUpperCase(); filas = filas.filter(r => String(r[2]).trim().toUpperCase() === a && parseInt(r[22]) === anioActual); } else if (modo === "MIS_PEDIDOS") { const a = String(areaUsuario||"").trim().toUpperCase(); filas = filas.filter(r => String(r[2]).trim().toUpperCase() === a && parseInt(r[22]) === anioActual); } if (filtros) { if (filtros.area && filtros.area !== "TODAS") filas = filas.filter(r => String(r[2]).toUpperCase() === filtros.area.toUpperCase()); if (filtros.estatus && filtros.estatus !== "TODOS") filas = filas.filter(r => String(r[19]).toUpperCase() === filtros.estatus.toUpperCase()); } filas.sort((a,b) => parseInt(b[1]) - parseInt(a[1])); return filas.slice(0, 500); } catch(e) { return []; } }
function actualizarEstatusCompra(folioLote, nuevoEstatus, observacion) { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_COMPRAS); const lastRow = sh.getLastRow(); if (lastRow < 2) return { success:false, msg:"Sin datos." }; const data = sh.getRange(2, 2, lastRow-1, 19).getValues(); for (let i = 0; i < data.length; i++) { if (String(data[i][0]).trim() === String(folioLote).trim()) { sh.getRange(i+2, 20).setValue(nuevoEstatus); if (observacion && observacion.trim() !== '') { sh.getRange(i+2, 24).setValue(observacion.trim()); } registrarAuditoria("COMPRAS","CAMBIO_ESTATUS","Lote: "+folioLote+" → "+nuevoEstatus+(observacion?" | Obs: "+observacion:""),folioLote); return { success:true }; } } return { success:false, msg:"Lote no encontrado." }; } catch(e) { return { success:false, msg:e.message }; } }
function _mailCompra(destinos, folio, loteInicio, items, cab, correoFernanda) { try { const count = items.length; const urlAutorizar = getWebAppUrl() + "?action=autorizar_compra&folio=" + folio + "&lote=" + loteInicio; const urlOmar = getWebAppUrl() + "?action=enviar_omar&folio=" + folio + "&lote=" + loteInicio; const urlIsela = getWebAppUrl() + "?action=enviar_isela&folio=" + folio + "&lote=" + loteInicio; const botonesFer = correoFernanda ? `<div style="display:flex;gap:10px;margin-top:20px"><a href="${urlAutorizar}" style="flex:1;background:#22C55E;color:#fff;text-align:center;padding:12px;border-radius:8px;text-decoration:none;font-weight:700;font-size:13px">YO AUTORIZO</a><a href="${urlOmar}" style="flex:1;background:#5DC1E8;color:#fff;text-align:center;padding:12px;border-radius:8px;text-decoration:none;font-weight:700;font-size:13px">→ ENVIAR A OMAR</a><a href="${urlIsela}" style="flex:1;background:#FF6B00;color:#fff;text-align:center;padding:12px;border-radius:8px;text-decoration:none;font-weight:700;font-size:13px">→ ENVIAR A ISELA</a></div>` : ''; const lista = items.slice(0,4).map((it,i) => `<tr style="border-bottom:1px solid #F1F5F9"><td style="padding:8px;font-size:12px">${String(i+1).padStart(2,'0')}</td><td style="padding:8px;font-size:12px;font-weight:600">${it.material}</td><td style="padding:8px;font-size:12px">${it.tipo}</td><td style="padding:8px;font-size:12px;font-weight:700;color:#FF6B00">${it.cantidad} ${it.unidad}</td></tr>`).join(""); const html = `<div style="font-family:Arial,sans-serif;background:#F7F9FC;padding:32px"><div style="max-width:600px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.1)"><div style="background:linear-gradient(135deg,#5DC1E8 0%,#3A9FD8 100%);padding:28px 32px;position:relative"><div style="position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#5DC1E8,#FF6B00)"></div><table width="100%"><tr valign="middle"><td><div style="font-size:10px;letter-spacing:3px;color:rgba(255,255,255,.7);font-weight:700;margin-bottom:8px">SANTILLÁN RAMÍREZ · SOLICITUD DE COMPRAS</div><div style="font-size:24px;font-weight:900;color:#fff">NUEVA SOLICITUD</div></td><td align="right"><div style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:12px;padding:12px 18px;text-align:center"><div style="font-size:9px;color:rgba(255,255,255,.7);letter-spacing:2px">FOLIO</div><div style="font-size:26px;color:#fff;font-weight:900;font-family:monospace">#${folio}</div></div></td></tr></table></div><div style="padding:28px 32px"><table width="100%" style="margin-bottom:20px"><tr><td width="50%" style="padding-right:8px"><div style="background:#F0F9FF;border-left:4px solid #5DC1E8;border-radius:8px;padding:14px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:4px">ÁREA</div><div style="font-size:16px;font-weight:800;color:#0F172A">${cab.area}</div></div></td><td width="50%" style="padding-left:8px"><div style="background:#FFF7ED;border-left:4px solid #FF6B00;border-radius:8px;padding:14px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:4px">RESPONSABLE</div><div style="font-size:16px;font-weight:800;color:#0F172A">${cab.solicitante}</div></div></td></tr></table><div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;margin-bottom:18px"><div style="background:#F8FAFC;padding:10px 14px;border-bottom:1px solid #E2E8F0"><span style="font-size:10px;letter-spacing:2px;color:#64748B;font-weight:700">PARTIDAS (${count} total)</span></div><table width="100%" cellpadding="0" cellspacing="0"><thead><tr style="background:#F1F5F9"><th style="padding:8px;font-size:9px;color:#64748B;text-align:left">#</th><th style="padding:8px;font-size:9px;color:#64748B;text-align:left">MATERIAL</th><th style="padding:8px;font-size:9px;color:#64748B;text-align:left">TIPO</th><th style="padding:8px;font-size:9px;color:#64748B;text-align:left">CANT.</th></tr></thead><tbody>${lista}</tbody></table></div>${botonesFer}</div></div></div>`; MailApp.sendEmail({ to: [...new Set(destinos)].join(","), name: "Sistema Corporativo · Compras", replyTo: _getReplyTo("COMPRAS"), subject: `Solicitud #${folio} · ${cab.area}`, htmlBody: html }); } catch(e) { throw new Error("EMAIL_COMPRA: " + e.message); } }

/* MÓDULO 02 ACTAS */
function registrarActa(payload) { const ss = SpreadsheetApp.openById(SS_MASTER_ID); if (!payload.puedeSolicitar) throw new Error("⛔ No tienes permisos."); const perfilSolicitante = obtenerPerfilCompleto(); const puestoJefe = perfilSolicitante.puesto.toUpperCase(); const areaJefe = perfilSolicitante.area.toUpperCase(); const esRH = areaJefe.includes("HUMANO") || areaJefe.includes("RH") || areaJefe.includes("CAPITAL"); const esSuperAdmin = perfilSolicitante.esSuperAdmin; const rangosAutoridad = ["JEFE", "JEFA", "LIDER", "LÍDER", "GERENTE", "DIRECTOR", "COORDINADOR", "SUPERVISOR", "ENCARGADO"]; if (!esRH && !esSuperAdmin && !rangosAutoridad.some(r => puestoJefe.includes(r))) throw new Error("⛔ JERARQUÍA INSUFICIENTE."); const de = ss.getSheetByName(SH_EMPLEADOS); const rows = de ? de.getDataRange().getValues() : []; let target = null; for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]).trim() === String(payload.idReportado).trim()) { target = { nombre: String(rows[i][1]), correo: String(rows[i][2]), puesto: String(rows[i][3]), area: String(rows[i][4]).toUpperCase(), tipo: String(rows[i][5]||"OPERADOR"), regimenFiscal: String(rows[i][6]||"NO FISCAL").toUpperCase().trim() }; break; } } if (!target) throw new Error("❌ ID no encontrado."); if (!esRH && !esSuperAdmin && areaJefe !== target.area) throw new Error(`⛔ FUERA DE JURISDICCIÓN.`); const dbActas = ss.getSheetByName(SH_ACTAS); const now = new Date(); const folio = "ACT-" + Utilities.formatDate(now,"GMT-6","yyyyMMdd") + "-" + String(dbActas.getLastRow()).padStart(4,"0"); const fecha = Utilities.formatDate(now,"GMT-6","dd/MM/yyyy"); const hora = Utilities.formatDate(now,"GMT-6","HH:mm"); const tipoActa = payload.tipoActa || target.regimenFiscal || "NO FISCAL"; dbActas.appendRow([ folio, fecha, hora, perfilSolicitante.email, perfilSolicitante.nombre, payload.idReportado, target.nombre, target.puesto, target.area, tipoActa, payload.descripcion, "SOLICITADA", "SI", areaJefe, payload.testigo1||"", payload.testigo2||"" ]); registrarAuditoria("ACTAS","ACTA_NUEVA","Colaborador: "+target.nombre+" ("+target.area+"), Tipo: "+tipoActa,folio); 
        const destinosActa = [...new Set([..._getEquipoNotif("CH"), perfilSolicitante.email])].filter(e => e && e.includes("@"));
        _mailActa([...new Set(destinosActa)], folio, fecha, hora, target, { solicitanteNombre: perfilSolicitante.nombre, areaSolicitante: areaJefe, descripcion: payload.descripcion }); return { folio, nombre: target.nombre, area: target.area }; }
function obtenerHistorialActas(modoAdmin, emailUsuario) { try { const ss = SpreadsheetApp.openById(SS_MASTER_ID); const email = Session.getActiveUser().getEmail().toLowerCase().trim(); let esAdminReal = false; if (modoAdmin) { if (_getSuperAdmins().includes(email)) { esAdminReal = true; } else { const shU = ss.getSheetByName(SH_USUARIOS); if (shU && shU.getLastRow() >= 2) { const uData = shU.getRange(2, 2, shU.getLastRow()-1, 11).getValues(); for (let i = 0; i < uData.length; i++) { if (String(uData[i][0]).trim().toLowerCase() === email) { esAdminReal = normalizarBooleano(uData[i][10]); break; } } } } } const sh = ss.getSheetByName(SH_ACTAS); if (!sh || sh.getLastRow() < 2) return []; const data = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues(); const result = []; const filtroEmail = (modoAdmin && esAdminReal) ? null : email; for (let i = data.length - 1; i >= 0; i--) { const r = data[i]; if (!String(r[0]).trim()) continue; if (filtroEmail && String(r[3]).toLowerCase().trim() !== filtroEmail) continue; let fechaStr = ""; try { fechaStr = (r[1] instanceof Date) ? Utilities.formatDate(r[1], "GMT-6", "dd/MM/yyyy") : String(r[1]); } catch(fe) { fechaStr = String(r[1]); } result.push({ folio: String(r[0]), fecha: fechaStr, hora: String(r[2]), solicitante: String(r[4]), idRep: String(r[5]), nombre: String(r[6]), puesto: String(r[7]), area: String(r[8]), tipo: String(r[9]), desc: String(r[10]), estatus: String(r[11]), areaSol: String(r[13]) }); } return result; } catch(e) { throw new Error("Error actas: " + e.message); } }
function obtenerActasPorEmpleado(idEmp) { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_ACTAS); if (!sh || sh.getLastRow() < 2) return []; return sh.getDataRange().getValues().slice(1).filter(r => String(r[5]).trim() === String(idEmp).trim() && String(r[0]).trim() !== "").map(r => ({ folio:r[0], fecha:r[1] instanceof Date ? Utilities.formatDate(r[1],"GMT-6","dd/MM/yyyy") : String(r[1]), tipo:r[9], desc:r[10], estatus:r[11] })).reverse(); } catch(e) { return []; } }
function actualizarEstatusActa(folio, nuevoEstatus) { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_ACTAS); const data = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues(); for (let i = 0; i < data.length; i++) { if (String(data[i][0]) === String(folio)) { sh.getRange(i+2,12).setValue(nuevoEstatus); if (nuevoEstatus === "RESUELTA") _mailActaResuelta(String(data[i][3]), String(data[i][4]), folio); registrarAuditoria("ACTAS","CAMBIO_ESTATUS","Folio: "+folio+" → "+nuevoEstatus,folio); return { success:true }; } } return { success:false, msg:"Folio no encontrado." }; } catch(e) { return { success:false, msg:e.message }; } }
function _mailActa(destinos, folio, fecha, hora, target, payload) { try { const html = `<div style="font-family:Arial;background:#F7F9FC;padding:32px"><div style="max-width:600px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.1)"><div style="background:linear-gradient(135deg,#FF6B00 0%,#FF8C00 100%);padding:28px 32px;position:relative"><div style="position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#FF6B00,#5DC1E8)"></div><table width="100%"><tr valign="middle"><td><div style="font-size:10px;letter-spacing:3px;color:rgba(255,255,255,.7);font-weight:700;margin-bottom:8px">ACTAS ADMINISTRATIVAS</div><div style="font-size:24px;font-weight:900;color:#fff">NUEVA ACTA</div></td><td align="right"><div style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:12px;padding:12px 16px;text-align:center"><div style="font-size:9px;color:rgba(255,255,255,.7);letter-spacing:2px">FOLIO</div><div style="font-size:12px;color:#fff;font-weight:900;font-family:monospace">${folio}</div></div></td></tr></table></div><div style="padding:28px 32px"><table width="100%" style="margin-bottom:18px"><tr><td width="50%" style="padding-right:8px"><div style="background:#FFF7ED;border-left:4px solid #FF6B00;border-radius:8px;padding:14px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:4px">REPORTADO</div><div style="font-size:15px;font-weight:800;color:#0F172A">${target.nombre}</div><div style="font-size:11px;color:#64748B;margin-top:2px">${target.puesto} · ${target.area}</div></div></td><td width="50%" style="padding-left:8px"><div style="background:#F0F9FF;border-left:4px solid #5DC1E8;border-radius:8px;padding:14px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:4px">SOLICITADO POR</div><div style="font-size:15px;font-weight:800;color:#0F172A">${payload.solicitanteNombre}</div></div></td></tr></table><div style="background:#FFF7ED;border-radius:10px;padding:14px;margin-top:16px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:6px;font-weight:700">MOTIVO</div><div style="font-size:13px;color:#0F172A;line-height:1.6">${payload.descripcion||"Sin descripcion"}</div></div><div style="background:#F0F9FF;border-radius:8px;padding:10px;margin-top:12px;text-align:center;font-size:11px;color:#64748B">Folio: <b style="font-family:monospace;color:#0F172A">${folio}</b> · ${fecha} ${hora}</div></div></div></div>`; MailApp.sendEmail({ to: destinos.join(","), name: "Sistema Corporativo · Capital Humano", replyTo: _getReplyTo("CH"), subject: `Acta ${folio} · ${target.nombre}`, htmlBody: html }); } catch(e) { throw new Error("EMAIL_ACTA: " + e.message); } }
function _mailActaResuelta(correo, nombre, folio) { try { const html = _emailShell({ headerGrad: "linear-gradient(135deg,#22C55E 0%,#16A34A 100%)", accent: "#22C55E", title: "ACTA RESUELTA", subtitle: "Capital Humano ha cerrado este caso", folioValue: folio, body: _emailInfoBlock("COLABORADOR", nombre, "#22C55E") + '<div style="background:#F0FDF4;border:1px solid #BBF7D0;border-radius:10px;padding:18px;margin-top:14px;text-align:center"><div style="font-size:28px;margin-bottom:8px">&#10003;</div><div style="font-size:14px;font-weight:700;color:#166534">Este caso ha sido resuelto satisfactoriamente</div><div style="font-size:12px;color:#4ADE80;margin-top:4px">No se requiere ninguna acción adicional.</div></div>', footerNote: "Este es un correo automático del sistema corporativo." }); MailApp.sendEmail({ to: correo, name: "Sistema Corporativo · Capital Humano", replyTo: _getReplyTo("CH"), subject: `Acta ${folio} resuelta`, htmlBody: html }); } catch(e) { throw new Error("EMAIL_ACTA_RESUELTA: " + e.message); } }

/* MÓDULO 03 SOPORTE */
function calcularFolioSoporte() { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_SOPORTE); const anio = new Date().getFullYear(); if (!sh || sh.getLastRow() < 2) return { nextFolio: "SOPORTE-"+anio+"-0001" }; const data = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues(); let maxNum = 0; data.forEach(r => { const f = String(r[0]); if (f.includes(String(anio))) { const num = parseInt(f.split("-").pop()); if (!isNaN(num) && num > maxNum) maxNum = num; } }); return { nextFolio: "SOPORTE-"+anio+"-" + String(maxNum + 1).padStart(4, "0") }; } catch(e) { return { nextFolio: "SOPORTE-"+new Date().getFullYear()+"-0001" }; } }
function registrarSoporte(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { folio: null, error: "Sistema ocupado, intenta de nuevo." };
  try {
    const ss = SpreadsheetApp.openById(SS_MASTER_ID);
    let dbSoporte = ss.getSheetByName(SH_SOPORTE);
    if (!dbSoporte) {
      dbSoporte = ss.insertSheet(SH_SOPORTE);
      const hdr = ["FOLIO","FECHA","HORA","SOLICITANTE_CORREO","SOLICITANTE_NOMBRE","AREA_SOLICITANTE","TIPO_SOLICITUD","PRIORIDAD","TITULO","DESCRIPCION","EQUIPO","UBICACION","ESTATUS","FECHA_RESOLUCION"];
      dbSoporte.getRange(1,1,1,hdr.length).setValues([hdr]).setBackground("#7C3AED").setFontColor("#fff").setFontWeight("bold");
      dbSoporte.setFrozenRows(1);
    }
    const { nextFolio } = calcularFolioSoporte();
    const now = new Date();
    const fecha = Utilities.formatDate(now, "GMT-6", "dd/MM/yyyy");
    const hora  = Utilities.formatDate(now, "GMT-6", "HH:mm");
    dbSoporte.appendRow([ nextFolio, fecha, hora, payload.solicitanteCorreo, payload.solicitanteNombre, payload.areaSolicitante, payload.tipoSolicitud, payload.prioridad, payload.titulo, payload.descripcion, payload.equipo||"", payload.ubicacion||"", "ABIERTA", "" ]);
    let emailErrSop = "";
    try { _mailSoporte(_getEquipoNotif("SOPORTE"), nextFolio, fecha, hora, payload); } catch(me) { emailErrSop = me.message; Logger.log("⚠️ Error email soporte: " + me.message); }
    registrarAuditoria("SOPORTE","SOLICITUD_NUEVA","Tipo: "+payload.tipoSolicitud+", Prioridad: "+payload.prioridad+", "+payload.titulo, nextFolio);
    return { folio: nextFolio, emailError: emailErrSop };
  } catch(e) { return { folio: null, error: e.message }; }
  finally { lock.releaseLock(); }
}
function obtenerHistorialSoporte(solicitudEsAdmin, emailUsuario) { try { const perfil = obtenerPerfilCompleto(); const esAdmin = perfil.acceso && perfil.soporte.admin; const adminArea = (perfil.soporte && perfil.soporte.adminArea) ? _normArea(perfil.soporte.adminArea) : ""; const mostrarTodo = esAdmin && !adminArea; const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_SOPORTE); if (!sh || sh.getLastRow() < 2) return []; const data = sh.getRange(2, 1, sh.getLastRow()-1, 14).getValues(); const result = []; for (let i = data.length-1; i >= 0; i--) { const r = data[i]; if (!mostrarTodo) { if (adminArea) { if (_normArea(r[5]) !== adminArea) continue; } else { if (String(r[3]).toLowerCase() !== perfil.email.toLowerCase()) continue; } } result.push({ folio: r[0], fecha: r[1] instanceof Date ? Utilities.formatDate(r[1], "GMT-6", "dd/MM/yyyy") : String(r[1]), hora: r[2], solicitante: r[4], area: r[5], tipo: r[6], prioridad: r[7], titulo: r[8], estatus: r[12] }); if (result.length >= 500) break; } return result; } catch(e) { return []; } }
function actualizarEstatusSoporte(folio, nuevoEstatus) { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_SOPORTE); const data = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues(); for (let i = 0; i < data.length; i++) { if (String(data[i][0]) === String(folio)) { sh.getRange(i+2, 13).setValue(nuevoEstatus); if (nuevoEstatus === "RESUELTA") { sh.getRange(i+2, 14).setValue(Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm")); _mailSoporteResuelta(String(data[i][3]), String(data[i][4]), folio); } registrarAuditoria("SOPORTE","CAMBIO_ESTATUS","Folio: "+folio+" -> "+nuevoEstatus,folio); return { success: true }; } } return { success: false, msg: "Folio no encontrado." }; } catch(e) { return { success: false, msg: e.message }; } }
function _mailSoporte(destinos, folio, fecha, hora, payload) { try { const urlResolver = getWebAppUrl() + "?action=resolver_soporte&folio=" + folio; const html = `<div style="font-family:Arial;background:#F7F9FC;padding:32px"><div style="max-width:600px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.1)"><div style="background:linear-gradient(135deg,#8B5CF6 0%,#7C3AED 100%);padding:28px 32px"><table width="100%"><tr valign="middle"><td><div style="font-size:10px;letter-spacing:3px;color:rgba(255,255,255,.7);font-weight:700;margin-bottom:8px">INFRAESTRUCTURA Y SOPORTE</div><div style="font-size:24px;font-weight:900;color:#fff">NUEVO TICKET</div></td><td align="right"><div style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:12px;padding:12px 16px;text-align:center"><div style="font-size:9px;color:rgba(255,255,255,.7);letter-spacing:2px">FOLIO</div><div style="font-size:11px;color:#fff;font-weight:900;font-family:monospace">${folio}</div></div></td></tr></table></div><div style="padding:28px 32px"><div style="background:#F5F3FF;border-left:4px solid #8B5CF6;border-radius:8px;padding:14px;margin-bottom:14px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:4px">SOLICITANTE</div><div style="font-size:15px;font-weight:800;color:#0F172A">${payload.solicitanteNombre}</div></div><a href="${urlResolver}" style="display:block;background:#8B5CF6;color:#fff;text-align:center;padding:14px;border-radius:8px;text-decoration:none;font-weight:700;margin-bottom:14px">MARCAR COMO RESUELTA</a></div></div></div>`; MailApp.sendEmail({ to: destinos.join(","), name: "Sistema Corporativo · Soporte TI", replyTo: _getReplyTo("SOPORTE"), subject: `Nuevo Ticket ${folio} · ${payload.titulo}`, htmlBody: html }); } catch(e) { throw new Error("EMAIL_SOPORTE: " + e.message); } }
function _mailSoporteResuelta(correo, nombre, folio) { try { const html = _emailShell({ headerGrad: "linear-gradient(135deg,#8B5CF6 0%,#7C3AED 100%)", accent: "#8B5CF6", title: "TICKET RESUELTO", subtitle: "Infraestructura y Soporte", folioValue: folio, body: _emailInfoBlock("SOLICITANTE", nombre, "#8B5CF6") + '<div style="background:#F5F3FF;border:1px solid #DDD6FE;border-radius:10px;padding:18px;margin-top:14px;text-align:center"><div style="font-size:28px;margin-bottom:8px">&#10003;</div><div style="font-size:14px;font-weight:700;color:#5B21B6">Tu ticket ha sido resuelto</div><div style="font-size:12px;color:#8B5CF6;margin-top:4px">Si el problema persiste, genera un nuevo ticket.</div></div>', footerNote: "Este es un correo automático del sistema corporativo." }); MailApp.sendEmail({ to: correo, name: "Sistema Corporativo · Soporte TI", replyTo: _getReplyTo("SOPORTE"), subject: `Ticket ${folio} resuelto`, htmlBody: html }); } catch(e) { throw new Error("EMAIL_SOPORTE_RESUELTA: " + e.message); } }

/* MANTENIMIENTO */
function calcularFolioMantenimiento() { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_MANTENIMIENTO"); const anio = new Date().getFullYear(); if (!sh || sh.getLastRow() < 2) return { nextFolio: "MANT-"+anio+"-0001" }; const data = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues(); let maxNum = 0; data.forEach(r => { const f = String(r[0]); if (f.includes(String(anio))) { const num = parseInt(f.split("-").pop()); if (!isNaN(num) && num > maxNum) maxNum = num; } }); return { nextFolio: "MANT-"+anio+"-" + String(maxNum + 1).padStart(4, "0") }; } catch(e) { return { nextFolio: "MANT-"+new Date().getFullYear()+"-0001" }; } }
function registrarMantenimiento(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { folio: null, error: "Sistema ocupado, intenta de nuevo." };
  try {
    const ss = SpreadsheetApp.openById(SS_MASTER_ID);
    const dbMant = ss.getSheetByName("DB_MANTENIMIENTO");
    const { nextFolio } = calcularFolioMantenimiento();
    const now = new Date();
    const fecha = Utilities.formatDate(now, "GMT-6", "dd/MM/yyyy");
    const hora  = Utilities.formatDate(now, "GMT-6", "HH:mm");
    dbMant.appendRow([ nextFolio, fecha, hora, payload.solicitanteCorreo, payload.solicitanteNombre, payload.areaSolicitante, payload.categoria, payload.elemento, payload.ubicacion, payload.tipoMantenimiento, payload.prioridad, payload.descripcion, payload.fechaSugerida, "PROGRAMADA", "" ]);
    let emailErrMant = "";
    try { _enviarAlertaMantenimiento(nextFolio, payload); } catch(me) { emailErrMant = me.message; Logger.log("⚠️ Error email mantenimiento: " + me.message); }
    registrarAuditoria("MANTENIMIENTO","SOLICITUD_NUEVA","Categoria: "+payload.categoria+", Elemento: "+payload.elemento+", Prioridad: "+payload.prioridad, nextFolio);
    return { folio: nextFolio, emailError: emailErrMant };
  } catch(e) { return { folio: null, error: e.message }; }
  finally { lock.releaseLock(); }
}
function obtenerHistorialMantenimiento(esAdmin) { try { const perfil = obtenerPerfilCompleto(); const mostrarTodo = perfil.acceso && perfil.mantenimiento.admin && esAdmin; const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_MANTENIMIENTO"); if (!sh || sh.getLastRow() < 2) return []; const data = sh.getRange(2, 1, sh.getLastRow()-1, 15).getValues(); const result = []; const filtroEmail = mostrarTodo ? null : perfil.email.toLowerCase(); for (let i = data.length-1; i >= 0; i--) { const r = data[i]; if (filtroEmail && String(r[3]).toLowerCase() !== filtroEmail) continue; result.push({ folio: r[0], fecha: r[1] instanceof Date ? Utilities.formatDate(r[1], "GMT-6", "dd/MM/yyyy") : String(r[1]), solicitante: r[4], categoria: r[6], elemento: r[7], tipo: r[9], prioridad: r[10], estatus: r[13] }); if (result.length >= 500) break; } return result; } catch(e) { return []; } }
function actualizarEstatusMantenimiento(folio, nuevoEstatus) { try { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_MANTENIMIENTO"); const data = sh.getRange(2, 1, sh.getLastRow()-1, 8).getValues(); for (let i = 0; i < data.length; i++) { if (String(data[i][0]) === String(folio)) { sh.getRange(i+2, 14).setValue(nuevoEstatus); if (nuevoEstatus === "RESUELTA" || nuevoEstatus === "TERMINADA") { sh.getRange(i+2, 15).setValue(Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm")); } else if (nuevoEstatus === "POR CONFIRMAR SOLICITANTE") { const correoSol = String(data[i][3]); const nombreSol = String(data[i][4]); const elemento = String(data[i][7]); try { _mailConfirmacionMantenimiento(correoSol, nombreSol, folio, elemento); } catch(me) {} } registrarAuditoria("MANTENIMIENTO","CAMBIO_ESTATUS","Folio: "+folio+" -> "+nuevoEstatus,folio); return { success: true }; } } return { success: false, msg: "Folio no encontrado." }; } catch(e) { return { success: false, msg: e.message }; } }
function _mailConfirmacionMantenimiento(correoSolicitante, nombreSolicitante, folio, elemento) { const urlConfirmar = getWebAppUrl() + "?action=confirmar_cierre_mant&folio=" + folio; const body = '<p style="font-size:14px;color:#334155;margin-bottom:14px">Estimado(a) <b style="color:#0F172A">' + nombreSolicitante + '</b>,</p>' + _emailInfoBlock("ELEMENTO ATENDIDO", elemento, "#22C55E") + '<div style="background:#F0FDF4;border:1px solid #BBF7D0;border-radius:10px;padding:16px;margin:16px 0"><div style="font-size:13px;color:#166534;font-weight:600;line-height:1.6">El equipo de Mantenimiento ha concluido el trabajo. Por favor confirme el cierre del ticket.</div></div>' + _emailButton(urlConfirmar, "CONFIRMAR CIERRE DEL TICKET", "#22C55E") + '<p style="font-size:11px;color:#94A3B8;margin-top:18px;text-align:center">Si el trabajo no fue satisfactorio, contacte al área de Mantenimiento.</p>'; const html = _emailShell({ headerGrad: "linear-gradient(135deg,#22C55E 0%,#16A34A 100%)", accent: "#22C55E", title: "TRABAJO COMPLETADO", subtitle: "Confirmación requerida", folioValue: folio, body: body }); MailApp.sendEmail({ to: correoSolicitante, name: "Sistema Corporativo · Mantenimiento", replyTo: _getReplyTo("MANT"), subject: `[Confirmación Requerida] Mantenimiento ${folio} - ${elemento}`, htmlBody: html }); }
function _enviarAlertaMantenimiento(folio, payload) { const destinos = _getEquipoNotif("MANT").join(","); const btnProceso = `${getWebAppUrl()}?action=updateMant&folio=${folio}&st=EN_PROCESO`; const btnResuelta = `${getWebAppUrl()}?action=updateMant&folio=${folio}&st=RESUELTA`; const prioColors = {"BAJA":"#54A9C8","MEDIA":"#D97706","ALTA":"#F0644F","CRÍTICA":"#DC2626"}; const prioColor = prioColors[payload.prioridad] || "#D97706"; const body = _emailInfoBlock("ELEMENTO", payload.elemento, "#22C55E") + '<table width="100%" cellpadding="0" cellspacing="0" style="margin:12px 0"><tr><td width="50%" style="padding-right:6px">' + _emailInfoBlock("PRIORIDAD", payload.prioridad, prioColor) + '</td><td width="50%" style="padding-left:6px">' + _emailInfoBlock("SOLICITANTE", payload.solicitanteNombre || "—", "#54A9C8") + '</td></tr></table>' + (payload.descripcion ? '<div style="background:#F8FAFC;border:1px solid #E8ECF0;border-radius:10px;padding:14px;margin:10px 0"><div style="font-family:monospace;font-size:8px;letter-spacing:2px;color:#94A3B8;margin-bottom:6px">DESCRIPCIÓN</div><div style="font-size:13px;color:#334155;line-height:1.6">' + payload.descripcion + '</div></div>' : '') + '<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px"><tr><td width="50%" style="padding-right:6px">' + _emailButton(btnProceso, "EN PROCESO", "#D97706") + '</td><td width="50%" style="padding-left:6px">' + _emailButton(btnResuelta, "RESUELTA", "#22C55E") + '</td></tr></table>'; const html = _emailShell({ headerGrad: "linear-gradient(135deg,#22C55E 0%,#16A34A 100%)", accent: "#22C55E", title: "NUEVO TICKET", subtitle: "Mantenimiento · " + (payload.categoria || "General"), folioValue: folio, body: body }); MailApp.sendEmail({ to: destinos, name: "Sistema Corporativo · Mantenimiento", replyTo: _getReplyTo("MANT"), subject: `[NUEVO] ${folio} - ${payload.elemento} (${payload.prioridad})`, htmlBody: html }); }
function enviarRecordatoriosMantenimiento() { const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_MANTENIMIENTO"); if (!sh || sh.getLastRow() < 2) return; const data = sh.getRange(2,1,sh.getLastRow()-1, 14).getValues(); const hoy = new Date(); hoy.setHours(0,0,0,0); const correos = _getEquipoNotif("MANT").join(","); let ticketsHoy = []; for(let i=0; i<data.length; i++) { const r = data[i]; const estatus = String(r[13]).toUpperCase(); if (estatus === "RESUELTA" || estatus === "TERMINADA" || estatus === "CANCELADA") continue; const fechaSug = r[12]; if (fechaSug) { const [yyyy, mm, dd] = String(fechaSug).split("-"); if (yyyy && mm && dd) { const fechaT = new Date(yyyy, mm-1, dd); if (fechaT.getTime() === hoy.getTime()) { ticketsHoy.push({ folio: r[0], elemento: r[7], ubicacion: r[8], solicitante: r[4] }); } } } } if (ticketsHoy.length > 0) { let items = ''; ticketsHoy.forEach(function(t){ items += '<div style="display:flex;align-items:center;gap:12px;padding:10px 14px;border-bottom:1px solid #F1F5F9"><div style="width:8px;height:8px;border-radius:50%;background:#D97706;flex-shrink:0"></div><div><div style="font-weight:700;font-size:13px;color:#0F172A">' + t.folio + ' — ' + t.elemento + '</div><div style="font-size:11px;color:#64748B">' + (t.solicitante||"") + '</div></div></div>'; }); const body = '<div style="background:#FFFBEB;border:1px solid #FDE68A;border-radius:10px;padding:16px;margin-bottom:16px;text-align:center"><div style="font-size:24px;margin-bottom:6px">&#9888;</div><div style="font-size:14px;font-weight:700;color:#92400E">Tienes ' + ticketsHoy.length + ' mantenimiento(s) programado(s) para hoy</div></div><div style="border:1px solid #E8ECF0;border-radius:10px;overflow:hidden">' + items + '</div>'; const html = _emailShell({ headerGrad: "linear-gradient(135deg,#D97706 0%,#B45309 100%)", accent: "#D97706", title: "RECORDATORIO", subtitle: "Mantenimientos programados para hoy", body: body }); MailApp.sendEmail({ to: correos, name: "Sistema Corporativo · Mantenimiento", replyTo: _getReplyTo("MANT"), subject: `[RECORDATORIO] ${ticketsHoy.length} mantenimientos programados para hoy`, htmlBody: html }); } }

/* ═══════════════════════════════════════════════════════════
   ☣️ MÓDULO 07: RESIDUOS PELIGROSOS 
═══════════════════════════════════════════════════════════ */

function calcularFolioResiduos() { 
  try { 
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_RESIDUOS"); 
    const anio = new Date().getFullYear();
    if (!sh || sh.getLastRow() < 2) return { nextFolio: "RESP-"+anio+"-0001" }; 
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues(); 
    let maxFolio = 0; 
    data.forEach(r => { 
      const f = String(r[0]); 
      if (f.includes(String(anio))) { 
        const num = parseInt(f.split("-").pop()); 
        if (!isNaN(num) && num > maxFolio) maxFolio = num; 
      } 
    }); 
    return { nextFolio: "RESP-"+anio+"-" + String(maxFolio + 1).padStart(4, "0") }; 
  } catch(e) { return { nextFolio: "RESP-"+new Date().getFullYear()+"-0001" }; } 
}

function procesarSolicitudResiduos(payload) { 
  const lock = LockService.getScriptLock(); 
  if (!lock.tryLock(20000)) return { success: false, msg: "Sistema ocupado, intenta de nuevo." }; 
  try { 
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_RESIDUOS"); 
    const { nextFolio } = calcularFolioResiduos(); 
    const now = new Date(); 
    const fecha = Utilities.formatDate(now,"GMT-6","dd/MM/yyyy"); 
    const hora = Utilities.formatDate(now,"GMT-6","HH:mm"); 
    
    const row = [ 
      nextFolio, "1", fecha, hora, payload.correo, payload.nombre, payload.area, payload.tipo, payload.descripcion, 
      "", "", "", "", "PENDIENTE", ""  
    ]; 
    
    sh.appendRow(row); 
    let emailErrRes = "";
    try { _enviarAlertaResiduos(nextFolio, payload.nombre, payload.area, payload.descripcion, payload.tipo); } catch(me){ emailErrRes = me.message; Logger.log("⚠️ Error email residuos: " + me.message); }
    return { success: true, folio: nextFolio, emailError: emailErrRes }; 
  } catch(e) { return { success: false, msg: e.message }; } finally { lock.releaseLock(); } 
}

function obtenerHistorialResiduos(esAdmin) { 
  try { 
    const perfil = obtenerPerfilCompleto(); 
    const mostrarTodo = perfil.acceso && perfil.residuos.admin && esAdmin; 
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_RESIDUOS"); 
    if (!sh || sh.getLastRow() < 2) return []; 
    
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 15).getValues(); 
    const result = []; 
    const filtroEmail = mostrarTodo ? null : perfil.email.toLowerCase(); 

    for (let i = data.length-1; i >= 0; i--) { 
      const r = data[i]; 
      if (filtroEmail && String(r[4]).toLowerCase() !== filtroEmail) continue; 
      result.push({ 
        folio: r[0], 
        fecha: r[2] instanceof Date ? Utilities.formatDate(r[2], "GMT-6", "dd/MM/yyyy") : String(r[2]), 
        solicitante: r[5], area: r[6], tipo: r[7], desc: r[8], recipiente: r[9], 
        cant: r[10], peso: r[11], uni: r[12], estatus: r[13] 
      }); 
      if (result.length >= 150) break; 
    } 
    return result; 
  } catch(e) { return []; } 
}

function actualizarEstatusResiduo(folio, nuevoEstatus) {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_RESIDUOS");
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(folio)) {
        sh.getRange(i+2, 14).setValue(nuevoEstatus);
        if (nuevoEstatus === "RECOLECTADO") {
          sh.getRange(i+2, 15).setValue(Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm"));
        }
        registrarAuditoria("RESIDUOS","CAMBIO_ESTATUS","Folio: "+folio+" -> "+nuevoEstatus,folio); return { success: true };
      }
    }
    return { success: false, msg: "Folio no encontrado." };
  } catch(e) { return { success: false, msg: e.message }; }
}

function _enviarAlertaResiduos(folio, solicitante, area, desc, tipo) { 
  const destinos = _getEquipoNotif("MANT").join(",");
  const urlPesar = `${getWebAppUrl()}?action=pesar_residuos&folio=${folio}`; 
  const body = '<table width="100%" cellpadding="0" cellspacing="0"><tr><td width="50%" style="padding-right:6px">' + _emailInfoBlock("ÁREA", area, "#EF4444") + '</td><td width="50%" style="padding-left:6px">' + _emailInfoBlock("GENERADOR", solicitante, "#54A9C8") + '</td></tr></table>' +
    '<div style="background:#FEF2F2;border:1px solid #FECACA;border-radius:10px;padding:16px;margin:12px 0"><div style="font-size:16px;font-weight:800;color:#991B1B;margin-bottom:4px">' + desc + '</div><div style="font-size:12px;color:#EF4444">Tipo: ' + tipo + '</div></div>' +
    '<div style="font-size:12px;color:#94A3B8;text-align:center;margin:14px 0">Da clic en el botón cuando recojas para ingresar el peso real:</div>' +
    _emailButton(urlPesar, "IR A PESAR Y RECIBIR", "#EF4444");
  const html = _emailShell({ headerGrad: "linear-gradient(135deg,#EF4444 0%,#B91C1C 100%)", accent: "#EF4444", title: "RECOLECCIÓN PENDIENTE", subtitle: "Residuos Peligrosos", folioValue: folio, body: body });
  MailApp.sendEmail({ to: destinos, name: "Sistema Corporativo · Residuos", replyTo: _getReplyTo("MANT"), subject: `[RESIDUOS] Recolectar en ${area} (Folio ${folio})`, htmlBody: html }); 
}

function guardarPesosResiduos(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: "Sistema ocupado, intenta de nuevo." };
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_RESIDUOS");
    if (!sh) return { ok: false, error: "Hoja DB_RESIDUOS no encontrada." };
    // Buscar fila por folio — NUNCA confiar en payload.row (inyección de fila)
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 14).getValues();
    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(payload.folio).trim()) { targetRow = i + 2; break; }
    }
    if (targetRow === -1) return { ok: false, error: "Folio no encontrado: " + payload.folio };
    const currentDesc = sh.getRange(targetRow, 9).getValue();
    const newDesc = payload.notas ? (currentDesc + " | [Detalle Pesos: " + payload.notas + "]") : currentDesc;
    sh.getRange(targetRow, 9).setValue(newDesc);
    sh.getRange(targetRow, 10).setValue(payload.recipiente);
    sh.getRange(targetRow, 11).setValue(payload.cantidad);
    sh.getRange(targetRow, 12).setValue(payload.pesoTotal);
    sh.getRange(targetRow, 13).setValue(payload.unidad);
    sh.getRange(targetRow, 14).setValue("EN ALMACÉN");
    const btnProveedor = `${getWebAppUrl()}?action=entregar_proveedor&folio=${payload.folio}`;
    const html2 = _emailShell({ headerGrad: "linear-gradient(135deg,#D97706 0%,#B45309 100%)", accent: "#D97706", title: "RESIDUOS EN ALMACÉN", subtitle: "Pesado y resguardado", folioValue: payload.folio, body: _emailInfoBlock("PESO TOTAL", payload.pesoTotal + " " + payload.unidad, "#D97706") + _emailButton(btnProveedor, "MARCAR COMO ENTREGADO A PROVEEDOR", "#22C55E") });
    MailApp.sendEmail({ to: _getEquipoNotif("MANT").join(","), name: "Sistema Corporativo · Residuos", replyTo: _getReplyTo("MANT"), subject: `[EN ALMACÉN] Folio ${payload.folio}`, htmlBody: html2 });
    registrarAuditoria("RESIDUOS","PESAJE_REGISTRADO","Folio: "+payload.folio+", Peso: "+payload.pesoTotal+" "+payload.unidad, payload.folio);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
  finally { lock.releaseLock(); }
}

/* ═══════════════════════════════════════════════════════════
   MÓDULO 08: EXPEDIENTE DIGITAL
═══════════════════════════════════════════════════════════ */

function buscarExpedientePorNE(ne) {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_EXPEDIENTES);
    if (!sh || sh.getLastRow() < 2) return null;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 50).getValues();
    const neStr = String(ne).trim();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === neStr) {
        const r = data[i];
        const fd = (v, fmt) => { try { return v instanceof Date ? Utilities.formatDate(v, "GMT-6", fmt || "yyyy-MM-dd") : String(v || ""); } catch(e) { return String(v || ""); } };
        return {
          rowIndex: i + 2, ne: String(r[0]), nombreCompleto: String(r[1]), fotoUrl: String(r[2]), curp: String(r[3]), rfc: String(r[4]), edad: String(r[5]), sexo: String(r[6]), estadoCivil: String(r[7]), nacionalidad: String(r[8]), cumpleanos: fd(r[9]), domicilio: String(r[10]), delegacionMunicipio: String(r[11]), correoElectronico: String(r[12]), telEmergencia: String(r[13]), telEmergencia2: String(r[14]), area: String(r[15]), puesto: String(r[16]), centroCosto: String(r[17]), nivelAcademico: String(r[18]), tipoContrato: String(r[19]), fechaIngreso: fd(r[20]), fechaBaja: fd(r[21]), bajasAnteriores: String(r[22]), numCic: String(r[23]), numImss: String(r[24]), tipoSangre: String(r[25]), padecimientos: String(r[26]), tallaFaja: String(r[27]), tallasUniforme: String(r[28]), beneficiario: String(r[29]), beneficiarioParentesco: String(r[30]), banco: String(r[31]), ctaBancaria: String(r[32]), claveInterbancaria: String(r[33]), archivoIne: String(r[34]), archivoRfc: String(r[35]), archivoCurp: String(r[36]), archivoDomicilio: String(r[37]), archivoContrato: String(r[38]), fechaRegistro: fd(r[39], "dd/MM/yyyy"), ultimaActualizacion: fd(r[40], "dd/MM/yyyy"), registradoPor: String(r[41]), estatus: String(r[42]), sueldoBase: String(r[43]), fiscal: String(r[44]), complemento: String(r[45]), sueldoMensualTotal: String(r[46]), salarioIntegradoDiario: String(r[47]), valesDespensa: String(r[48] || ''), idCiudadano: String(r[49] || '')
        };
      }
    }
    return null;
  } catch (e) { throw new Error("Error buscando expediente: " + e.message); }
}

function guardarExpediente(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(30000);
    const ss = SpreadsheetApp.openById(SS_MASTER_ID);
    const sh = ss.getSheetByName(SH_EXPEDIENTES);
    const now = new Date();
    const fechaHoy = Utilities.formatDate(now, "GMT-6", "dd/MM/yyyy");
    const email = Session.getActiveUser().getEmail();
    const neStr = String(payload.ne).trim();
    let rowIdx = 0; let existing = new Array(50).fill("");
    if (sh.getLastRow() >= 2) {
      const data = sh.getRange(2, 1, sh.getLastRow() - 1, 50).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === neStr) { rowIdx = i + 2; existing = data[i].map(v => v instanceof Date ? Utilities.formatDate(v, "GMT-6", "dd/MM/yyyy") : String(v || "")); break; }
      }
    }
    const v = (key, idx) => (payload[key] !== undefined && payload[key] !== null && String(payload[key]).trim() !== "") ? String(payload[key]).trim() : existing[idx];
    const row = [ neStr, v("nombreCompleto", 1), v("fotoUrl", 2), v("curp", 3), v("rfc", 4), v("edad", 5), v("sexo", 6), v("estadoCivil", 7), v("nacionalidad", 8), v("cumpleanos", 9), v("domicilio", 10), v("delegacionMunicipio", 11), v("correoElectronico", 12), v("telEmergencia", 13), v("telEmergencia2", 14), v("area", 15), v("puesto", 16), v("centroCosto", 17), v("nivelAcademico", 18), v("tipoContrato", 19), v("fechaIngreso", 20), v("fechaBaja", 21), v("bajasAnteriores", 22), v("numCic", 23), v("numImss", 24), v("tipoSangre", 25), v("padecimientos", 26), v("tallaFaja", 27), v("tallasUniforme", 28), v("beneficiario", 29), v("beneficiarioParentesco", 30), v("banco", 31), v("ctaBancaria", 32), v("claveInterbancaria", 33), v("archivoIne", 34), v("archivoRfc", 35), v("archivoCurp", 36), v("archivoDomicilio", 37), v("archivoContrato", 38), rowIdx ? existing[39] : fechaHoy, fechaHoy, email, v("estatus", 42) || "PENDIENTE", v("sueldoBase", 43), v("fiscal", 44), v("complemento", 45), v("sueldoMensualTotal", 46), v("salarioIntegradoDiario", 47), v("valesDespensa", 48), v("idCiudadano", 49) ];
    if (rowIdx) { sh.getRange(rowIdx, 1, 1, 50).setValues([row]); } else { sh.appendRow(row); }
    return { success: true, mensaje: rowIdx ? "Expediente actualizado" : "Expediente creado", esNuevo: !rowIdx };
  } catch (e) { return { success: false, mensaje: "Error: " + e.message }; }
  finally { try { lock.releaseLock(); } catch(le) {} }
}

function obtenerTodosExpedientes() {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_EXPEDIENTES);
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 50).getValues();
    return data.map((r, i) => {
      const fd = (v) => { try { return v instanceof Date ? Utilities.formatDate(v, "GMT-6", "dd/MM/yyyy") : String(v || ""); } catch(e) { return String(v || ""); } };
      return { rowIndex: i + 2, ne: String(r[0]), nombreCompleto: String(r[1]), fotoUrl: String(r[2] || ""), area: String(r[15]), puesto: String(r[16]), correoElectronico: String(r[12]), fechaIngreso: fd(r[20]), fechaBaja: fd(r[21]), tipoContrato: String(r[19]), estatus: String(r[42]), tieneIne: !!String(r[34]).trim(), tieneRfc: !!String(r[35]).trim(), tieneCurp: !!String(r[36]).trim(), tieneDomicilio: !!String(r[37]).trim(), tieneContrato: !!String(r[38]).trim(), urlIne: String(r[34] || ""), urlRfc: String(r[35] || ""), urlCurp: String(r[36] || ""), urlDomicilio: String(r[37] || ""), urlContrato: String(r[38] || ""), fechaRegistro: fd(r[39]), ultimaActualizacion: fd(r[40]), sueldoBase: String(r[43] || ""), fiscal: String(r[44] || ""), salarioIntegradoDiario: String(r[47] || ""), valesDespensa: String(r[48] || "") };
    });
  } catch (e) { throw new Error("Error cargando expedientes: " + e.message); }
}

function subirArchivoExpediente(base64Data, mimeType, nombreArchivo, ne, nombreCompleto, colArchivo) {
  try {
    const parentFolder = DriveApp.getFolderById(DRIVE_EXP_ID);
    const nombreCarpeta = String(nombreCompleto).trim().toUpperCase();
    let folder; const folders = parentFolder.getFoldersByName(nombreCarpeta);
    if (folders.hasNext()) { folder = folders.next(); } else { folder = parentFolder.createFolder(nombreCarpeta); }
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, nombreArchivo);
    const file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();
    const colMap = { archivoIne: 35, archivoRfc: 36, archivoCurp: 37, archivoDomicilio: 38, archivoContrato: 39 };
    const colNum = colMap[colArchivo];
    if (colNum) {
      const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_EXPEDIENTES);
      if (sh.getLastRow() >= 2) {
        const nes = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < nes.length; i++) { if (String(nes[i][0]).trim() === String(ne).trim()) { sh.getRange(i + 2, colNum).setValue(url); break; } }
      }
    }
    return { success: true, url: url, mensaje: "Archivo subido correctamente" };
  } catch (e) { return { success: false, mensaje: "Error subiendo archivo: " + e.message }; }
}

function subirFotoExpediente(base64Data, mimeType, nombreArchivo, ne, nombreCompleto) {
  try {
    const parentFolder = DriveApp.getFolderById(DRIVE_EXP_ID);
    const nombreCarpeta = String(nombreCompleto).trim().toUpperCase();
    let folder; const folders = parentFolder.getFoldersByName(nombreCarpeta);
    if (folders.hasNext()) { folder = folders.next(); } else { folder = parentFolder.createFolder(nombreCarpeta); }
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, "FOTO_" + ne + "." + mimeType.split("/")[1]);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_EXPEDIENTES);
    if (sh.getLastRow() >= 2) {
      const nes = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < nes.length; i++) {
        if (String(nes[i][0]).trim() === String(ne).trim()) { sh.getRange(i + 2, 3).setValue(url); break; }
      }
    }
    return { success: true, url: url };
  } catch(e) { return { success: false, mensaje: "Error subiendo foto: " + e.message }; }
}

function crearExpedienteNuevo(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(30000);
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_EXPEDIENTES);
    const now = new Date();
    const fechaHoy = Utilities.formatDate(now, "GMT-6", "dd/MM/yyyy");
    const email = Session.getActiveUser().getEmail();
    const neStr = String(payload.ne || "").trim().toUpperCase();
    if (!neStr) return { success: false, mensaje: "NE es requerido" };
    // Verificar que no exista
    if (sh.getLastRow() >= 2) {
      const nes = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
      for (let i = 0; i < nes.length; i++) {
        if (String(nes[i][0]).trim().toUpperCase() === neStr) return { success: false, mensaje: "Ya existe un expediente con NE: " + neStr };
      }
    }
    const row = new Array(50).fill("");
    row[0]  = neStr;
    row[1]  = String(payload.nombreCompleto || "").trim().toUpperCase();
    row[3]  = String(payload.curp || "").trim().toUpperCase();
    row[4]  = String(payload.rfc || "").trim().toUpperCase();
    row[5]  = String(payload.edad || "");
    row[6]  = String(payload.sexo || "");
    row[7]  = String(payload.estadoCivil || "");
    row[8]  = String(payload.nacionalidad || "MEXICANA").toUpperCase();
    row[9]  = String(payload.cumpleanos || "");
    row[10] = String(payload.domicilio || "");
    row[11] = String(payload.delegacionMunicipio || "");
    row[12] = String(payload.correoElectronico || "");
    row[13] = String(payload.telEmergencia || "");
    row[14] = String(payload.telEmergencia2 || "");
    row[15] = String(payload.area || "").toUpperCase();
    row[16] = String(payload.puesto || "").toUpperCase();
    row[17] = String(payload.centroCosto || "");
    row[18] = String(payload.nivelAcademico || "");
    row[19] = String(payload.tipoContrato || "INDETERMINADO");
    row[20] = String(payload.fechaIngreso || "");
    row[21] = "";
    row[22] = "0";
    row[23] = String(payload.numCic || "");
    row[24] = String(payload.numImss || "");
    row[25] = String(payload.tipoSangre || "");
    row[26] = String(payload.padecimientos || "");
    row[27] = String(payload.tallaFaja || "");
    row[28] = String(payload.tallasUniforme || "");
    row[29] = String(payload.beneficiario || "");
    row[30] = String(payload.beneficiarioParentesco || "");
    row[31] = String(payload.banco || "").toUpperCase();
    row[32] = String(payload.ctaBancaria || "");
    row[33] = String(payload.claveInterbancaria || "");
    row[39] = fechaHoy;
    row[40] = fechaHoy;
    row[41] = email;
    row[42] = "PENDIENTE";
    row[43] = String(payload.sueldoBase || "");
    row[44] = String(payload.fiscal || payload.sueldoBase || "");
    row[45] = String(payload.complemento || "");
    row[46] = String(payload.sueldoMensualTotal || "");
    row[47] = String(payload.salarioIntegradoDiario || "");
    row[48] = String(payload.valesDespensa || "");
    row[49] = String(payload.idCiudadano || "");
    sh.appendRow(row);
    registrarAuditoria("EXPEDIENTES","EXPEDIENTE_CREADO","NE: "+neStr+" | "+row[1],neStr);
    return { success: true, mensaje: "Expediente creado: " + neStr, ne: neStr };
  } catch(e) { return { success: false, mensaje: "Error: " + e.message }; }
  finally { try { lock.releaseLock(); } catch(le) {} }
}

/* ═══════════════════════════════════════════════════════════
   ADMINISTRACIÓN DE USUARIOS
═══════════════════════════════════════════════════════════ */

function obtenerTodosLosUsuarios() { 
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS); 
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 25).getValues(); 
    const usuarios = []; 
    data.forEach((r, i) => { 
      usuarios.push({ 
        rowIndex: i + 2, numeroOperador: String(r[0]), correo: String(r[1]).toLowerCase().trim(), nombre: String(r[2]), area: String(r[3]).toUpperCase(), puesto: String(r[4]), tipo: String(r[5]).toUpperCase(), activo: normalizarBooleano(r[6]), 
        compras_acceso: normalizarBooleano(r[7]), compras_admin: normalizarBooleano(r[8]), compras_area_responsable: String(r[9]).toUpperCase(), 
        actas_acceso: normalizarBooleano(r[10]), actas_admin: normalizarBooleano(r[11]), actas_puede_solicitar: normalizarBooleano(r[12]), 
        soporte_acceso: normalizarBooleano(r[13]), soporte_admin: normalizarBooleano(r[14]), soporte_admin_raw: String(r[14] || "").toUpperCase().trim(), 
        mantenimiento_acceso: normalizarBooleano(r[15]), mantenimiento_admin: normalizarBooleano(r[16]), 
        residuos_acceso: normalizarBooleano(r[17]), residuos_admin: normalizarBooleano(r[18]),
        transporte_acceso: normalizarBooleano(r[19]||false), transporte_admin: normalizarBooleano(r[20]||false),
        expedientes_acceso: normalizarBooleano(r[21]||false), expedientes_admin: normalizarBooleano(r[22]||false),
        limpieza_acceso: normalizarBooleano(r[23]||false), limpieza_admin: normalizarBooleano(r[24]||false), limpieza_admin_raw: String(r[24] || "").toUpperCase().trim()
      }); 
    }); 
    return usuarios;
  } catch(e) { throw new Error("Error cargando usuarios: " + e.message); }
}

function crearNuevoUsuario(payload) { 
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
    if (sh.getLastRow() >= 2) {
      const correos = sh.getRange(2, 2, sh.getLastRow()-1, 1).getValues();
      for (let i = 0; i < correos.length; i++) {
        if (String(correos[i][0]).toLowerCase().trim() === payload.correo.toLowerCase().trim()) { return { success: false, mensaje: "Error: Ya existe un usuario con ese correo" }; }
      }
    }
    const row = [ 
      payload.numeroOperador, payload.correo, payload.nombre, payload.area, payload.puesto, payload.tipo, "SI", 
      payload.compras_acceso?"SI":"NO", payload.compras_admin?"SI":"NO", payload.compras_area_responsable||"", 
      payload.actas_acceso?"SI":"NO", payload.actas_admin?"SI":"NO", payload.actas_puede_solicitar?"SI":"NO", 
      payload.soporte_acceso?"SI":"NO", payload.soporte_admin||"NO", 
      payload.mantenimiento_acceso?"SI":"NO", payload.mantenimiento_admin?"SI":"NO", 
      payload.residuos_acceso?"SI":"NO", payload.residuos_admin?"SI":"NO",
      payload.transporte_acceso?"SI":"NO", payload.transporte_admin?"SI":"NO",
      payload.expedientes_acceso?"SI":"NO", payload.expedientes_admin?"SI":"NO",
      payload.limpieza_acceso?"SI":"NO", payload.limpieza_admin||"NO"
    ]; 
    sh.appendRow(row); 
    CacheService.getUserCache().remove('perfil_v4_' + payload.correo.toLowerCase().trim());
    registrarAuditoria("USUARIOS","USUARIO_CREADO","Correo: "+payload.correo+", Area: "+payload.area+", Tipo: "+payload.tipo,""); return { success: true, mensaje: "Usuario creado exitosamente" };
  } catch(e) { return { success: false, mensaje: "Error: " + e.message }; }
}

function actualizarUsuario(payload) { 
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
    const row = [
      payload.numeroOperador, payload.correo, payload.nombre, payload.area, payload.puesto, payload.tipo, 
      payload.activo ? "SI" : "NO",
      payload.compras_acceso ? "SI" : "NO", payload.compras_admin ? "SI" : "NO", payload.compras_area_responsable || "",
      payload.actas_acceso ? "SI" : "NO", payload.actas_admin ? "SI" : "NO", payload.actas_puede_solicitar ? "SI" : "NO",
      payload.soporte_acceso ? "SI" : "NO", payload.soporte_admin || "NO",
      payload.mantenimiento_acceso ? "SI" : "NO", payload.mantenimiento_admin ? "SI" : "NO",
      payload.residuos_acceso ? "SI" : "NO", payload.residuos_admin ? "SI" : "NO",
      payload.transporte_acceso ? "SI" : "NO", payload.transporte_admin ? "SI" : "NO",
      payload.expedientes_acceso ? "SI" : "NO", payload.expedientes_admin ? "SI" : "NO",
      payload.limpieza_acceso ? "SI" : "NO", payload.limpieza_admin || "NO"
    ];
    sh.getRange(payload.rowIndex, 1, 1, 25).setValues([row]);
    CacheService.getUserCache().remove('perfil_v4_' + payload.correo.toLowerCase().trim());
    registrarAuditoria("USUARIOS","USUARIO_ACTUALIZADO","Correo: "+payload.correo+", Area: "+payload.area+", Activo: "+(payload.activo?"SI":"NO"),"");
    return { success: true, mensaje: "Usuario actualizado" };
  } catch(e) { return { success: false, mensaje: "Error: " + e.message }; }
}

function desactivarUsuario(correo) { 
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS); 
    const data = sh.getRange(2, 2, sh.getLastRow()-1, 1).getValues(); 
    for (let i = 0; i < data.length; i++) { 
      if (String(data[i][0]).toLowerCase().trim() === correo.toLowerCase().trim()) { 
        sh.getRange(i+2, 7).setValue("NO"); 
        CacheService.getUserCache().remove('perfil_v4_' + correo.toLowerCase().trim());
        registrarAuditoria("USUARIOS","USUARIO_DESACTIVADO","Correo desactivado: "+correo,""); return { success: true, mensaje: "Usuario desactivado" }; 
      } 
    } 
    return { success: false, mensaje: "Usuario no encontrado" };
  } catch(e) { return { success: false, mensaje: "Error: " + e.message }; }
}

/* ═══════════════════════════════════════════════════════════
   MANEJADOR DE ENLACES EXTERNOS (CORREOS / HTML GET)
═══════════════════════════════════════════════════════════ */

function handleCallback(params) {
  const action = params.action; const folio = params.folio; const lote = params.lote;
  if (action === "autorizar_compra") { actualizarEstatusCompra(lote, "AUTORIZADA"); return HtmlService.createHtmlOutput(`<div style="font-family:Arial;text-align:center;padding:60px 20px"><h1 style="color:#22C55E">AUTORIZADO</h1><p>Lote #${lote} autorizado.</p></div>`); }
  if (action === "enviar_omar") { const urlOmar = getWebAppUrl() + "?action=confirmar_omar&folio=" + folio + "&lote=" + lote; const hmOmar = `<div style="font-family:Arial,sans-serif;background:#F7F9FC;padding:32px"><div style="max-width:580px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.1)"><div style="background:linear-gradient(135deg,#5DC1E8,#3A9FD8);padding:24px 28px"><div style="font-size:10px;letter-spacing:3px;color:rgba(255,255,255,.7);margin-bottom:6px">SANTILLÁN RAMÍREZ · COMPRAS</div><div style="font-size:22px;font-weight:900;color:#fff">AUTORIZACIÓN REQUERIDA</div></div><div style="padding:24px 28px"><div style="background:#F0F9FF;border-left:4px solid #5DC1E8;border-radius:8px;padding:14px;margin-bottom:20px"><div style="font-size:11px;color:#64748B;margin-bottom:4px">FOLIO / LOTE</div><div style="font-size:18px;font-weight:800;color:#0F172A;font-family:monospace">#${folio} / #${lote}</div></div><a href="${urlOmar}" style="display:block;background:linear-gradient(135deg,#22C55E,#16A34A);color:#fff;text-align:center;padding:16px;border-radius:10px;text-decoration:none;font-weight:700;font-size:15px;box-shadow:0 4px 14px rgba(34,197,94,.4)">✓ AUTORIZAR SOLICITUD DE COMPRA</a></div></div></div>`; MailApp.sendEmail({ to: _getEquipoNotif("COMPRAS_APROBADOR").find(e=>e.includes("omar")) || _getReplyTo("COMPRAS"), name: "Sistema Corporativo · Compras", replyTo: _getReplyTo("COMPRAS"), subject: "Autorización Requerida · Folio #" + folio, htmlBody: hmOmar }); return HtmlService.createHtmlOutput('<div style="font-family:Arial;text-align:center;padding:60px 20px"><h2 style="color:#22C55E">✓ Enviado a Omar</h2><p>El correo de autorización fue enviado exitosamente.</p></div>'); }
  if (action === "enviar_isela") { const urlIsela = getWebAppUrl() + "?action=confirmar_isela&folio=" + folio + "&lote=" + lote; const hmIsela = `<div style="font-family:Arial,sans-serif;background:#F7F9FC;padding:32px"><div style="max-width:580px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.1)"><div style="background:linear-gradient(135deg,#FF6B00,#D97706);padding:24px 28px"><div style="font-size:10px;letter-spacing:3px;color:rgba(255,255,255,.7);margin-bottom:6px">SANTILLÁN RAMÍREZ · COMPRAS</div><div style="font-size:22px;font-weight:900;color:#fff">AUTORIZACIÓN REQUERIDA</div></div><div style="padding:24px 28px"><div style="background:#FFF7ED;border-left:4px solid #FF6B00;border-radius:8px;padding:14px;margin-bottom:20px"><div style="font-size:11px;color:#64748B;margin-bottom:4px">FOLIO / LOTE</div><div style="font-size:18px;font-weight:800;color:#0F172A;font-family:monospace">#${folio} / #${lote}</div></div><a href="${urlIsela}" style="display:block;background:linear-gradient(135deg,#FF6B00,#D97706);color:#fff;text-align:center;padding:16px;border-radius:10px;text-decoration:none;font-weight:700;font-size:15px;box-shadow:0 4px 14px rgba(255,107,0,.4)">✓ AUTORIZAR SOLICITUD DE COMPRA</a></div></div></div>`; MailApp.sendEmail({ to: _getEquipoNotif("COMPRAS_APROBADOR").find(e=>e.includes("isela")) || _getReplyTo("COMPRAS"), name: "Sistema Corporativo · Compras", replyTo: _getReplyTo("COMPRAS"), subject: "Autorización Requerida · Folio #" + folio, htmlBody: hmIsela }); return HtmlService.createHtmlOutput('<div style="font-family:Arial;text-align:center;padding:60px 20px"><h2 style="color:#22C55E">✓ Enviado a Isela</h2><p>El correo de autorización fue enviado exitosamente.</p></div>'); }
  if (action === "confirmar_omar") { actualizarEstatusCompra(lote, "AUTORIZADA POR OMAR"); return HtmlService.createHtmlOutput("Autorizado por Omar"); }
  if (action === "confirmar_isela") { actualizarEstatusCompra(lote, "AUTORIZADA POR ISELA"); return HtmlService.createHtmlOutput("Autorizado por Isela"); }
  if (action === "resolver_soporte") { actualizarEstatusSoporte(folio, "RESUELTA"); return HtmlService.createHtmlOutput(`<div style="font-family:Arial;text-align:center;padding:60px 20px"><h1 style="color:#22C55E">TICKET RESUELTO</h1></div>`); }
  if (action === 'confirmar_cierre_mant') { actualizarEstatusMantenimiento(folio, 'RESUELTA'); return HtmlService.createHtmlOutput(`<div style="font-family:Arial,sans-serif;text-align:center;padding:60px 20px;background:#F7F9FC;min-height:100vh;display:flex;align-items:center;justify-content:center"><div style="background:#fff;border-radius:16px;padding:48px 40px;box-shadow:0 8px 32px rgba(0,0,0,.12);max-width:480px"><div style="width:72px;height:72px;border-radius:50%;background:rgba(34,197,94,.12);border:3px solid rgba(34,197,94,.3);display:flex;align-items:center;justify-content:center;font-size:2rem;margin:0 auto 20px;color:#22C55E"><i class='fa-solid fa-circle-check'></i></div><h1 style="color:#22C55E;font-family:sans-serif">TICKET CERRADO</h1><p style="font-size:16px;color:#475569">El folio <b>${folio}</b> ha sido marcado como <b style='color:#22C55E'>RESUELTA</b>. Gracias por confirmar.</p></div></div>`); }
  if (action === 'updateMant') { const st = params.st === 'EN_PROCESO' ? 'EN PROCESO' : 'RESUELTA'; actualizarEstatusMantenimiento(folio, st); const color = st === 'RESUELTA' ? '#22C55E' : '#F59E0B'; return HtmlService.createHtmlOutput(`<div style="font-family:sans-serif;text-align:center;padding:50px;background:#F7F9FC;height:100vh;"><div style="background:white;padding:40px;border-radius:15px;box-shadow:0 10px 30px rgba(0,0,0,0.1);max-width:500px;margin:0 auto;"><h1 style="color:${color};margin-top:0;">TICKET ACTUALIZADO</h1><p style="font-size:18px;">El folio <b>${folio}</b> ha cambiado a: <b style="color:${color}">${st}</b>.</p></div></div>`); }

  if (action === 'pesar_residuos') {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_RESIDUOS");
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 14).getValues(); 
    let targetRow = -1; let desc = "";
    for(let i=0; i<data.length; i++){ if(String(data[i][0]).trim() === String(folio).trim() && String(data[i][13]).trim().toUpperCase() === "PENDIENTE"){ targetRow = i + 2; desc = data[i][8]; break; } }
    if(targetRow === -1) { return HtmlService.createHtmlOutput(`<div style="text-align:center;padding:50px;font-family:sans-serif;"><h2>Este folio ya fue pesado o no existe.</h2></div>`); }
    return HtmlService.createHtmlOutput(`
      <html><head><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{font-family:'Arial',sans-serif;background:#F1F5F9;margin:0;padding:20px;}.card{max-width:450px;margin:0 auto;background:white;padding:30px;border-radius:12px;box-shadow:0 4px 15px rgba(0,0,0,0.1);}h2{color:#EF4444;margin-top:0;}.desc{background:#FEE2E2;padding:15px;border-radius:8px;color:#991B1B;font-weight:bold;margin-bottom:20px;}.row{display:flex;gap:10px;margin-bottom:15px;}select,input{width:100%;padding:12px;border:2px solid #CBD5E1;border-radius:6px;font-size:15px;}.btn-add{background:#E2E8F0;color:#475569;border:none;padding:10px;width:100%;border-radius:6px;font-weight:bold;cursor:pointer;margin-bottom:20px;}.btn-add:active{background:#CBD5E1;}.envase-row{display:flex;align-items:center;gap:10px;margin-bottom:10px;background:#F8FAFC;padding:10px;border-radius:6px;}.total-box{font-size:24px;font-weight:bold;color:#0F172A;text-align:center;margin:20px 0;padding:15px;border-top:2px dashed #E2E8F0;}.btn-save{background:#EF4444;color:white;border:none;padding:15px;width:100%;border-radius:6px;font-weight:bold;font-size:16px;cursor:pointer;box-shadow:0 4px 10px rgba(239,68,68,0.3);}</style></head>
      <body><div class="card" id="form-container"><h2>Calculadora de Pesaje</h2><p style="color:#64748B;">Folio: <b>${folio}</b></p><div class="desc">${desc}</div><div class="row"><select id="recipiente"><option value="" disabled selected>Tipo de Recipiente...</option><option value="Bolsa">Bolsa</option><option value="Tambo">Tambo</option><option value="Bidón">Bidón</option></select><select id="unidad"><option value="Kg">Kg</option><option value="Litros">Litros</option></select></div><div id="lista-envases"></div><button class="btn-add" onclick="addEnvase()">+ Agregar peso de otro envase</button><div class="total-box">Total: <span id="total-lbl">0.00</span></div><button class="btn-save" onclick="guardarTodo()">Guardar en Almacén</button></div><div id="msg" style="text-align:center; margin-top:30px; font-weight:bold; font-size:18px;"></div>
      <script>let envasesCount=0;function addEnvase(){envasesCount++;const div=document.createElement('div');div.className='envase-row';div.innerHTML=\`<span style="color:#64748B;font-weight:bold;width:80px;">Envase \${envasesCount}:</span><input type="number" step="0.1" class="peso-input" placeholder="0.0" oninput="calcTotal()">\`;document.getElementById('lista-envases').appendChild(div);}function calcTotal(){let total=0;document.querySelectorAll('.peso-input').forEach(input=>{const val=parseFloat(input.value);if(!isNaN(val))total+=val;});document.getElementById('total-lbl').innerText=total.toFixed(2);return total;}function guardarTodo(){const recipiente=document.getElementById('recipiente').value;const unidad=document.getElementById('unidad').value;const total=calcTotal();if(!recipiente){alert('Selecciona el tipo de recipiente (Bolsa, Tambo, etc.)');return;}if(envasesCount===0||total===0){alert('Debes agregar al menos el peso de un envase.');return;}let notas=[];document.querySelectorAll('.peso-input').forEach(input=>{const val=parseFloat(input.value);if(!isNaN(val))notas.push(val);});const payload={folio:'${folio}',row:${targetRow},recipiente:recipiente,cantidad:envasesCount,pesoTotal:total,unidad:unidad,notas:notas.join(' + ')};document.getElementById('form-container').style.display='none';document.getElementById('msg').innerHTML='Guardando calculo en la base de datos...';document.getElementById('msg').style.color='#F59E0B';google.script.run.withSuccessHandler(function(r){document.getElementById('msg').innerHTML='Guardado exitosamente!<br><span style="font-size:14px;color:#666;font-weight:normal;display:block;margin-top:10px;">La base de datos se ha actualizado limpiamente.<br>Ya puedes cerrar esta ventana.</span>';document.getElementById('msg').style.color='#22C55E';}).guardarPesosResiduos(payload);}addEnvase();</script></body></html>
    `);
  }

  if (action === 'entregar_proveedor') {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_RESIDUOS");
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for(let i=0; i<data.length; i++){
      if(String(data[i][0]).trim() === String(folio).trim()){ sh.getRange(i+2, 14).setValue("RECOLECTADO"); sh.getRange(i+2, 15).setValue(Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm")); }
    }
    return HtmlService.createHtmlOutput(`<div style="font-family:sans-serif;text-align:center;padding:50px;background:#F7F9FC;height:100vh;"><div style="background:white;padding:40px;border-radius:15px;box-shadow:0 10px 30px rgba(0,0,0,0.1);max-width:500px;margin:0 auto;"><h1 style="color:#22C55E;margin-top:0;">ENTREGADO AL PROVEEDOR</h1><p style="font-size:18px;">El folio <b>${folio}</b> ha sido entregado exitosamente.</p></div></div>`);
  }
  return HtmlService.createHtmlOutput("<h1>Acción no reconocida</h1>");
}

/* ═══════════════════════════════════════════════════════════════════════════
   MÓDULO LIMPIEZA DE EQUIPOS — DENTRO DE INFRAESTRUCTURA & SOPORTE
═══════════════════════════════════════════════════════════════════════════ */

function _getCatalogoInfraCompleto() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('catalogo_infra');
  if (cachedData) return JSON.parse(cachedData);
  const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DATA_INFRA");
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  const data = rows.filter(r => String(r[0]).trim() !== "").map(r => ({ noEquipo : String(r[0]).trim(), ne : String(r[1]).trim(), area : String(r[2]).trim(), correo : String(r[3]).trim(), usuario  : String(r[4]).trim(), so : String(r[5]).trim() }));
  cache.put('catalogo_infra', JSON.stringify(data), 600);
  return data;
}

function obtenerCatalogoInfra() {
  try {
    const catalogoCompleto = _getCatalogoInfraCompleto();
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    const esSA = _getSuperAdmins().includes(email);
    const esInfra = _getEquipoNotif("INFRA").includes(email);

    if (esSA || esInfra) return { ok: true, data: catalogoCompleto };

    let userArea = "";
    let limpAdminRaw = "";
    try {
      const shU = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
      if (shU && shU.getLastRow() >= 2) {
        const uRows = shU.getRange(2, 1, shU.getLastRow() - 1, 25).getValues();
        for (let i = 0; i < uRows.length; i++) {
          if (String(uRows[i][1]).toLowerCase().trim() === email) {
            userArea     = _normArea(uRows[i][3]);
            limpAdminRaw = String(uRows[i][24] || "").toUpperCase().trim();
            break;
          }
        }
      }
    } catch(ue) {}

    const esGlobal = (limpAdminRaw === "SI" || limpAdminRaw === "TRUE" || limpAdminRaw === "SÍ" || limpAdminRaw === "YES");
    if (esGlobal) return { ok: true, data: catalogoCompleto };

    const esAdminArea = limpAdminRaw !== "" && limpAdminRaw !== "NO" && limpAdminRaw !== "FALSE";
    const areaAdmin = _normArea(limpAdminRaw);
    if (esAdminArea) {
      return { ok: true, data: catalogoCompleto.filter(function(eq) {
        return _normArea(eq.area) === areaAdmin;
      }) };
    }

    return { ok: true, data: catalogoCompleto.filter(function(eq) {
      return userArea && _normArea(eq.area) === userArea;
    }) };
  } catch(e) { return { ok: false, error: e.message, data: [] }; }
}

function obtenerDatosLimpieza() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    let esAdmin = _getSuperAdmins().includes(email) || _getEquipoNotif("INFRA").includes(email);
    let userArea = "";
    let limpAdminRaw = "";
    
    try {
      const shU = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
      if (shU && shU.getLastRow() >= 2) {
        const uRows = shU.getRange(2, 1, shU.getLastRow()-1, 25).getValues();
        for (let i = 0; i < uRows.length; i++) {
          if (String(uRows[i][1]).toLowerCase().trim() === email) {
            userArea     = _normArea(uRows[i][3]);
            limpAdminRaw = String(uRows[i][24] || "").toUpperCase().trim();
            break;
          }
        }
      }
    } catch(ue) {}

    const esGlobal    = (limpAdminRaw === "SI" || limpAdminRaw === "TRUE" || limpAdminRaw === "SÍ" || limpAdminRaw === "YES");
    const esAdminArea = !esGlobal && limpAdminRaw !== "" && limpAdminRaw !== "NO" && limpAdminRaw !== "FALSE";
    if (esGlobal) esAdmin = true;

    // ÁREA EFECTIVA normalizada sin acentos
    const areaFiltro = esAdmin ? "" : _normArea(esAdminArea ? limpAdminRaw : userArea);

    const catalogoCompleto = _getCatalogoInfraCompleto();
    
    const catalogo = areaFiltro
      ? catalogoCompleto.filter(function(eq) { return _normArea(eq.area) === areaFiltro; })
      : catalogoCompleto;
    
    const shLimp = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_LIMPIEZA");
    let solicitudes = [];
    if (shLimp && shLimp.getLastRow() >= 2) {
      const fd  = v => { try { return (v instanceof Date) ? Utilities.formatDate(v,"GMT-6","dd/MM/yyyy") : String(v||""); } catch(e) { return String(v||""); } };
      const fdt = v => { try { return (v instanceof Date) ? Utilities.formatDate(v,"GMT-6","dd/MM/yyyy HH:mm") : String(v||""); } catch(e) { return String(v||""); } };
      const ft  = v => { try { return (v instanceof Date) ? Utilities.formatDate(v,"GMT-6","HH:mm") : String(v||""); } catch(e) { return String(v||""); } };
      const ffh = v => { try { return (v instanceof Date) ? Utilities.formatDate(v,"GMT-6","yyyy-MM-dd HH:mm") : String(v||""); } catch(e) { return String(v||""); } };
      const rows = shLimp.getRange(2,1,shLimp.getLastRow()-1,18).getValues();
      for (let i = rows.length-1; i >= 0; i--) {
        const r = rows[i];
        if (areaFiltro) {
          if (_normArea(r[5]) !== areaFiltro) continue;
        }
        solicitudes.push({ folio: String(r[0]||""), fecha: fd(r[1]), hora: ft(r[2]), correo: String(r[3]||""), nombre: String(r[4]||""), area: String(r[5]||""), noEquipo: String(r[6]||""), neEquipo: String(r[7]||""), usuarioEq: String(r[8]||""), articulos: String(r[9]||""), notas: String(r[10]||""), estatus: String(r[11]||"PENDIENTE"), fechaRes: fdt(r[12]), resueltoPor: String(r[13]||""), fechaHorario: ffh(r[17]) });
      }
    }

    return { ok: true, esAdmin: esAdmin || esAdminArea, adminArea: esAdminArea ? limpAdminRaw : "", userArea: userArea, catalogo: catalogo, solicitudes: solicitudes };
  } catch(e) { return { ok: false, error: e.message, esAdmin: false, adminArea: "", userArea: "", catalogo: [], solicitudes: [] }; }
}

// CRITICO: Google Sheets auto-convierte "2026-03-10 09:00" a Date object al guardarlo.
// String(Date) produce "Mon Mar 10 2026 09:00:00 GMT-0600" — jamas va a coincidir.
// _normFH detecta Date object y lo formatea antes de normalizar.
function _normFH(fh) {
  if (!fh) return "";
  if (fh instanceof Date) {
    return Utilities.formatDate(fh, "GMT-6", "yyyy-MM-dd HH:mm");
  }
  var s = String(fh).trim();
  if (!s || s === "0" || s === "false") return "";
  var p = s.split(" ");
  if (p.length >= 2) {
    var h = p[1].trim();
    if (/^\d:\d{2}$/.test(h)) h = "0" + h;
    return p[0] + " " + h;
  }
  return s;
}

function obtenerHorasOcupadasLimpieza(fecha) {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName('DB_LIMPIEZA');
    if (!sh || sh.getLastRow() < 2) return { pendientes: [], resueltas: [] };
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues();
    const pendientes = [];
    const resueltas  = [];
    for (let i = 0; i < data.length; i++) {
      const estatus = String(data[i][11]).toUpperCase().trim();
      if (estatus === 'CANCELADA') continue; // Solo canceladas liberan el slot
      const fh = _normFH(data[i][17]);
      if (fh.startsWith(fecha) && fh.length > 10) {
        const hora = fh.substring(11, 16);
        if (!hora) continue;
        if (estatus === 'RESUELTA') {
          if (!resueltas.includes(hora)) resueltas.push(hora);
        } else {
          if (!pendientes.includes(hora)) pendientes.push(hora);
        }
      }
    }
    return { pendientes: pendientes, resueltas: resueltas };
  } catch(e) { return { pendientes: [], resueltas: [] }; }
}

function registrarLimpieza(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: "Sistema ocupado, intenta de nuevo." };
  try {
    // ── VALIDACIÓN DE ÁREA: bloquear solicitudes para equipos de otra área ──────
    const userEmailReg = Session.getActiveUser().getEmail().toLowerCase().trim();
    const esAdminReg   = _getSuperAdmins().includes(userEmailReg) ||
                         _getEquipoNotif("INFRA").includes(userEmailReg);
    if (!esAdminReg && payload.noEquipo && payload.area) {
      try {
        // Verificar si es admin de área
        let limpAdminVal = "";
        const shUVal = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
        if (shUVal && shUVal.getLastRow() >= 2) {
          const uRowsVal = shUVal.getRange(2, 1, shUVal.getLastRow()-1, 25).getValues();
          for (let iv = 0; iv < uRowsVal.length; iv++) {
            if (String(uRowsVal[iv][1]).toLowerCase().trim() === userEmailReg) { limpAdminVal = String(uRowsVal[iv][24] || "").toUpperCase().trim(); break; }
          }
        }
        const esGlobalVal = (limpAdminVal === "SI" || limpAdminVal === "TRUE" || limpAdminVal === "SÍ" || limpAdminVal === "YES");
        if (!esGlobalVal) {
          const catItems = _getCatalogoInfraCompleto();
          const eqFound  = catItems.find(function(eq){ return eq.noEquipo === payload.noEquipo; });
          if (eqFound) {
            const areaEq = _normArea(eqFound.area);
            // Admin de área puede solicitar para su área administrada
            const areaPermitida = (limpAdminVal && limpAdminVal !== "NO" && limpAdminVal !== "FALSE") ? _normArea(limpAdminVal) : _normArea(payload.area);
            if (areaEq && areaPermitida && areaEq !== areaPermitida) {
              return { ok: false, error: "⛔ No tienes permiso para solicitar limpieza de equipos de " + areaEq + ". Tu área es: " + areaPermitida };
            }
          }
        }
      } catch(ve) { Logger.log("Validación área limpieza: " + ve.message); }
    }
    // ──────────────────────────────────────────────────────────────────────────────
    const ss  = SpreadsheetApp.openById(SS_MASTER_ID);
    let sh = ss.getSheetByName("DB_LIMPIEZA");
    if (!sh) {
      sh = ss.insertSheet("DB_LIMPIEZA");
      const hdr = ["FOLIO","FECHA","HORA","SOLICITANTE_CORREO","SOLICITANTE_NOMBRE","AREA_SOLICITANTE","NO_EQUIPO","NE_EQUIPO","USUARIO_EQUIPO","ARTICULOS","NOTAS","ESTATUS","FECHA_RESOLUCION","RESUELTO_POR","REEMPLAZADO","DETALLE_REEMPLAZO","OBSERVACIONES","FECHA_HORARIO"];
      sh.getRange(1,1,1,hdr.length).setValues([hdr]).setBackground("#0D9488").setFontColor("#fff").setFontWeight("bold"); sh.setFrozenRows(1);
    }
    // ── VALIDACIÓN ANTI-DUPLICADO: verificar que el slot no esté ya ocupado ──────
    // Se hace DENTRO del lock para evitar condición de carrera
    if (payload.fechaHorario && sh.getLastRow() >= 2) {
      const existRows = sh.getRange(2, 1, sh.getLastRow() - 1, 18).getValues();
      for (let i = 0; i < existRows.length; i++) {
        const estatus = String(existRows[i][11]).toUpperCase().trim();
        if (estatus === "CANCELADA") continue; // Solo canceladas liberan el slot
        const fhExist = _normFH(existRows[i][17]);        // SIN String() — maneja Date directamente
        const fhNew   = _normFH(payload.fechaHorario);
        if (fhExist === fhNew && fhExist !== "") {
          const horaOcup = fhExist.length > 10 ? fhExist.substring(11, 16) : fhExist;
          return { ok: false, error: "SLOT_OCUPADO", horario: horaOcup, fechaHorario: fhExist };
        }
      }
    }
    // ─────────────────────────────────────────────────────────────────────────────

    const now = new Date(); const fecha = Utilities.formatDate(now,"GMT-6","dd/MM/yyyy"); const hora  = Utilities.formatDate(now,"GMT-6","HH:mm");
    const seq = String(sh.getLastRow()).padStart(4,"0");
    const folio = "LIM-" + Utilities.formatDate(now,"GMT-6","yyyyMMdd") + "-" + seq;

    sh.appendRow([ folio, fecha, hora, payload.correo||"", payload.nombre||"", payload.area||"", payload.noEquipo||"", payload.neEquipo||"", payload.usuarioEquipo||"", (payload.articulos||[]).join(", "), payload.notas||"", "PENDIENTE", "", "", "", "", "", payload.fechaHorario||"" ]);

    try {
      const artLi = (payload.articulos||[]).map(a=>`<li style="padding:3px 0;color:#0F172A">${a}</li>`).join("");
      const html = `<div style="font-family:Arial,sans-serif;background:#F7F9FC;padding:28px"><div style="max-width:560px;margin:0 auto;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.10)"><div style="background:linear-gradient(135deg,#0D9488,#0F766E);padding:22px 28px"><div style="font-size:9px;letter-spacing:3px;color:rgba(255,255,255,.7);margin-bottom:5px">INFRAESTRUCTURA · SANTILLÁN RAMÍREZ</div><div style="font-size:20px;font-weight:900;color:#fff">SOLICITUD DE LIMPIEZA</div></div><div style="padding:22px 28px"><table width="100%" style="margin-bottom:14px"><tr><td width="50%" style="padding-right:6px"><div style="background:#F0FDF4;border-left:4px solid #10B981;border-radius:8px;padding:11px"><div style="font-size:9px;color:#64748B;letter-spacing:2px;margin-bottom:3px">FOLIO</div><div style="font-size:16px;font-weight:900;color:#064E3B;font-family:monospace">${folio}</div></div></td><td width="50%" style="padding-left:6px"><div style="background:#F0FDF4;border-left:4px solid #10B981;border-radius:8px;padding:11px"><div style="font-size:9px;color:#64748B;letter-spacing:2px;margin-bottom:3px">EQUIPO</div><div style="font-size:16px;font-weight:900;color:#064E3B">${payload.noEquipo}</div></div></td></tr></table><div style="margin-bottom:12px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:5px">SOLICITANTE</div><div style="font-weight:700;color:#0F172A">${payload.nombre} · ${payload.area}</div></div><div style="margin-bottom:12px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:5px">ARTÍCULOS A LIMPIAR</div><ul style="margin:0;padding-left:18px;line-height:2">${artLi}</ul></div>${payload.notas?`<div style="background:#F8FAFC;border-radius:8px;padding:11px;font-size:13px;color:#475569"><b>Notas:</b> ${payload.notas}</div>`:""}</div></div></div>`;
      const dests = _getEquipoNotif("INFRA");
      MailApp.sendEmail({ to: dests.join(","), name:"Sistema Corporativo · Infraestructura", replyTo: payload.correo||"", subject:`[LIMPIEZA] ${folio} · ${payload.noEquipo}`, htmlBody: html });
      // Confirmación al solicitante
      if (payload.correo && payload.correo.includes("@")) {
        const fhMostrar = payload.fechaHorario || "Por confirmar";
        const htmlConfirm = _emailShell({ headerGrad: "linear-gradient(135deg,#0D9488 0%,#0F766E 100%)", accent: "#0D9488", title: "SOLICITUD REGISTRADA", subtitle: "Limpieza de equipo · Infraestructura", folioValue: folio, body: _emailInfoBlock("EQUIPO", payload.noEquipo, "#0D9488") + _emailInfoBlock("HORARIO SOLICITADO", fhMostrar, "#0F766E") + _emailInfoBlock("SOLICITANTE", payload.nombre, "#14B8A6") + '<div style="background:#F0FDFA;border:1px solid #99F6E4;border-radius:10px;padding:16px;margin-top:14px;text-align:center"><div style="font-size:12px;color:#134E4A;line-height:1.6">Tu solicitud ha sido recibida. El equipo de Infraestructura se pondrá en contacto contigo para coordinar el servicio.</div></div>', footerNote: "Este es un correo automático. No responder directamente." });
        MailApp.sendEmail({ to: payload.correo, name: "Sistema Corporativo · Infraestructura", subject: `[LIMPIEZA CONFIRMADA] Folio ${folio} · ${payload.noEquipo}`, htmlBody: htmlConfirm });
      }
    } catch(me) { Logger.log("Email limpieza: " + me.message); }

    return { ok: true, folio: folio };
  } catch(e) { return { ok: false, error: e.message }; } finally { lock.releaseLock(); }
}

function resolverLimpieza(folio, reemplazado, detalleReemplazo, observaciones) {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    let esAdmin = _getSuperAdmins().includes(email) || _getEquipoNotif("INFRA").includes(email);
    if (!esAdmin) {
      try {
        const shU = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
        if (shU && shU.getLastRow() >= 2) {
          const uRows = shU.getRange(2, 2, shU.getLastRow()-1, 24).getValues();
          for (let i = 0; i < uRows.length; i++) { if (String(uRows[i][0]).toLowerCase().trim() === email) { var _raw = String(uRows[i][23] || "").toUpperCase().trim(); esAdmin = _raw !== "" && _raw !== "NO" && _raw !== "FALSE"; break; } }
        }
      } catch(ue) {}
    }
    if (!esAdmin) return { ok: false, error: "Sin permisos de infraestructura." };
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_LIMPIEZA");
    if (!sh) return { ok: false, error: "DB_LIMPIEZA no existe aún." };
    const vals = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === String(folio).trim()) {
        sh.getRange(i+2, 12).setValue("RESUELTA");
        sh.getRange(i+2, 13).setValue(Utilities.formatDate(new Date(),"GMT-6","dd/MM/yyyy HH:mm"));
        sh.getRange(i+2, 14).setValue(email);
        if (reemplazado !== undefined) sh.getRange(i+2, 15).setValue(reemplazado || "NO");
        if (detalleReemplazo !== undefined) sh.getRange(i+2, 16).setValue(detalleReemplazo || "");
        if (observaciones !== undefined && observaciones !== null) sh.getRange(i+2, 17).setValue(observaciones || "");
        registrarAuditoria("LIMPIEZA","LIMPIEZA_RESUELTA","Folio: "+folio+", Reemplazo: "+(reemplazado||"NO"),folio);
        return { ok: true };
      }
    }
    return { ok: false, error: "Folio no encontrado: " + folio };
  } catch(e) { return { ok: false, error: e.message }; }
}

function cancelarLimpieza(folio, motivo) {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase().trim();
    // NOTA: ver L1183 — reemplazado abajo){return x.toLowerCase()}).indexOf(email) > -1;
    if (!esAdmin) {
      try {
        var shU = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
        if (shU && shU.getLastRow() >= 2) {
          var uRows = shU.getRange(2, 2, shU.getLastRow()-1, 24).getValues();
          for (var i = 0; i < uRows.length; i++) { if (String(uRows[i][0]).toLowerCase().trim() === email) { var _raw = String(uRows[i][23] || "").toUpperCase().trim(); esAdmin = _raw !== "" && _raw !== "NO" && _raw !== "FALSE"; break; } }
        }
      } catch(ue) {}
    }
    if (!esAdmin) return { ok: false, error: "Sin permisos de infraestructura." };
    var sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_LIMPIEZA");
    if (!sh) return { ok: false, error: "DB_LIMPIEZA no existe aún." };
    var vals = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === String(folio).trim()) {
        sh.getRange(i+2, 12).setValue("CANCELADA");
        sh.getRange(i+2, 13).setValue(Utilities.formatDate(new Date(),"GMT-6","dd/MM/yyyy HH:mm"));
        sh.getRange(i+2, 14).setValue(email);
        sh.getRange(i+2, 17).setValue(motivo || "");
        registrarAuditoria("LIMPIEZA","LIMPIEZA_CANCELADA","Folio: "+folio+", Motivo: "+(motivo||"Sin motivo"),folio);
        return { ok: true };
      }
    }
    return { ok: false, error: "Folio no encontrado: " + folio };
  } catch(e) { return { ok: false, error: e.message }; }
}
function posponerLimpieza(folio, nuevaFechaHorario, motivo) {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase().trim();
    var esAdmin = _getSuperAdmins().indexOf(email) > -1
               || _getEquipoNotif("INFRA").indexOf(email) > -1;
    if (!esAdmin) {
      try {
        var shU = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
        if (shU && shU.getLastRow() >= 2) {
          var uRows = shU.getRange(2, 2, shU.getLastRow()-1, 24).getValues();
          for (var i = 0; i < uRows.length; i++) {
            if (String(uRows[i][0]).toLowerCase().trim() === email) {
              var _raw = String(uRows[i][23] || "").toUpperCase().trim();
              esAdmin = _raw !== "" && _raw !== "NO" && _raw !== "FALSE";
              break;
            }
          }
        }
      } catch(ue) {}
    }
    if (!esAdmin) return { ok: false, error: "Sin permisos de infraestructura." };

    var sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName("DB_LIMPIEZA");
    if (!sh) return { ok: false, error: "DB_LIMPIEZA no existe." };

    // Verificar que el nuevo slot no esté ocupado (solo CANCELADA libera)
    var nuevaFHNorm = _normFH(nuevaFechaHorario);
    if (nuevaFHNorm && sh.getLastRow() >= 2) {
      var existRows = sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues();
      for (var i = 0; i < existRows.length; i++) {
        if (String(existRows[i][0]).trim() === String(folio).trim()) continue;
        var estSl = String(existRows[i][11]).toUpperCase().trim();
        if (estSl === "CANCELADA") continue;
        var fhExist = _normFH(existRows[i][17]);
        if (fhExist === nuevaFHNorm && fhExist !== "") {
          var horaOcup = fhExist.length > 10 ? fhExist.substring(11, 16) : fhExist;
          return { ok: false, error: "SLOT_OCUPADO", horario: horaOcup };
        }
      }
    }

    // Encontrar el folio y reprogramar
    var vals = sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === String(folio).trim()) {
        var solicitanteCorreo = String(vals[i][3] || "");
        var solicitanteNombre = String(vals[i][4] || "");
        var noEquipo          = String(vals[i][6] || "");
        var fhAnterior        = _normFH(vals[i][17]);

        sh.getRange(i+2, 18).setValue(nuevaFechaHorario);
        var obsActual = String(vals[i][16] || "");
        var obsNueva  = "POSPUESTO (" + Utilities.formatDate(new Date(),"GMT-6","dd/MM/yyyy HH:mm") + "): " + (motivo || "") + " — Anterior: " + fhAnterior + (obsActual ? (" | " + obsActual) : "");
        sh.getRange(i+2, 17).setValue(obsNueva);

        // Notificar al solicitante por correo
        try {
          if (solicitanteCorreo && solicitanteCorreo.indexOf("@") > -1) {
            var htmlPosp = _emailShell({
              headerGrad: "linear-gradient(135deg,#D97706 0%,#B45309 100%)",
              accent: "#D97706",
              title: "SERVICIO REPROGRAMADO",
              subtitle: "Limpieza de equipo · Infraestructura",
              folioValue: folio,
              body: _emailInfoBlock("EQUIPO", noEquipo, "#D97706") +
                    _emailInfoBlock("NUEVO HORARIO", nuevaFechaHorario, "#B45309") +
                    _emailInfoBlock("HORARIO ANTERIOR", fhAnterior, "#92400E") +
                    _emailInfoBlock("MOTIVO", motivo || "Reprogramación por Infraestructura", "#78350F") +
                    '<div style="background:#FFFBEB;border:1px solid #FDE68A;border-radius:10px;padding:16px;margin-top:14px;text-align:center">' +
                    '<div style="font-size:12px;color:#92400E;line-height:1.6">Tu servicio de limpieza fue reprogramado por el equipo de Infraestructura. Si tienes alguna duda, comunícate con ellos directamente.</div></div>',
              footerNote: "Este es un correo automático. No responder directamente."
            });
            MailApp.sendEmail({
              to: solicitanteCorreo,
              name: "Sistema Corporativo · Infraestructura",
              subject: "[LIMPIEZA REPROGRAMADA] Folio " + folio + " · " + noEquipo,
              htmlBody: htmlPosp
            });
          }
        } catch(me) { Logger.log("Email posponer limpieza: " + me.message); }

        registrarAuditoria("LIMPIEZA","LIMPIEZA_POSPUESTA",
          "Folio: "+folio+", Nuevo: "+nuevaFechaHorario+", Anterior: "+fhAnterior, folio);
        return { ok: true, nuevaFechaHorario: nuevaFechaHorario };
      }
    }
    return { ok: false, error: "Folio no encontrado: " + folio };
  } catch(e) { return { ok: false, error: e.message }; }
}
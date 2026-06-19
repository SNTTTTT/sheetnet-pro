/**
 * ╔═══════════════════════════════════════════════════════════════════╗
 * ║  DISEÑO Y SOLUCIÓN INTEGRAL SANTILLÁN RAMÍREZ                      ║
 * ║  Sistema Corporativo · V5.0 (Núcleo Blindado)                     ║
 * ║                                                                   ║
 * ║  MÓDULOS ACTIVOS:                                                 ║
 * ║    · Actas Administrativas                                        ║
 * ║    · Expediente Digital                                           ║
 * ║    · Residuos Peligrosos                                          ║
 * ║                                                                   ║
 * ║  Módulos retirados (respaldados en /_archivo_memoria):            ║
 * ║    Compras · Soporte TI · Mantenimiento · Limpieza · Transporte  ║
 * ╚═══════════════════════════════════════════════════════════════════╝
 */

const SS_MASTER_ID    = "1rdfAXSEml3cDamy48r02FMrrZ19G0C1-LyRaK2oRQ4w";
const SH_USUARIOS     = "DATA_USUARIOS";
const SH_EMPLEADOS    = "DATA_EMPLEADOS";
const SH_ACTAS        = "DB_ACTAS";
const SH_EXPEDIENTES  = "DB_EXPEDIENTES";
const SH_RESIDUOS     = "DB_RESIDUOS";
const SH_FOTOS        = "DATA_FOTOS";
const SH_AUDITORIA    = "DB_AUDITORIA";
const DRIVE_EXP_ID    = "1r5_jIBIFHVCnGGkmy36BS9F34IbH8BJn";

/* ═══════════════════════════════════════════════════════════════════════════
   📋 DATA_USUARIOS — Columna 26 (índice 25): equipoNotificacion
   ───────────────────────────────────────────────────────────────────────────
   Valores válidos (separados por coma): SUPER_ADMIN | CH | RESIDUOS | MANT
     SUPER_ADMIN → Acceso total al sistema
     CH          → Capital Humano, recibe actas
     RESIDUOS    → Recibe alertas de residuos peligrosos (equipo de recolección)
     MANT        → Equipo de recolección/almacén (respaldo de RESIDUOS)
   Para alta/baja de un equipo: editar la hoja, sin tocar código.
   NOTA: Las columnas de permisos de módulos retirados se conservan en la hoja
   por integridad de datos, pero el sistema sólo expone Actas/Expediente/Residuos.
═══════════════════════════════════════════════════════════════════════════ */

/**
 * _cargarUsuariosBD()
 * Lee DATA_USUARIOS completo (cols 1-26) con caché de script 10 min.
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

/** _getSuperAdmins() — emails con SUPER_ADMIN en col 26 (equipoNotif). */
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
    if (activo && correo.includes('@') && equipo.includes('SUPER_ADMIN')) admins.push(correo);
  });
  try { cache.put('_sa_v1', JSON.stringify(admins), 600); } catch(e) {}
  return admins;
}

/** _getEquipoNotif(modulo) — emails activos de un equipo de notificación. */
function _getEquipoNotif(modulo) {
  const rows = _cargarUsuariosBD();
  const equipo = [];
  const moduloUp = String(modulo || '').toUpperCase().trim();
  rows.forEach(r => {
    const notif = String(r[25] || '').toUpperCase();
    const correo = String(r[1] || '').toLowerCase().trim();
    const activo = normalizarBooleano(r[6]);
    if (activo && correo.includes('@') && notif.includes(moduloUp)) equipo.push(correo);
  });
  return [...new Set(equipo)];
}

/** _getReplyTo(modulo) — primer correo del equipo para replyTo. */
function _getReplyTo(modulo) {
  const equipo = _getEquipoNotif(modulo);
  return equipo.length > 0 ? equipo[0] : '';
}

/** Equipo de recolección de residuos: RESIDUOS, con respaldo a MANT. */
function _getEquipoResiduos() {
  const r = _getEquipoNotif("RESIDUOS");
  return r.length ? r : _getEquipoNotif("MANT");
}

function getWebAppUrl() {
  try { return ScriptApp.getService().getUrl(); } catch(e) { return ""; }
}

/* ═══════════════════════════════════════════════════════════════════
   📧 SISTEMA DE EMAILS PREMIUM — Plantilla unificada corporativa
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
    '<div style="background:#FFFFFF;padding:28px 32px;border-left:1px solid #E8ECF0;border-right:1px solid #E8ECF0">' + body + '</div>' +
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



/* ═══════════════════════════════════════════════════════════════
   BASE DE DATOS — inicialización de hojas (sólo módulos activos)
═══════════════════════════════════════════════════════════════ */
function _inicializarBase() {
  const cache = CacheService.getScriptCache();
  if (cache.get('_base_init_v5')) return;
  const ss = SpreadsheetApp.openById(SS_MASTER_ID);
  // DATA_USUARIOS conserva el esquema completo de 25 columnas por integridad de datos.
  const esquemas = [
    { n: SH_USUARIOS, h: ["NE", "CORREO", "NOMBRE", "AREA", "PUESTO", "TIPO", "ACTIVO", "COMPRAS_ACCESO", "COMPRAS_ADMIN", "COMPRAS_AREA_RESPONSABLE", "ACTAS_ACCESO", "ACTAS_ADMIN", "ACTAS_PUEDE_SOLICITAR", "SOPORTE_ACCESO", "SOPORTE_ADMIN", "MANTENIMIENTO_ACCESO", "MANTENIMIENTO_ADMIN", "RESIDUOS_ACCESO", "RESIDUOS_ADMIN", "TRANSPORTE_ACCESO", "TRANSPORTE_ADMIN", "EXPEDIENTES_ACCESO", "EXPEDIENTES_ADMIN", "LIMPIEZA_ACCESO", "LIMPIEZA_ADMIN"] },
    { n: SH_EMPLEADOS, h: ["ID_EMPLEADO", "NOMBRE", "CORREO", "PUESTO", "AREA", "TIPO", "REGIMEN_FISCAL"] },
    { n: SH_ACTAS, h: ["FOLIO", "FECHA", "HORA", "SOLICITANTE_CORREO", "SOLICITANTE_NOMBRE", "ID_REPORTADO", "NOMBRE_REPORTADO", "PUESTO_REPORTADO", "AREA_REPORTADO", "TIPO_INCIDENCIA", "DESCRIPCION", "ESTATUS", "NOTIFICADO", "AREA_SOLICITANTE", "TESTIGO_1", "TESTIGO_2"] },
    { n: SH_RESIDUOS, h: ["FOLIO", "PARTIDA", "FECHA", "HORA", "SOLICITANTE_CORREO", "SOLICITANTE_NOMBRE", "AREA", "TIPO_RESIDUO", "DESCRIPCION", "RECIPIENTE", "CANTIDAD", "PESO_TOTAL", "UNIDAD", "ESTATUS", "FECHA_RECOLECCION"] },
    { n: SH_FOTOS, h: ["NE", "URL_FOTO"] },
    { n: SH_AUDITORIA, h: ["ID", "FECHA", "HORA", "EMAIL", "NOMBRE", "AREA", "MODULO", "ACCION", "DETALLE", "REFERENCIA", "IP_SESION"] }
  ];
  esquemas.forEach(e => { let sheet = ss.getSheetByName(e.n); if (!sheet) { sheet = ss.insertSheet(e.n); sheet.getRange(1, 1, 1, e.h.length).setValues([e.h]).setBackground("#5DC1E8").setFontColor("#FFFFFF").setFontWeight("bold"); sheet.setFrozenRows(1); } });
  cache.put('_base_init_v5', '1', 1800);
}

function normalizarBooleano(valor) {
  if (valor === true || valor === 1) return true;
  const s = String(valor).toUpperCase().trim();
  return s === "SI" || s === "TRUE" || s === "YES" || s === "SÍ";
}

// Normaliza un área: MAYÚSCULAS + sin acentos
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

/* ═══════════════════════════════════════════════════════════════
   PERFIL — caché de usuario + seguridad estricta de accesos
   Sólo expone: actas · expedientes · residuos
═══════════════════════════════════════════════════════════════ */
function obtenerPerfilCompleto(forceRefresh = false) {
  const email = Session.getActiveUser().getEmail().toLowerCase().trim();
  const cache = CacheService.getUserCache();
  const cacheKey = 'perfil_v5_' + email;

  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);
  }

  _inicializarBase();
  const ss = SpreadsheetApp.openById(SS_MASTER_ID);
  const esSA = _getSuperAdmins().includes(email);
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
        actas_acceso: normalizarBooleano(data[i][10]), actas_admin: normalizarBooleano(data[i][11]), actas_puede_solicitar: normalizarBooleano(data[i][12]),
        residuos_acceso: normalizarBooleano(data[i][17]), residuos_admin: normalizarBooleano(data[i][18]),
        expedientes_acceso: normalizarBooleano(data[i][21]), expedientes_admin: normalizarBooleano(data[i][22])
      };
      break;
    }
  }

  if (!u) { return esSA ? crearPerfilSuperAdmin(email) : { acceso: false, email, razon: "USUARIO_NO_REGISTRADO" }; }

  const perfilFinal = {
    acceso: true, fotoUrl: _getFotoUrl(u.numeroOperador), email: u.correo, numeroOperador: u.numeroOperador, nombre: u.nombre,
    area: u.area, puesto: u.puesto, tipo: u.tipo, esSuperAdmin: esSA, esAdminUsuarios: esSA,
    actas: { acceso: u.actas_acceso || esSA, admin: u.actas_admin || esSA, puedeSolicitar: u.actas_puede_solicitar || esSA },
    residuos: { acceso: u.residuos_acceso || esSA, admin: u.residuos_admin || esSA },
    expedientes: { acceso: u.expedientes_acceso || esSA, admin: u.expedientes_admin || esSA }
  };

  cache.put(cacheKey, JSON.stringify(perfilFinal), 300);
  return perfilFinal;
}

function crearPerfilSuperAdmin(email) {
  return {
    acceso: true, fotoUrl: '', email: email, numeroOperador: "SA-999", nombre: email.split("@")[0].toUpperCase(),
    area: "DIRECCIÓN", puesto: "SUPER ADMIN", tipo: "ADMINISTRATIVO", esSuperAdmin: true, esAdminUsuarios: true,
    actas: { acceso: true, admin: true, puedeSolicitar: true },
    residuos: { acceso: true, admin: true },
    expedientes: { acceso: true, admin: true }
  };
}

/** cargarAreas() — lista de áreas únicas (alimenta el filtro de Gestión de Actas). */
function cargarAreas() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('areas_v5');
    if (cached) return JSON.parse(cached);
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
    if(!sh || sh.getLastRow()<2) return [];
    const rows = sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
    const areas = new Set();
    rows.forEach(r => { const a = String(r[3]).trim().toUpperCase(); if(a && a !== "UNDEFINED") areas.add(a); });
    const result = [...areas].sort();
    try { cache.put('areas_v5', JSON.stringify(result), 600); } catch(ce) {}
    return result;
  } catch(e) { return []; }
}

function buscarEmpleadoPorId(id) {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_EMPLEADOS);
    const rows = sh.getDataRange().getValues();
    for(let i=1; i<rows.length; i++) {
      if(String(rows[i][0]).trim() === String(id).trim()) {
        return { encontrado: true, nombre: String(rows[i][1]), puesto: String(rows[i][3]), area: String(rows[i][4]).toUpperCase(), tipo: String(rows[i][5]||"OPERADOR"), regimenFiscal: String(rows[i][6]||"NO FISCAL").toUpperCase().trim() };
      }
    }
    return { encontrado: false };
  } catch(e) { return { encontrado: false }; }
}

/* ═══════════════════════════════════════════════════════════════
   MÓDULO · ACTAS ADMINISTRATIVAS
═══════════════════════════════════════════════════════════════ */
function registrarActa(payload) {
  const ss = SpreadsheetApp.openById(SS_MASTER_ID);
  if (!payload.puedeSolicitar) throw new Error("⛔ No tienes permisos.");
  const perfilSolicitante = obtenerPerfilCompleto();
  const puestoJefe = perfilSolicitante.puesto.toUpperCase();
  const areaJefe = perfilSolicitante.area.toUpperCase();
  const esRH = areaJefe.includes("HUMANO") || areaJefe.includes("RH") || areaJefe.includes("CAPITAL");
  const esSuperAdmin = perfilSolicitante.esSuperAdmin;
  const rangosAutoridad = ["JEFE", "JEFA", "LIDER", "LÍDER", "GERENTE", "DIRECTOR", "COORDINADOR", "SUPERVISOR", "ENCARGADO"];
  if (!esRH && !esSuperAdmin && !rangosAutoridad.some(r => puestoJefe.includes(r))) throw new Error("⛔ JERARQUÍA INSUFICIENTE.");
  const de = ss.getSheetByName(SH_EMPLEADOS);
  const rows = de ? de.getDataRange().getValues() : [];
  let target = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(payload.idReportado).trim()) {
      target = { nombre: String(rows[i][1]), correo: String(rows[i][2]), puesto: String(rows[i][3]), area: String(rows[i][4]).toUpperCase(), tipo: String(rows[i][5]||"OPERADOR"), regimenFiscal: String(rows[i][6]||"NO FISCAL").toUpperCase().trim() };
      break;
    }
  }
  if (!target) throw new Error("❌ ID no encontrado.");
  if (!esRH && !esSuperAdmin && areaJefe !== target.area) throw new Error("⛔ FUERA DE JURISDICCIÓN.");
  const dbActas = ss.getSheetByName(SH_ACTAS);
  const now = new Date();
  const folio = "ACT-" + Utilities.formatDate(now,"GMT-6","yyyyMMdd") + "-" + String(dbActas.getLastRow()).padStart(4,"0");
  const fecha = Utilities.formatDate(now,"GMT-6","dd/MM/yyyy");
  const hora = Utilities.formatDate(now,"GMT-6","HH:mm");
  const tipoActa = payload.tipoActa || target.regimenFiscal || "NO FISCAL";
  dbActas.appendRow([ folio, fecha, hora, perfilSolicitante.email, perfilSolicitante.nombre, payload.idReportado, target.nombre, target.puesto, target.area, tipoActa, payload.descripcion, "SOLICITADA", "SI", areaJefe, payload.testigo1||"", payload.testigo2||"" ]);
  registrarAuditoria("ACTAS","ACTA_NUEVA","Colaborador: "+target.nombre+" ("+target.area+"), Tipo: "+tipoActa,folio);
  const destinosActa = [...new Set([..._getEquipoNotif("CH"), perfilSolicitante.email])].filter(e => e && e.includes("@"));
  _mailActa([...new Set(destinosActa)], folio, fecha, hora, target, { solicitanteNombre: perfilSolicitante.nombre, areaSolicitante: areaJefe, descripcion: payload.descripcion });
  return { folio, nombre: target.nombre, area: target.area };
}

function obtenerHistorialActas(modoAdmin, emailUsuario) {
  try {
    const ss = SpreadsheetApp.openById(SS_MASTER_ID);
    const email = Session.getActiveUser().getEmail().toLowerCase().trim();
    let esAdminReal = false;
    if (modoAdmin) {
      if (_getSuperAdmins().includes(email)) { esAdminReal = true; }
      else {
        const shU = ss.getSheetByName(SH_USUARIOS);
        if (shU && shU.getLastRow() >= 2) {
          const uData = shU.getRange(2, 2, shU.getLastRow()-1, 11).getValues();
          for (let i = 0; i < uData.length; i++) {
            if (String(uData[i][0]).trim().toLowerCase() === email) { esAdminReal = normalizarBooleano(uData[i][10]); break; }
          }
        }
      }
    }
    const sh = ss.getSheetByName(SH_ACTAS);
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues();
    const result = [];
    const filtroEmail = (modoAdmin && esAdminReal) ? null : email;
    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      if (!String(r[0]).trim()) continue;
      if (filtroEmail && String(r[3]).toLowerCase().trim() !== filtroEmail) continue;
      let fechaStr = "";
      try { fechaStr = (r[1] instanceof Date) ? Utilities.formatDate(r[1], "GMT-6", "dd/MM/yyyy") : String(r[1]); } catch(fe) { fechaStr = String(r[1]); }
      result.push({ folio: String(r[0]), fecha: fechaStr, hora: String(r[2]), solicitante: String(r[4]), idRep: String(r[5]), nombre: String(r[6]), puesto: String(r[7]), area: String(r[8]), tipo: String(r[9]), desc: String(r[10]), estatus: String(r[11]), areaSol: String(r[13]) });
    }
    return result;
  } catch(e) { throw new Error("Error actas: " + e.message); }
}

function obtenerActasPorEmpleado(idEmp) {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_ACTAS);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1).filter(r => String(r[5]).trim() === String(idEmp).trim() && String(r[0]).trim() !== "").map(r => ({ folio:r[0], fecha:r[1] instanceof Date ? Utilities.formatDate(r[1],"GMT-6","dd/MM/yyyy") : String(r[1]), tipo:r[9], desc:r[10], estatus:r[11] })).reverse();
  } catch(e) { return []; }
}

function actualizarEstatusActa(folio, nuevoEstatus) {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_ACTAS);
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(folio)) {
        sh.getRange(i+2,12).setValue(nuevoEstatus);
        if (nuevoEstatus === "RESUELTA") _mailActaResuelta(String(data[i][3]), String(data[i][4]), folio);
        registrarAuditoria("ACTAS","CAMBIO_ESTATUS","Folio: "+folio+" → "+nuevoEstatus,folio);
        return { success:true };
      }
    }
    return { success:false, msg:"Folio no encontrado." };
  } catch(e) { return { success:false, msg:e.message }; }
}

function _mailActa(destinos, folio, fecha, hora, target, payload) {
  try {
    const html = `<div style="font-family:Arial;background:#F7F9FC;padding:32px"><div style="max-width:600px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.1)"><div style="background:linear-gradient(135deg,#FF6B00 0%,#FF8C00 100%);padding:28px 32px;position:relative"><div style="position:absolute;top:0;left:0;right:0;height:4px;background:linear-gradient(90deg,#FF6B00,#5DC1E8)"></div><table width="100%"><tr valign="middle"><td><div style="font-size:10px;letter-spacing:3px;color:rgba(255,255,255,.7);font-weight:700;margin-bottom:8px">ACTAS ADMINISTRATIVAS</div><div style="font-size:24px;font-weight:900;color:#fff">NUEVA ACTA</div></td><td align="right"><div style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:12px;padding:12px 16px;text-align:center"><div style="font-size:9px;color:rgba(255,255,255,.7);letter-spacing:2px">FOLIO</div><div style="font-size:12px;color:#fff;font-weight:900;font-family:monospace">${folio}</div></div></td></tr></table></div><div style="padding:28px 32px"><table width="100%" style="margin-bottom:18px"><tr><td width="50%" style="padding-right:8px"><div style="background:#FFF7ED;border-left:4px solid #FF6B00;border-radius:8px;padding:14px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:4px">REPORTADO</div><div style="font-size:15px;font-weight:800;color:#0F172A">${target.nombre}</div><div style="font-size:11px;color:#64748B;margin-top:2px">${target.puesto} · ${target.area}</div></div></td><td width="50%" style="padding-left:8px"><div style="background:#F0F9FF;border-left:4px solid #5DC1E8;border-radius:8px;padding:14px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:4px">SOLICITADO POR</div><div style="font-size:15px;font-weight:800;color:#0F172A">${payload.solicitanteNombre}</div></div></td></tr></table><div style="background:#FFF7ED;border-radius:10px;padding:14px;margin-top:16px"><div style="font-size:9px;letter-spacing:2px;color:#64748B;margin-bottom:6px;font-weight:700">MOTIVO</div><div style="font-size:13px;color:#0F172A;line-height:1.6">${payload.descripcion||"Sin descripcion"}</div></div><div style="background:#F0F9FF;border-radius:8px;padding:10px;margin-top:12px;text-align:center;font-size:11px;color:#64748B">Folio: <b style="font-family:monospace;color:#0F172A">${folio}</b> · ${fecha} ${hora}</div></div></div></div>`;
    MailApp.sendEmail({ to: destinos.join(","), name: "Sistema Corporativo · Capital Humano", replyTo: _getReplyTo("CH"), subject: `Acta ${folio} · ${target.nombre}`, htmlBody: html });
  } catch(e) { throw new Error("EMAIL_ACTA: " + e.message); }
}

function _mailActaResuelta(correo, nombre, folio) {
  try {
    const html = _emailShell({ headerGrad: "linear-gradient(135deg,#22C55E 0%,#16A34A 100%)", accent: "#22C55E", title: "ACTA RESUELTA", subtitle: "Capital Humano ha cerrado este caso", folioValue: folio, body: _emailInfoBlock("COLABORADOR", nombre, "#22C55E") + '<div style="background:#F0FDF4;border:1px solid #BBF7D0;border-radius:10px;padding:18px;margin-top:14px;text-align:center"><div style="font-size:28px;margin-bottom:8px">&#10003;</div><div style="font-size:14px;font-weight:700;color:#166534">Este caso ha sido resuelto satisfactoriamente</div><div style="font-size:12px;color:#4ADE80;margin-top:4px">No se requiere ninguna acción adicional.</div></div>', footerNote: "Este es un correo automático del sistema corporativo." });
    MailApp.sendEmail({ to: correo, name: "Sistema Corporativo · Capital Humano", replyTo: _getReplyTo("CH"), subject: `Acta ${folio} resuelta`, htmlBody: html });
  } catch(e) { throw new Error("EMAIL_ACTA_RESUELTA: " + e.message); }
}

/* ═══════════════════════════════════════════════════════════════
   ☣️ MÓDULO · RESIDUOS PELIGROSOS
═══════════════════════════════════════════════════════════════ */
function calcularFolioResiduos() {
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_RESIDUOS);
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
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_RESIDUOS);
    const { nextFolio } = calcularFolioResiduos();
    const now = new Date();
    const fecha = Utilities.formatDate(now,"GMT-6","dd/MM/yyyy");
    const hora = Utilities.formatDate(now,"GMT-6","HH:mm");
    const row = [ nextFolio, "1", fecha, hora, payload.correo, payload.nombre, payload.area, payload.tipo, payload.descripcion, "", "", "", "", "PENDIENTE", "" ];
    sh.appendRow(row);
    let emailErrRes = "";
    try { _enviarAlertaResiduos(nextFolio, payload.nombre, payload.area, payload.descripcion, payload.tipo); } catch(me){ emailErrRes = me.message; Logger.log("⚠️ Error email residuos: " + me.message); }
    registrarAuditoria("RESIDUOS","SOLICITUD_NUEVA","Area: "+payload.area+", Tipo: "+payload.tipo+", "+payload.descripcion, nextFolio);
    return { success: true, folio: nextFolio, emailError: emailErrRes };
  } catch(e) { return { success: false, msg: e.message }; } finally { lock.releaseLock(); }
}

function obtenerHistorialResiduos(esAdmin) {
  try {
    const perfil = obtenerPerfilCompleto();
    const mostrarTodo = perfil.acceso && perfil.residuos.admin && esAdmin;
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_RESIDUOS);
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
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_RESIDUOS);
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(folio)) {
        sh.getRange(i+2, 14).setValue(nuevoEstatus);
        if (nuevoEstatus === "RECOLECTADO") {
          sh.getRange(i+2, 15).setValue(Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm"));
        }
        registrarAuditoria("RESIDUOS","CAMBIO_ESTATUS","Folio: "+folio+" -> "+nuevoEstatus,folio);
        return { success: true };
      }
    }
    return { success: false, msg: "Folio no encontrado." };
  } catch(e) { return { success: false, msg: e.message }; }
}

function _enviarAlertaResiduos(folio, solicitante, area, desc, tipo) {
  const destinos = _getEquipoResiduos().join(",");
  const urlPesar = `${getWebAppUrl()}?action=pesar_residuos&folio=${folio}`;
  const body = '<table width="100%" cellpadding="0" cellspacing="0"><tr><td width="50%" style="padding-right:6px">' + _emailInfoBlock("ÁREA", area, "#EF4444") + '</td><td width="50%" style="padding-left:6px">' + _emailInfoBlock("GENERADOR", solicitante, "#54A9C8") + '</td></tr></table>' +
    '<div style="background:#FEF2F2;border:1px solid #FECACA;border-radius:10px;padding:16px;margin:12px 0"><div style="font-size:16px;font-weight:800;color:#991B1B;margin-bottom:4px">' + desc + '</div><div style="font-size:12px;color:#EF4444">Tipo: ' + tipo + '</div></div>' +
    '<div style="font-size:12px;color:#94A3B8;text-align:center;margin:14px 0">Da clic en el botón cuando recojas para ingresar el peso real:</div>' +
    _emailButton(urlPesar, "IR A PESAR Y RECIBIR", "#EF4444");
  const html = _emailShell({ headerGrad: "linear-gradient(135deg,#EF4444 0%,#B91C1C 100%)", accent: "#EF4444", title: "RECOLECCIÓN PENDIENTE", subtitle: "Residuos Peligrosos", folioValue: folio, body: body });
  MailApp.sendEmail({ to: destinos, name: "Sistema Corporativo · Residuos", replyTo: _getReplyTo("RESIDUOS"), subject: `[RESIDUOS] Recolectar en ${area} (Folio ${folio})`, htmlBody: html });
}

function guardarPesosResiduos(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: "Sistema ocupado, intenta de nuevo." };
  try {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_RESIDUOS);
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
    MailApp.sendEmail({ to: _getEquipoResiduos().join(","), name: "Sistema Corporativo · Residuos", replyTo: _getReplyTo("RESIDUOS"), subject: `[EN ALMACÉN] Folio ${payload.folio}`, htmlBody: html2 });
    registrarAuditoria("RESIDUOS","PESAJE_REGISTRADO","Folio: "+payload.folio+", Peso: "+payload.pesoTotal+" "+payload.unidad, payload.folio);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
  finally { lock.releaseLock(); }
}

/* ═══════════════════════════════════════════════════════════════
   MÓDULO · EXPEDIENTE DIGITAL
═══════════════════════════════════════════════════════════════ */
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





/* ═══════════════════════════════════════════════════════════════
   MANEJADOR DE ENLACES EXTERNOS (correos / HTML GET)
   Sólo acciones de Residuos Peligrosos.
═══════════════════════════════════════════════════════════════ */
function handleCallback(params) {
  const action = params.action;
  const folio  = params.folio;

  if (action === 'pesar_residuos') {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_RESIDUOS);
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 14).getValues();
    let targetRow = -1; let desc = "";
    for(let i=0; i<data.length; i++){ if(String(data[i][0]).trim() === String(folio).trim() && String(data[i][13]).trim().toUpperCase() === "PENDIENTE"){ targetRow = i + 2; desc = data[i][8]; break; } }
    if(targetRow === -1) { return HtmlService.createHtmlOutput(`<div style="text-align:center;padding:50px;font-family:sans-serif;"><h2>Este folio ya fue pesado o no existe.</h2></div>`); }
    return HtmlService.createHtmlOutput(`
      <html><head><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{font-family:'Arial',sans-serif;background:#F1F5F9;margin:0;padding:20px;}.card{max-width:450px;margin:0 auto;background:white;padding:30px;border-radius:12px;box-shadow:0 4px 15px rgba(0,0,0,0.1);}h2{color:#EF4444;margin-top:0;}.desc{background:#FEE2E2;padding:15px;border-radius:8px;color:#991B1B;font-weight:bold;margin-bottom:20px;}.row{display:flex;gap:10px;margin-bottom:15px;}select,input{width:100%;padding:12px;border:2px solid #CBD5E1;border-radius:6px;font-size:15px;}.btn-add{background:#E2E8F0;color:#475569;border:none;padding:10px;width:100%;border-radius:6px;font-weight:bold;cursor:pointer;margin-bottom:20px;}.btn-add:active{background:#CBD5E1;}.envase-row{display:flex;align-items:center;gap:10px;margin-bottom:10px;background:#F8FAFC;padding:10px;border-radius:6px;}.total-box{font-size:24px;font-weight:bold;color:#0F172A;text-align:center;margin:20px 0;padding:15px;border-top:2px dashed #E2E8F0;}.btn-save{background:#EF4444;color:white;border:none;padding:15px;width:100%;border-radius:6px;font-weight:bold;font-size:16px;cursor:pointer;box-shadow:0 4px 10px rgba(239,68,68,0.3);}</style></head>
      <body><div class="card" id="form-container"><h2>Calculadora de Pesaje</h2><p style="color:#64748B;">Folio: <b>${folio}</b></p><div class="desc">${desc}</div><div class="row"><select id="recipiente"><option value="" disabled selected>Tipo de Recipiente...</option><option value="Bolsa">Bolsa</option><option value="Tambo">Tambo</option><option value="Bidón">Bidón</option></select><select id="unidad"><option value="Kg">Kg</option><option value="Litros">Litros</option></select></div><div id="lista-envases"></div><button class="btn-add" onclick="addEnvase()">+ Agregar peso de otro envase</button><div class="total-box">Total: <span id="total-lbl">0.00</span></div><button class="btn-save" onclick="guardarTodo()">Guardar en Almacén</button></div><div id="msg" style="text-align:center; margin-top:30px; font-weight:bold; font-size:18px;"></div>
      <script>let envasesCount=0;function addEnvase(){envasesCount++;const div=document.createElement('div');div.className='envase-row';div.innerHTML=\`<span style="color:#64748B;font-weight:bold;width:80px;">Envase \${envasesCount}:</span><input type="number" step="0.1" class="peso-input" placeholder="0.0" oninput="calcTotal()">\`;document.getElementById('lista-envases').appendChild(div);}function calcTotal(){let total=0;document.querySelectorAll('.peso-input').forEach(input=>{const val=parseFloat(input.value);if(!isNaN(val))total+=val;});document.getElementById('total-lbl').innerText=total.toFixed(2);return total;}function guardarTodo(){const recipiente=document.getElementById('recipiente').value;const unidad=document.getElementById('unidad').value;const total=calcTotal();if(!recipiente){alert('Selecciona el tipo de recipiente (Bolsa, Tambo, etc.)');return;}if(envasesCount===0||total===0){alert('Debes agregar al menos el peso de un envase.');return;}let notas=[];document.querySelectorAll('.peso-input').forEach(input=>{const val=parseFloat(input.value);if(!isNaN(val))notas.push(val);});const payload={folio:'${folio}',recipiente:recipiente,cantidad:envasesCount,pesoTotal:total,unidad:unidad,notas:notas.join(' + ')};document.getElementById('form-container').style.display='none';document.getElementById('msg').innerHTML='Guardando calculo en la base de datos...';document.getElementById('msg').style.color='#F59E0B';google.script.run.withSuccessHandler(function(r){document.getElementById('msg').innerHTML='Guardado exitosamente!<br><span style="font-size:14px;color:#666;font-weight:normal;display:block;margin-top:10px;">La base de datos se ha actualizado limpiamente.<br>Ya puedes cerrar esta ventana.</span>';document.getElementById('msg').style.color='#22C55E';}).guardarPesosResiduos(payload);}addEnvase();</script></body></html>
    `);
  }

  if (action === 'entregar_proveedor') {
    const sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_RESIDUOS);
    const data = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for(let i=0; i<data.length; i++){
      if(String(data[i][0]).trim() === String(folio).trim()){ sh.getRange(i+2, 14).setValue("RECOLECTADO"); sh.getRange(i+2, 15).setValue(Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm")); }
    }
    return HtmlService.createHtmlOutput(`<div style="font-family:sans-serif;text-align:center;padding:50px;background:#F7F9FC;height:100vh;"><div style="background:white;padding:40px;border-radius:15px;box-shadow:0 10px 30px rgba(0,0,0,0.1);max-width:500px;margin:0 auto;"><h1 style="color:#22C55E;margin-top:0;">ENTREGADO AL PROVEEDOR</h1><p style="font-size:18px;">El folio <b>${folio}</b> ha sido entregado exitosamente.</p></div></div>`);
  }

  return HtmlService.createHtmlOutput("<h1>Acción no reconocida</h1>");
}


/* ═══════════════════════════════════════════════════════════════
   🔎 DIAGNÓSTICO — En el editor de Apps Script: Ejecutar ▸ diagnostico
   Revisa el resultado (o Ver ▸ Registro). Te dice si el sistema detecta
   tu correo y tu fila en DATA_USUARIOS. Útil cuando "no carga tu nombre".
═══════════════════════════════════════════════════════════════ */
function diagnostico() {
  var out = {};
  try { out.correoDetectado = Session.getActiveUser().getEmail(); } catch(e){ out.correoDetectado = 'ERROR: ' + e.message; }
  try { out.correoEfectivo  = Session.getEffectiveUser().getEmail(); } catch(e){ out.correoEfectivo = 'ERROR: ' + e.message; }
  out.idHojaMaestra = SS_MASTER_ID;
  try {
    var sh = SpreadsheetApp.openById(SS_MASTER_ID).getSheetByName(SH_USUARIOS);
    if (!sh) { out.hojaUsuarios = 'NO EXISTE la hoja ' + SH_USUARIOS; }
    else {
      out.hojaUsuarios = 'OK · ' + (sh.getLastRow() - 1) + ' usuarios · ' + sh.getLastColumn() + ' columnas';
      var email = String(out.correoDetectado || '').toLowerCase().trim();
      out.totalSuperAdmins = _getSuperAdmins().length;
      if (sh.getLastRow() >= 2) {
        var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
        var fila = null;
        for (var i = 0; i < data.length; i++) {
          if (String(data[i][1]).toLowerCase().trim() === email) { fila = { fila: i + 2, ne: data[i][0], correo: data[i][1], nombre: data[i][2] }; break; }
        }
        out.miFila = fila || ('NO ENCONTRADA para el correo "' + email + '"  (revisa que esté escrito igual en la columna CORREO)');
      }
    }
  } catch(e){ out.hojaUsuarios = 'ERROR: ' + e.message; }
  try { out.perfil = obtenerPerfilCompleto(true); } catch(e){ out.perfil = 'ERROR: ' + e.message; }
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

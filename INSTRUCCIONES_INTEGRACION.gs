/**
 * ============================================================================
 * INSTRUCCIONES DE INTEGRACIÓN — CLEAVER DISC v3.2
 * Sistema SOLU · Capital Humano
 * ============================================================================
 * QUÉ CAMBIAR EN TU PROYECTO GAS EXISTENTE
 * ============================================================================
 *
 * PASO 1 — REEMPLAZAR doGet() en tu archivo principal .gs
 * ─────────────────────────────────────────────────────────
 * BORRA la función doGet() actual y PON esto en su lugar:
 */

function doGet(e) {
  const modo = (e && e.parameter && e.parameter.modo) ? e.parameter.modo : 'reportes';

  if (modo === 'aplicacion') {
    // Candidatos / Personal — prueba DISC
    return HtmlService
      .createTemplateFromFile('Index_Aplicacion')
      .evaluate()
      .setTitle('SOLU · Evaluación DISC — CLEAVER')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport','width=device-width,initial-scale=1.0');
  }

  // RRHH → dashboard de reportes (comportamiento original intacto)
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SOLU · Capital Humano — CLEAVER Reportes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width,initial-scale=1.0');
}

/**
 * ============================================================================
 * PASO 2 — AGREGAR ARCHIVOS NUEVOS AL MISMO PROYECTO
 * ─────────────────────────────────────────────────────
 * En el editor de Apps Script, botón "+" → Archivo:
 *
 *   → Nuevo archivo .gs   → pegar contenido de CLEAVER_APLICACION.gs
 *   → Nuevo archivo HTML  → pegar contenido de Index_Aplicacion.html
 *
 * NO toques nada más del proyecto existente.
 *
 * ============================================================================
 * PASO 3 — EJECUTAR UNA SOLA VEZ (inicialización)
 * ─────────────────────────────────────────────────
 * En el editor, selecciona función: inicializarHojasApp → Ejecutar
 * Crea automáticamente:
 *   → Hoja: Tokens_Candidatos     (token | correo | nombre | puesto | usado | fecha)
 *   → Hoja: Dominios_Permitidos   (dominio | descripción)
 *
 * ============================================================================
 * PASO 4 — AGREGAR TUS DOMINIOS
 * ──────────────────────────────
 * Abre hoja "Dominios_Permitidos" y agrega uno por fila:
 *
 *   Dominio              | Descripción
 *   tuempresa.com        | Dominio principal
 *   otrodominio.com      | Dominio secundario
 *
 * ============================================================================
 * PASO 5 — SINCRONIZAR TOKENS
 * ────────────────────────────
 * Cada vez que registres candidatos en "Nuevos_Ingresos":
 * Selecciona función: sincronizarTokensDesdeNuevosIngresos → Ejecutar
 *
 * ============================================================================
 * PASO 6 — REPUBLICAR EL WEBAPP
 * ──────────────────────────────
 * Implementar → Administrar implementaciones → editar la existente
 * → Nueva versión → Guardar
 * (La URL no cambia, solo se actualiza el código)
 *
 * ============================================================================
 * URLS RESULTANTES
 * ─────────────────
 * Capital Humano (dashboard actual, sin cambios):
 *   https://script.google.com/macros/s/TU_ID/exec
 *
 * Candidatos / Personal (prueba DISC):
 *   https://script.google.com/macros/s/TU_ID/exec?modo=aplicacion
 *
 * ============================================================================
 */

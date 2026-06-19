# 🗄️ Archivo de memoria — Respaldo del sistema original

Este directorio conserva el sistema **completo, sin recortes**, tal como estaba
antes de la reducción a 3 módulos. Nada se perdió: aquí está todo el código de los
módulos retirados, listo para consultarse o restaurarse.

## Contenido

| Archivo | Qué es |
|---------|--------|
| `Sistema_Corporativo.ORIGINAL.gs` | Backend original completo (1322 líneas, todos los módulos). |
| `Sistema_Corporativo.ORIGINAL.Index.html` | Frontend original completo (4894 líneas, todos los módulos). |
| `indexx_sheetnet_placeholder.html` | Placeholder antiguo "SheetNet Pro" (no relacionado con el sistema). |

## Módulos retirados (y dónde encontrarlos en el original)

### Backend — `Sistema_Corporativo.ORIGINAL.gs`
- **Compras** — `calcularFolioCompra`, `procesarSolicitudCompra`, `obtenerPedidosVisor`,
  `actualizarEstatusCompra`, `_mailCompra`, `_getAprobadoresCompras`, `resolverJefeArea`.
- **Soporte TI** — `calcularFolioSoporte`, `registrarSoporte`, `obtenerHistorialSoporte`,
  `actualizarEstatusSoporte`, `_mailSoporte`, `_mailSoporteResuelta`.
- **Mantenimiento** — `calcularFolioMantenimiento`, `registrarMantenimiento`,
  `obtenerHistorialMantenimiento`, `actualizarEstatusMantenimiento`,
  `_enviarAlertaMantenimiento`, `_mailConfirmacionMantenimiento`, `enviarRecordatoriosMantenimiento`.
- **Limpieza de Equipos** — `_getCatalogoInfraCompleto`, `obtenerCatalogoInfra`,
  `obtenerDatosLimpieza`, `_normFH`, `obtenerHorasOcupadasLimpieza`, `registrarLimpieza`,
  `resolverLimpieza`, `cancelarLimpieza`, `posponerLimpieza`.
- **Callbacks de correo** retirados de `handleCallback`: `autorizar_compra`, `enviar_omar`,
  `enviar_isela`, `confirmar_omar`, `confirmar_isela`, `resolver_soporte`,
  `confirmar_cierre_mant`, `updateMant`.

### Frontend — `Sistema_Corporativo.ORIGINAL.Index.html`
- Vistas: `vComprasMenu`, `vFormCompra`, `vHistCompra`, `vAdminCompra`,
  `vSoporteMenu`, `vFormSoporte`, `vMisSoporte`, `vAdminSoporte`,
  `vMantenimientoMenu`, `vFormMantenimiento`, `vMisMantenimiento`, `vAdminMantenimiento`,
  `vFormLimpieza`, `vMisLimpieza`, `vAdminLimpieza`.
- Tarjeta "Transporte" (módulo *próximo*, placeholder).
- Modales: cancelar-compra (`obsModal*`) y posponer-limpieza (`pospModal*`).

## Nota sobre datos
Las columnas de permisos de estos módulos en la hoja `DATA_USUARIOS` **siguen existiendo**
(no se borraron) para no alterar la integridad de la base. El sistema activo simplemente
ya no las lee ni las muestra.

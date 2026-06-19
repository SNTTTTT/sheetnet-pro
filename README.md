# Sistema Corporativo · Santillán Ramírez — V5.0 (Núcleo Blindado)

Aplicación Google Apps Script (Web App) reducida, blindada y optimizada.
Se conservaron **únicamente 3 módulos** y se retiraron el resto. Todo el código
retirado queda **respaldado íntegro** en [`_archivo_memoria/`](./_archivo_memoria).

---

## 🧩 Módulos activos

| # | Módulo | Descripción |
|---|--------|-------------|
| 01 | **Actas Administrativas** | Registro de incidencias laborales y resolución por Capital Humano. |
| 02 | **Expediente Digital** | Alta y gestión de expedientes (50 campos), carga de documentos a Drive y escáner de INE. |
| 03 | **Residuos Peligrosos** | Solicitud de recolección, pesaje en almacén y entrega a proveedor. |

Más: **Administración de Usuarios** y **Bitácora / Auditoría** (solo Super-Admin).

### Módulos retirados (respaldados en `_archivo_memoria/`)
Compras · Soporte TI · Mantenimiento · Limpieza de Equipos · Transporte.

---

## 📁 Estructura del repositorio

```
Codigo.gs            → Backend Apps Script (núcleo + 3 módulos + usuarios + auditoría)
Index.html           → Frontend (SPA). En Apps Script el archivo debe llamarse "Index"
appsscript.json      → Manifiesto (zona horaria, runtime V8, scopes OAuth)
_archivo_memoria/    → Respaldo íntegro del sistema original (TODOS los módulos)
   ├─ Sistema_Corporativo.ORIGINAL.gs
   ├─ Sistema_Corporativo.ORIGINAL.Index.html
   ├─ indexx_sheetnet_placeholder.html   (placeholder antiguo, no relacionado)
   └─ _LEEME.md                          (qué se retiró y dónde estaba)
```

---

## 🚀 Despliegue en Google Apps Script

1. Abre el proyecto Apps Script ligado a tu hoja maestra (`SS_MASTER_ID`).
2. Pega el contenido de **`Codigo.gs`** en el archivo de script.
3. Crea/usa un archivo HTML llamado **`Index`** y pega **`Index.html`**.
   > El backend hace `HtmlService.createTemplateFromFile("Index")`, por eso el
   > nombre **debe** ser `Index` (sin extensión en el editor).
4. (Opcional, si usas `clasp`) `appsscript.json` ya trae los scopes necesarios.
5. **Implementar → Nueva implementación → Aplicación web.**
   - *Ejecutar como:* el propietario/administrador (para escribir en la hoja central).
   - *Quién tiene acceso:* tu dominio corporativo.

### Permisos / OAuth
Sheets · Drive · Envío de correo · Email del usuario · Solicitudes externas · ScriptApp.

---

## 🔐 Modelo de accesos

- Dominios permitidos: `@ssolu.com.mx`, `@bur.mx`, `@bursatron.com.mx`.
- Los permisos por usuario viven en `DATA_USUARIOS`. **El esquema de 25 columnas
  se conserva intacto** por integridad de datos: las columnas de los módulos
  retirados (Compras, Soporte, Mantenimiento, Transporte, Limpieza) **no se borran**
  de la hoja, pero el sistema ya **no las expone ni las usa**.
- `actualizarUsuario()` está **blindado**: lee la fila previa y **preserva** esas
  columnas heredadas, de modo que editar un usuario nunca destruye datos.
- Equipos de notificación (columna 26 `equipoNotificacion`): `SUPER_ADMIN`, `CH`
  (Capital Humano → actas), `RESIDUOS`/`MANT` (recolección de residuos).

---

## ⚡ Velocidad, limpieza y blindaje

- **Frontend:** caché en `sessionStorage` (`AppCache`), *debounce* en búsquedas,
  render bajo demanda por vista, animaciones con `will-change`, guardas `try/catch`
  y validación de campos antes de cada envío.
- **Backend:** `CacheService` (perfil de usuario y catálogo de áreas), `LockService`
  en todas las escrituras concurrentes (residuos, expedientes), búsqueda de fila por
  folio (nunca se confía en índices del cliente) y `try/catch` con mensajes claros.
- **Cero código muerto:** se eliminaron vistas, rutas, handlers, modales y variables
  de los módulos retirados; el contrato frontend↔backend quedó verificado al 100 %.

---

## ♻️ Restaurar un módulo retirado
El código original completo está en `_archivo_memoria/Sistema_Corporativo.ORIGINAL.gs`
y `…ORIGINAL.Index.html`. Cada módulo está delimitado por comentarios
(`/* MÓDULO 01 COMPRAS */`, etc.); basta copiar su backend, su vista `vXxx`, sus
handlers JS y volver a registrar la ruta en `nav()` y su tarjeta en el lobby.

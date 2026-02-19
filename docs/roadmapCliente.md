---
title: Roadmap Cliente (Validación dinámica)
project: DistribuidoraPalomarNet
scope: frmDatosUsuario + frmPrincipal + Sync/Actualización de base
---

# Objetivo
Implementar una **validación dinámica de clientes** basada en la base local **`BASEPALOMARNET.mdb`** (tabla `clientes`), para asegurar:

- Identidad: **`CLCODI` + `CLCUIT`** (coincidencia exacta)
- Habilitación comercial: **`CLDES1` y `CLDES2`**

El sistema debe soportar 3 estados:

- **Autorizado**: `CLCODI+CLCUIT` válido y descuentos > 0
- **Restringido**: `CLCODI+CLCUIT` válido pero descuentos = 0
- **No autorizado**: `CLCODI+CLCUIT` inválido o no existe

# Consulta de validación (fuente de verdad)

```sql
SELECT C.CLRAZO, C.CLCUIT, C.CLDES1, C.CLDES2
FROM [clientes] C
WHERE C.CLCODI = @CodigoIngresado
  AND C.CLCUIT = @CuitIngresado;
```

# Estado actual del proyecto (hallazgos)

## Formularios
- **`frmDatosUsuario`** ya pide `Código` y `CUIT`.
- Validaciones actuales en UI llaman:
  - `Module.ValidarCodiCliente(...)`
  - `Module.ValidarCuitCliente(codi, cuit)`

## Módulos
- Existe persistencia local en tabla `USUARIO` (DB usuario) con:
  - `US_CODI`, `US_CUIT`, `US_DESCUENTO1`, `US_DESCUENTO2`, etc.
- `ModuloUsuario.ValidarCuitCliente(...)` actualmente valida contra `clientes` pero **no lee `CLDES1/CLDES2`**.

## Ofertas / Catálogo
- `frmPrincipal` ya marca filas como “oferta” con `artt > 0`.

# Arquitectura propuesta (componentes)

## 1) `ClienteRepository` (Acceso a datos)
Responsabilidad:
- Ejecutar el SELECT con parámetros por `CLCODI + CLCUIT`.
- Devolver un DTO con `CLRAZO`, `CLCUIT`, `CLDES1`, `CLDES2`.

Notas técnicas:
- Usar `OleDbCommand` parametrizado.
- Cuidar tipos: Access suele guardar CUIT como texto o numérico. Definirlo por inspección de esquema.

## 2) `UserSessionManager` (estado global)
Responsabilidad:
- Guardar en memoria (y cargar desde tabla `USUARIO`) los campos:
  - `Codigo`, `Cuit`
  - `Descuento1`, `Descuento2`
  - `IsAuthenticated` (dupla válida)
  - `IsAuthorized` (dupla válida y descuentos > 0)
  - `IsRestricted` (dupla válida y descuentos = 0)
- Ofrecer:
  - `ValidateFromLocalDb()` (revalidación completa)
  - `ApplySessionToModule()` (puente temporal con `Module.US_*` mientras se migra)

## 3) Guardia de UI (catálogo / carrito)
Responsabilidad:
- Centralizar el “bloqueo” del catálogo y carrito si:
  - `IsAuthorized == false` (incluye restringido y no autorizado)

Implementación sugerida (WinForms):
- Método común, ej.: `ApplyAuthorizationStateToUI()` en `frmPrincipal`.
- Deshabilitar controles:
  - `DataGridView1`, `DataGridView2` (catálogo)
  - botones “agregar”, “comprar”, “ver carrito”, etc.
- Mostrar aviso:
  - `IsRestricted`: “Notifique a la Distribuidora que no tiene descuentos aplicados”.
  - `NoAuth`: “Datos inválidos”.

## 4) Integración con Sync / Actualización de base
Responsabilidad:
- Después de `Module.ActualizacionXFTP()` (o flujo equivalente):
  - revalidar `CLCODI+CLCUIT` contra la **nueva** base descargada.
  - si pasa a `IsAuthorized=false`, cortar acceso inmediatamente.

Nota crítica:
- Evitar que la MDB esté bloqueada.
  - Asegurar que `OleDbConnection` esté `Close()` / `Dispose()` antes de reemplazar archivo.

# Workflow funcional (criterios)

## A) Primer ingreso
1. Usuario ingresa `Código` + `CUIT` en `frmDatosUsuario`.
2. Se valida contra DB local.
3. Resultado:
- No existe dupla: bloqueado.
- Existe dupla y descuentos 0: modo restringido (solo sync).
- Existe dupla y descuentos > 0: acceso total.

## B) Validación recurrente (3 puntos)
- **Al iniciar la app**: `ValidateFromLocalDb()` y setear estado.
- **Post-sincronización**: revalidar y aplicar guard.
- **Antes de mostrar catálogo** (guard): si no autorizado, no permitir navegar/comprar.

# Matriz de permisos

| Escenario | Catálogo | Precios/Carrito | Descarga Base | Mensaje |
|---|---:|---:|---:|---|
| Código+CUIT correcto (con desc.) | Sí | Sí | Sí | - |
| Código+CUIT correcto (desc. 0) | No | No | Sí | “Contacte Distribuidora” |
| Error en Código o CUIT | No | No | Sí | “Datos inválidos” |

# Plan de trabajo (milestones)

## Estado de implementación (snapshot)
- M3 (frmDatosUsuario): implementado (revalidación + refresh de `frmPrincipal` al validar/grabar/salir).
- M4 (guardia UI en frmPrincipal): implementado (bloqueo catálogo/carrito si `ClienteAutorizado=false`).
- Validación al iniciar: implementado en `Module.Cargar_Usuario()` llamando `ValidarCuitCliente(US_CODI, US_CUIT)`.
- M5 (post-sync): implementado en `Module.ActualizacionXFTP()` tras `UpdateBase(...)`/`Cargar_Usuario()` con `ApplyEstadoClienteToUI()`.

## M1 — Repositorio de clientes (DB Access)
- Implementar `ClienteRepository`.
- Query parametrizada `CLCODI+CLCUIT`.
- DTO con `CLRAZO`, `CLCUIT`, `CLDES1`, `CLDES2`.

**Estimación**: 3h

## M2 — Session manager + estado global
- Implementar `UserSessionManager`.
- Persistencia de credenciales:
  - lectura desde tabla `USUARIO` al iniciar.
- Revalidación contra `clientes`.
- Determinar estado: autorizado / restringido / no auth.

**Estimación**: 4h

## M3 — `frmDatosUsuario` (login/registro y feedback)
- Unificar validación:
  - validar `Código` + `CUIT` juntos (no solo por separado).
- Mostrar resultado (OK / restringido / inválido).
- Si restringido:
  - permitir guardar usuario (si corresponde) pero avisar que catálogo quedará bloqueado.

**Estado**: implementado.

**Estimación**: 6h (2h backend + 4h UI)

## M4 — Guardia de UI en `frmPrincipal`
- Implementar `ApplyAuthorizationStateToUI()`.
- Bloquear:
  - catálogo (grillas)
  - botones de compra/carrito
- Permitir:
  - sincronización/actualización base
- Mensajes consistentes.

**Estado**: implementado.

**Estimación**: 8h (4h backend + 4h UI)

## M5 — Integración con Sync (post-FTP)
- En el flujo de actualización (actualmente `Module.ActualizacionXFTP()`):
  - al finalizar descarga y swap de base, ejecutar `ValidateFromLocalDb()`.
  - si cambia a no autorizado/restringido, aplicar guard en caliente.

**Estado**: implementado.

**Estimación**: 4h (3h backend + 1h UI)

## M6 — Testing y casos borde
- CUIT inválido (no numérico / formato)
- Dupla válida, descuentos 0
- Dupla válida, descuentos > 0
- Cliente eliminado en base nueva (post-sync)
- Reemplazo de MDB con conexiones abiertas (archivo en uso)

**Estimación**: 5h (3h backend + 2h UI)

# Estimación total
- Backend: 19h
- Frontend: 11h
- **Total: 30h**

# Riesgos / decisiones pendientes
- Tipo de dato real de `CLCUIT` en Access (texto vs numérico): afecta parámetros.
- Dónde está físicamente `BASEPALOMARNET.mdb` y cómo se reemplaza en el sync.
- Definir si se permite “navegar” catálogo en modo restringido pero sin comprar (hoy el plan dice bloqueo total).
- Seguridad local: si se encripta o no el CUIT persistido.

# Definición de terminado (DoD)
- La app revalida identidad y descuentos al iniciar.
- Luego de sincronizar base, revalida y aplica restricciones inmediatamente.
- `frmPrincipal` bloquea catálogo/carrito si `IsAuthorized=false`.
- Mensajes correctos según matriz.
- Sin errores de archivo en uso durante actualización de MDB.

# Plan de Pruebas - Sistema de Equivalencias y Búsqueda Multi-Campo

## Información General
- **Fecha de creación**: 2026-08-22
- **Componentes a probar**: 
  - Sistema de equivalencias (DataGridView2)
  - Búsqueda multi-campo en txtDescripcion
  - Actualización automática por SelectionChanged
  - Sistema de limpieza y flags de control

---

## 1. PRUEBAS DE EQUIVALENCIAS - Carga Automática

### Test 1.1: Selección por Click
**Objetivo**: Verificar que al hacer clic en un artículo se carguen sus equivalencias correctamente

**Pasos**:
1. Abrir la aplicación y esperar que cargue la grilla principal (DataGridView1)
2. Hacer clic en cualquier artículo que tenga equivalencias (verificar que AROBSE no esté vacío)
3. Observar la grilla de equivalencias (DataGridView2)

**Resultado esperado**:
- ✅ DataGridView2 muestra las equivalencias del artículo seleccionado
- ✅ El cuadro amarillo muestra "Observación" y "Equivalencia" del artículo seleccionado
- ✅ Los datos corresponden exactamente al artículo clickeado (NO al anterior)

**Criterio de fallo**:
- ❌ Muestra equivalencias del artículo anterior
- ❌ Necesita más de un click para actualizar
- ❌ No muestra nada o muestra error

---

### Test 1.2: Selección por Teclado (flechas)
**Objetivo**: Verificar que al navegar con teclado se actualicen las equivalencias

**Pasos**:
1. Hacer clic en cualquier artículo de DataGridView1
2. Usar las teclas de flecha ARRIBA/ABAJO para navegar entre artículos
3. Observar DataGridView2 y el cuadro amarillo en cada cambio

**Resultado esperado**:
- ✅ Las equivalencias se actualizan inmediatamente al cambiar de fila
- ✅ El cuadro amarillo se actualiza con los datos del artículo actual
- ✅ No hay retraso ni necesidad de hacer clic adicional

**Criterio de fallo**:
- ❌ Las equivalencias no se actualizan al navegar con teclado
- ❌ Hay un desfase de una fila

---

### Test 1.3: Equivalencias después de Filtrar
**Objetivo**: Verificar que después de filtrar, el primer artículo muestra sus equivalencias

**Pasos**:
1. Escribir en txtCodigo: "11" (o cualquier código que exista)
2. Esperar que se cargue la grilla filtrada
3. Observar DataGridView2 y cuadro amarillo

**Resultado esperado**:
- ✅ El primer artículo de la grilla filtrada queda seleccionado automáticamente
- ✅ DataGridView2 muestra las equivalencias del primer artículo (si tiene)
- ✅ El cuadro amarillo muestra los datos del primer artículo

**Criterio de fallo**:
- ❌ No muestra equivalencias aunque el artículo las tenga
- ❌ Muestra equivalencias de otro artículo

---

## 2. PRUEBAS DE LIMPIEZA - Estado Vacío

### Test 2.1: Limpieza con botón btnLimpiar
**Objetivo**: Verificar que al limpiar se borren equivalencias y cuadro amarillo

**Pasos**:
1. Seleccionar un artículo con equivalencias
2. Verificar que DataGridView2 y cuadro amarillo tengan datos
3. Hacer clic en el botón "Limpiar" (escobita)
4. Observar DataGridView2 y cuadro amarillo

**Resultado esperado**:
- ✅ DataGridView1 queda vacía (solo headers)
- ✅ DataGridView2 queda vacía (sin filas)
- ✅ lblObservaciones está vacío
- ✅ lblRubro está vacío
- ✅ lblOrigen está vacío

**Criterio de fallo**:
- ❌ DataGridView2 todavía muestra equivalencias viejas
- ❌ El cuadro amarillo mantiene datos anteriores

---

### Test 2.2: Filtro sin resultados
**Objetivo**: Verificar comportamiento cuando un filtro no encuentra artículos

**Pasos**:
1. Escribir en txtCodigo: "XYZABC999" (código que no existe)
2. Esperar que se ejecute el filtro
3. Observar todas las grillas y campos

**Resultado esperado**:
- ✅ DataGridView1 vacía o sin resultados
- ✅ DataGridView2 vacía
- ✅ Cuadro amarillo vacío
- ✅ No hay error ni excepción

**Criterio de fallo**:
- ❌ Muestra datos antiguos
- ❌ Se produce un error o excepción

---

## 3. PRUEBAS DE BÚSQUEDA MULTI-CAMPO - txtDescripcion

### Test 3.1: Búsqueda de una palabra
**Objetivo**: Verificar que busca en los 3 campos (ARDESC, rbdesc, armode)

**Pasos**:
1. Limpiar todos los filtros
2. Escribir en txtDescripcion: "cazoleta"
3. Esperar resultados

**Resultado esperado**:
- ✅ Muestra artículos que tengan "cazoleta" en:
  - ARDESC (Descripción), O
  - r.rbdesc (Montadora), O
  - a.armode (Línea)
- ✅ Se ven resultados

**Criterio de fallo**:
- ❌ No encuentra artículos que existen con esa palabra
- ❌ Solo busca en un campo

---

### Test 3.2: Búsqueda multi-palabra en campos diferentes
**Objetivo**: Verificar que encuentra artículos aunque las palabras estén en campos diferentes

**Pasos**:
1. Limpiar todos los filtros
2. Escribir en txtDescripcion: "cazoleta strada"
3. Esperar resultados

**Resultado esperado**:
- ✅ Encuentra artículos donde:
  - "cazoleta" está en ARDESC (ej: "CAZOLETA SUSPENSION"), Y
  - "strada" está en armode (ej: "STRADA")
- ✅ Muestra resultados que cumplan con AMBAS palabras

**Dato a validar**: 
- Debe encontrar un artículo que tenga "cazoleta" en descripción y "strada" en línea/modelo

**Criterio de fallo**:
- ❌ No encuentra nada aunque existan artículos así
- ❌ Solo encuentra si ambas palabras están en el mismo campo

---

### Test 3.3: Búsqueda multi-palabra - 3 palabras
**Objetivo**: Verificar búsqueda con más de 2 palabras

**Pasos**:
1. Limpiar todos los filtros
2. Escribir en txtDescripcion: "amortiguador delantero fiat"
3. Esperar resultados

**Resultado esperado**:
- ✅ Encuentra artículos donde:
  - "amortiguador" está en alguno de los 3 campos, Y
  - "delantero" está en alguno de los 3 campos, Y
  - "fiat" está en alguno de los 3 campos
- ✅ Las 3 palabras deben estar presentes (cada una en cualquier campo)

**Criterio de fallo**:
- ❌ No encuentra artículos válidos
- ❌ Requiere que todas las palabras estén en el mismo campo

---

### Test 3.4: Búsqueda con palabras parciales
**Objetivo**: Verificar que el LIKE funciona con palabras parciales

**Pasos**:
1. Escribir en txtDescripcion: "amor del"
2. Esperar resultados

**Resultado esperado**:
- ✅ Encuentra "AMORTIGUADOR DELANTERO" (palabras parciales)
- ✅ No requiere palabras completas

**Criterio de fallo**:
- ❌ No encuentra nada
- ❌ Requiere palabras exactas completas

---

## 4. PRUEBAS DE CÓDIGOS CLICKEABLES EN EQUIVALENCIAS

### Test 4.1: Click en código de equivalencia
**Objetivo**: Verificar que al hacer clic en un código de la grilla de equivalencias, filtra por ese código

**Pasos**:
1. Seleccionar un artículo que tenga equivalencias
2. Observar la columna "Codigo" en DataGridView2 (debe estar en azul/link)
3. Hacer clic en uno de los códigos de equivalencia
4. Observar qué sucede

**Resultado esperado**:
- ✅ txtCodigo se llena con el código clickeado
- ✅ Los demás campos de búsqueda se limpian
- ✅ Se ejecuta un filtro por ese código
- ✅ DataGridView1 muestra el artículo buscado
- ✅ Las equivalencias se actualizan al nuevo artículo seleccionado

**Criterio de fallo**:
- ❌ No hace nada al hacer clic
- ❌ No filtra por el código
- ❌ Muestra error

---

### Test 4.2: Tooltip en códigos
**Objetivo**: Verificar que aparece tooltip "VER" sobre los códigos

**Pasos**:
1. Seleccionar un artículo con equivalencias
2. Pasar el mouse sobre un código en la columna "Codigo" de DataGridView2
3. Esperar 1-2 segundos

**Resultado esperado**:
- ✅ Aparece un tooltip que dice "VER"
- ✅ El cursor cambia a manita (hand cursor)
- ✅ El texto está en azul y subrayado (estilo link)

**Criterio de fallo**:
- ❌ No aparece tooltip
- ❌ El cursor no cambia
- ❌ No tiene estilo de link

---

## 5. PRUEBAS DE CONTROL DE FLAGS Y LIMPIEZA

### Test 5.1: Limpieza no dispara búsquedas
**Objetivo**: Verificar que el flag PorLimpieza evita búsquedas durante limpieza

**Pasos**:
1. Filtrar por algún código (ejemplo: "002")
2. Observar que hay resultados
3. Hacer clic en "Limpiar"
4. Observar el comportamiento

**Resultado esperado**:
- ✅ Los campos de texto se limpian
- ✅ La grilla queda vacía
- ✅ NO se ejecuta una búsqueda automática al limpiar los TextBox
- ✅ Permanece vacío sin "parpadeos" o cargas intermedias

**Criterio de fallo**:
- ❌ Después de limpiar se ve brevemente una búsqueda ejecutándose
- ❌ La grilla se llena con todos los artículos automáticamente

---

### Test 5.2: Escribir después de limpiar
**Objetivo**: Verificar que después de limpiar, se puede filtrar normalmente

**Pasos**:
1. Hacer clic en "Limpiar"
2. Esperar que todo quede vacío
3. Escribir en txtCodigo: "002"
4. Esperar que se ejecute el filtro

**Resultado esperado**:
- ✅ El filtro funciona normalmente
- ✅ Muestra resultados
- ✅ Al seleccionar un artículo, carga sus equivalencias

**Criterio de fallo**:
- ❌ No filtra después de limpiar
- ❌ Se borra todo al hacer clic en la grilla

---

### Test 5.3: Interactuar con grilla después de limpiar y filtrar
**Objetivo**: Verificar que el problema original está resuelto

**Pasos**:
1. Hacer clic en "Limpiar"
2. Escribir en txtCodigo: "11"
3. Esperar que filtre
4. Hacer clic en la grilla (DataGridView1)
5. Navegar con flechas entre artículos

**Resultado esperado**:
- ✅ La grilla NO se borra al hacer clic
- ✅ Las equivalencias se actualizan normalmente
- ✅ No hay borrados inesperados
- ✅ Todo funciona como si nunca se hubiera limpiado

**Criterio de fallo**:
- ❌ Al hacer clic en la grilla, se borra todo
- ❌ Necesita múltiples clics para que funcione
- ❌ Las equivalencias no se cargan

---

## 6. PRUEBAS DE BÚSQUEDA POR CÓDIGO

### Test 6.1: Búsqueda por código completo
**Objetivo**: Verificar búsqueda en arprim+arterc

**Pasos**:
1. Escribir en txtCodigo: "002"
2. Esperar resultados

**Resultado esperado**:
- ✅ Encuentra códigos que contengan "002" en arprim+arterc
- ✅ Encuentra códigos que contengan "002" en arterc solo

**Criterio de fallo**:
- ❌ No encuentra resultados válidos

---

### Test 6.2: Búsqueda por código parcial
**Objetivo**: Verificar que "11" encuentra resultados

**Pasos**:
1. Escribir en txtCodigo: "11"
2. Esperar resultados

**Resultado esperado**:
- ✅ Encuentra códigos como "RE4211", "001123", etc.
- ✅ Busca en arprim+arterc Y también en arterc solo

**Criterio de fallo**:
- ❌ No encuentra nada aunque existan códigos con "11"

---

## 7. PRUEBA INTEGRADA COMPLETA - Flujo Real de Usuario

### Test 7.1: Flujo completo de búsqueda y navegación
**Objetivo**: Simular un uso real completo del sistema

**Pasos**:
1. **Inicio**: Abrir aplicación, dejar que cargue
2. **Búsqueda inicial**: Escribir "cazoleta strada" en txtDescripcion
3. **Verificar resultados**: Ver que encuentra artículos
4. **Seleccionar**: Hacer clic en un artículo → ver equivalencias
5. **Navegar**: Usar flechas para cambiar de artículo → ver que equivalencias se actualizan
6. **Click en equivalencia**: Hacer clic en un código en DataGridView2
7. **Verificar filtro**: Ver que filtra por ese código
8. **Limpiar**: Hacer clic en "Limpiar"
9. **Nueva búsqueda**: Escribir "002" en txtCodigo
10. **Interactuar**: Hacer clic en grilla, navegar con flechas

**Resultado esperado en cada paso**:
- ✅ Paso 2: Encuentra artículos con palabras en campos diferentes
- ✅ Paso 4: Muestra equivalencias del artículo seleccionado
- ✅ Paso 5: Equivalencias se actualizan automáticamente
- ✅ Paso 6: El código clickeado filtra correctamente
- ✅ Paso 7: Muestra el artículo correspondiente al código
- ✅ Paso 8: Todo se limpia (grillas y cuadro amarillo)
- ✅ Paso 9: Filtra normalmente después de limpiar
- ✅ Paso 10: NO hay borrados inesperados, todo funciona normalmente

**Criterio de fallo**:
- ❌ Cualquier paso no cumple con el resultado esperado
- ❌ Se produce algún error o excepción
- ❌ Hay comportamientos inconsistentes

---

## 8. PRUEBAS DE EDGE CASES

### Test 8.1: Artículo sin equivalencias
**Objetivo**: Verificar comportamiento con artículos sin equivalencias

**Pasos**:
1. Buscar y seleccionar un artículo que NO tenga equivalencias (AROBSE vacío)
2. Observar DataGridView2 y cuadro amarillo

**Resultado esperado**:
- ✅ DataGridView2 está vacío (sin error)
- ✅ El cuadro amarillo muestra "Observación: " y "Equivalencia: " vacíos
- ✅ No hay error ni excepción

---

### Test 8.2: Cambio rápido de selección
**Objetivo**: Verificar que múltiples cambios rápidos no causan problemas

**Pasos**:
1. Filtrar para tener varios artículos
2. Mantener presionada la tecla de flecha ABAJO durante 3-5 segundos
3. Observar comportamiento

**Resultado esperado**:
- ✅ La aplicación responde sin colgarse
- ✅ Las equivalencias se actualizan (puede haber delay, pero no error)
- ✅ No hay excepciones

---

### Test 8.3: Búsqueda con caracteres especiales
**Objetivo**: Verificar que no rompe con caracteres especiales

**Pasos**:
1. Escribir en txtDescripcion: "amortiguador's"
2. Esperar resultados

**Resultado esperado**:
- ✅ No produce error SQL
- ✅ Busca correctamente (escapa el apóstrofe)

**Criterio de fallo**:
- ❌ Error de SQL syntax

---

## 9. RESUMEN DE VALIDACIÓN

### Checklist Final
Marcar cada funcionalidad validada:

**Equivalencias**:
- [ ] Carga correcta al hacer clic en artículo
- [ ] Carga correcta al navegar con teclado
- [ ] Carga correcta después de filtrar
- [ ] Se limpia correctamente con btnLimpiar
- [ ] Se limpia cuando no hay selección

**Búsqueda multi-campo**:
- [ ] Busca en ARDESC, rbdesc y armode
- [ ] Encuentra palabras en campos diferentes
- [ ] Funciona con 2+ palabras
- [ ] Funciona con palabras parciales

**Códigos clickeables**:
- [ ] Los códigos son azules y subrayados
- [ ] Tooltip "VER" aparece
- [ ] Cursor cambia a manita
- [ ] Click ejecuta filtro correctamente

**Control de limpieza**:
- [ ] Limpiar no dispara búsquedas automáticas
- [ ] Después de limpiar se puede filtrar normalmente
- [ ] No hay borrados inesperados al interactuar con grilla

**Búsqueda por código**:
- [ ] "11" encuentra resultados
- [ ] "002" encuentra resultados
- [ ] Busca en arprim+arterc y arterc

---

## 10. REPORTE DE BUGS (Si se encuentran)

**Formato para reportar**:

```
Bug #: [número]
Test: [nombre del test donde se encontró]
Descripción: [qué salió mal]
Pasos para reproducir:
1. 
2. 
3. 
Resultado esperado: 
Resultado actual: 
Severidad: [Crítica/Alta/Media/Baja]
```

---

## Notas Finales

- **Tiempo estimado**: 30-45 minutos para pruebas completas
- **Prerequisitos**: Base de datos con artículos y equivalencias reales
- **Herramientas**: Solo la aplicación corriendo
- **Responsable de pruebas**: Mara / QA Team

**Fecha de última actualización**: 2026-08-22

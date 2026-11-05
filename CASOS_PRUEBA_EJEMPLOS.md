# Casos de Prueba con Ejemplos Específicos

Este documento complementa el TEST_PLAN_EQUIVALENCIAS.md con casos de prueba específicos usando datos reales de la base de datos.

---

## Búsquedas Multi-Campo Recomendadas para Probar

### Caso 1: "cazoleta strada"
**Qué debe encontrar**:
- Artículos donde "cazoleta" aparece en ARDESC (descripción)
- Y "strada" aparece en armode (línea/modelo)

**Ejemplo de registro que DEBE encontrar**:
```
ARDESC: "CAZOLETA SUSPENSION DELANTERA"
armode: "STRADA"
r.rbdesc: "FIAT"
```

**Qué validar**:
- ✅ Encuentra artículos aunque las palabras estén en campos diferentes
- ❌ NO debe requerir que ambas palabras estén en el mismo campo

---

### Caso 2: "amortiguador delantero fiat"
**Qué debe encontrar**:
- "amortiguador" en cualquiera de los 3 campos
- "delantero" en cualquiera de los 3 campos
- "fiat" en cualquiera de los 3 campos

**Ejemplo de registro que DEBE encontrar**:
```
ARDESC: "AMORTIGUADOR DELANTERO"
armode: "PALIO"
r.rbdesc: "FIAT"
```

**Qué validar**:
- ✅ Las 3 palabras pueden estar distribuidas en diferentes campos
- ✅ Todas las palabras deben estar presentes

---

### Caso 3: "filtro aceite"
**Qué debe encontrar**:
- Artículos con "filtro" Y "aceite" en cualquier combinación de campos

**Ejemplo de registro que DEBE encontrar**:
```
ARDESC: "FILTRO DE ACEITE MOTOR"
armode: "VARIOS"
r.rbdesc: "MANN"
```

---

### Caso 4: Solo una palabra "cazoleta"
**Qué debe encontrar**:
- Todos los artículos que tengan "cazoleta" en:
  - ARDESC, O
  - armode, O
  - r.rbdesc

---

## Búsquedas por Código Recomendadas

### Caso 5: Código "11"
**Problema anterior**: No encontraba nada
**Ahora debe encontrar**:
- Códigos como: RE4211, 001123, 11ABC, etc.
- Busca en: arprim+arterc Y también en arterc solo

**Qué validar**:
- ✅ Encuentra resultados
- ✅ No importa si "11" está al inicio, medio o final del código

---

### Caso 6: Código "002"
**Debe encontrar**:
- Códigos que contengan "002"
- Ejemplo: 002123, AB002XY, 00245, etc.

---

## Casos de Equivalencias

### Caso 7: Artículo CON equivalencias
**Preparación**:
1. Buscar un artículo que tenga el campo AROBSE lleno
2. Hacer clic en ese artículo

**Qué validar**:
- ✅ DataGridView2 se llena con artículos equivalentes
- ✅ Los códigos en DataGridView2 están en azul
- ✅ Al pasar mouse sobre código, aparece tooltip "VER"
- ✅ Al hacer clic en un código, filtra por ese código

---

### Caso 8: Artículo SIN equivalencias
**Preparación**:
1. Buscar un artículo con AROBSE vacío o NULL

**Qué validar**:
- ✅ DataGridView2 queda vacío (sin error)
- ✅ El cuadro amarillo muestra campos vacíos
- ✅ No hay excepción ni error

---

### Caso 9: Navegación entre artículos con y sin equivalencias
**Preparación**:
1. Filtrar para tener varios artículos
2. Algunos con equivalencias, otros sin equivalencias

**Pasos**:
1. Hacer clic en artículo CON equivalencias → verificar que carga
2. Presionar flecha ABAJO para ir a artículo SIN equivalencias → verificar que limpia
3. Presionar flecha ABAJO para ir a artículo CON equivalencias → verificar que carga nuevamente

**Qué validar**:
- ✅ Las equivalencias se actualizan correctamente en cada cambio
- ✅ No hay residuos de artículos anteriores
- ✅ La transición es limpia

---

## Casos de Limpieza y Flags

### Caso 10: Secuencia Limpiar → Filtrar → Interactuar
**Pasos detallados**:
1. Filtrar por código "002" → verificar que hay resultados
2. Hacer clic en "Limpiar" (escobita)
3. Verificar que TODO queda vacío (DataGridView1, DataGridView2, cuadro amarillo)
4. Escribir en txtCodigo: "11"
5. Esperar que filtre
6. Hacer clic en la grilla (DataGridView1)
7. Hacer clic en otra fila
8. Navegar con flechas

**Qué validar en cada paso**:
- ✅ Paso 2: Limpia todo sin disparar búsquedas automáticas
- ✅ Paso 3: Todo realmente vacío
- ✅ Paso 5: Filtra correctamente
- ✅ Paso 6-8: NO se borra todo inesperadamente
- ✅ Paso 6-8: Las equivalencias se actualizan normalmente

**Este es el caso que tenía el bug original**

---

### Caso 11: Múltiples limpiezas seguidas
**Pasos**:
1. Limpiar
2. Filtrar algo
3. Limpiar nuevamente
4. Filtrar otra cosa
5. Interactuar con grilla

**Qué validar**:
- ✅ Cada limpieza funciona correctamente
- ✅ No hay efectos acumulativos
- ✅ Siempre se puede filtrar después de limpiar

---

## Casos de Códigos Clickeables

### Caso 12: Click en código de equivalencia
**Preparación**:
1. Buscar "cazoleta" en txtDescripcion
2. Seleccionar un artículo con equivalencias
3. Observar DataGridView2 con códigos en azul

**Pasos**:
1. Hacer clic en cualquier código de la columna "Codigo" en DataGridView2
2. Observar qué sucede

**Qué validar**:
- ✅ txtCodigo se llena con el código clickeado
- ✅ Los demás campos se limpian (txtDescripcion, txtModelo, etc.)
- ✅ Los combos se resetean (índice -1)
- ✅ Se ejecuta un filtro automático
- ✅ DataGridView1 muestra el artículo correspondiente
- ✅ Las equivalencias de ese nuevo artículo se cargan automáticamente

---

### Caso 13: Navegación por equivalencias (cadena)
**Escenario avanzado**:
1. Buscar artículo A (tiene equivalencias)
2. Click en equivalencia B → muestra B y sus equivalencias
3. Click en equivalencia C → muestra C y sus equivalencias
4. Y así sucesivamente

**Qué validar**:
- ✅ Cada click funciona correctamente
- ✅ Siempre muestra las equivalencias del artículo actual
- ✅ No hay confusión de datos entre artículos

---

## Casos de SelectionChanged

### Caso 14: SelectionChanged vs Click
**Objetivo**: Verificar que SelectionChanged carga datos del artículo correcto

**Pasos**:
1. Filtrar para tener al menos 3 artículos
2. Hacer clic en el primer artículo
3. Anotar qué equivalencias muestra
4. Presionar flecha ABAJO (sin hacer clic)
5. Verificar que las equivalencias cambiaron
6. Hacer clic en el artículo actual
7. Verificar que las equivalencias NO cambiaron (son las mismas)

**Qué validar**:
- ✅ Paso 3: Muestra equivalencias del primer artículo
- ✅ Paso 5: Las equivalencias se actualizaron al segundo artículo (por SelectionChanged)
- ✅ Paso 7: Las equivalencias siguen siendo del segundo artículo (no vuelven al primero)

**Este valida que el bug "muestra equivalencias del anterior" está resuelto**

---

### Caso 15: Primera selección automática después de filtrar
**Pasos**:
1. Limpiar todo
2. Filtrar por código "002"
3. SIN hacer clic en nada, observar DataGridView2 y cuadro amarillo

**Qué validar**:
- ✅ El primer artículo de la grilla queda seleccionado automáticamente
- ✅ Sus equivalencias se cargan automáticamente
- ✅ El cuadro amarillo muestra sus datos
- ✅ No necesita click manual

---

## Casos de Edge Cases

### Caso 16: Caracteres especiales en búsqueda
**Búsquedas a probar**:
- "o'ring"
- "kit's"
- "100%"
- "5/8"

**Qué validar**:
- ✅ No produce error SQL
- ✅ Busca correctamente (escapa caracteres especiales)
- ✅ No rompe la aplicación

---

### Caso 17: Búsqueda vacía
**Pasos**:
1. Limpiar todo
2. Hacer clic en txtDescripcion
3. Presionar Enter sin escribir nada

**Qué validar**:
- ✅ No ejecuta búsqueda (o muestra todo)
- ✅ No hay error

---

### Caso 18: Búsqueda muy larga
**Pasos**:
1. Escribir en txtDescripcion: "palabra1 palabra2 palabra3 palabra4 palabra5 palabra6"

**Qué validar**:
- ✅ Funciona correctamente
- ✅ Busca TODAS las palabras (cada una en cualquier campo)
- ✅ No hay límite artificial de palabras

---

## Matriz de Pruebas Cruzadas

### Combinaciones a probar:

| Acción 1 | Acción 2 | Resultado Esperado |
|----------|----------|-------------------|
| Filtrar código | Navegar con flechas | Equivalencias se actualizan |
| Filtrar descripción | Hacer clic en equivalencia | Filtra y muestra nuevo artículo |
| Limpiar | Filtrar código | Funciona normalmente |
| Limpiar | Filtrar descripción | Funciona normalmente |
| Filtrar | Limpiar | Todo se vacía |
| Filtrar | Cambiar de combo | Se aplica filtro combinado |
| Click equivalencia | Navegar con flechas | Funciona desde artículo nuevo |

---

## Registro de Pruebas Rápidas

**Para documentar pruebas rápidas**:

```
Fecha: _______________
Probado por: _______________

[ ] Caso 1: "cazoleta strada" - Resultado: ___________
[ ] Caso 5: Código "11" - Resultado: ___________
[ ] Caso 7: Artículo con equivalencias - Resultado: ___________
[ ] Caso 10: Secuencia Limpiar → Filtrar → Interactuar - Resultado: ___________
[ ] Caso 12: Click en código de equivalencia - Resultado: ___________
[ ] Caso 14: SelectionChanged vs Click - Resultado: ___________

Observaciones:
_________________________________________________
_________________________________________________
_________________________________________________
```

---

## Datos de Prueba Recomendados

**Tener preparados en la base de datos**:

1. **Al menos 3 artículos con equivalencias** (AROBSE lleno)
2. **Al menos 3 artículos sin equivalencias** (AROBSE vacío)
3. **Un artículo con "cazoleta" en descripción y "strada" en modelo**
4. **Códigos que contengan "11" y "002"**
5. **Artículos con palabras compuestas en diferentes campos**

---

## Scripts SQL Útiles para Preparar Datos

```sql
-- Encontrar artículos con equivalencias
SELECT TOP 10 arterc, ardesc, arobse, armode 
FROM listas 
WHERE arobse IS NOT NULL AND arobse <> ''

-- Encontrar artículos sin equivalencias
SELECT TOP 10 arterc, ardesc, armode 
FROM listas 
WHERE arobse IS NULL OR arobse = ''

-- Buscar artículos con "cazoleta"
SELECT arterc, ardesc, armode 
FROM listas 
WHERE ardesc LIKE '%cazoleta%' OR armode LIKE '%cazoleta%'

-- Buscar códigos con "11"
SELECT arterc, ardesc 
FROM listas 
WHERE (arprim+arterc) LIKE '%11%' OR arterc LIKE '%11%'
```

---

## Conclusión

Este documento proporciona casos de prueba específicos y prácticos. Úsalo junto con el TEST_PLAN_EQUIVALENCIAS.md para realizar pruebas completas y documentadas.

**Tiempo estimado por caso**: 2-3 minutos
**Total para todos los casos**: 30-50 minutos

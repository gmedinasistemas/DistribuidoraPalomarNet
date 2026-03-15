# Documentación Técnica: Sistema de Filtros y Búsquedas

**Proyecto:** DistribuidoraPalomarNet  
**Archivos principales:** `Module.cs`, `Forms/frmPrincipal.cs`  
**Última actualización:** Marzo 2026

---

## 1. Mapa de controles de filtro

| Control | Tipo | Filtro SQL | Motor de búsqueda |
|---|---|---|---|
| `cmbArmadora` | ComboBox | `r.rbdesc = '...'` (JOIN RUBROS) | `consultas_Relacionales` / `consultas_Ofertas` |
| `cmbLineas` | ComboBox | `a.armode = '...'` | `Filtrar_x_linea` (exclusivo) |
| `cmbRubros` | ComboBox | `a.armode = '...'` | `consultas_Relacionales` / `consultas_Ofertas` |
| `cmbMarcas` | ComboBox | `a.arprov = {int}` | `consultas_Relacionales` / `consultas_Ofertas` |
| `txtDescripcion` | TextBox | `a.ARDESC LIKE '%...%'` (multi-palabra) | `consultas_Relacionales` / `consultas_Ofertas` |
| `txtModelo` | TextBox | `a.AROBSE LIKE '%...%'` (multi-palabra) | `consultas_Relacionales` / `consultas_Ofertas` |
| `txtCodigo` | TextBox | `a.arprim + a.arterc LIKE '%...%'` | `FiltrarPorCodigo` (exclusivo) |
| `ckbOfertas` | CheckBox | `a.ARTT <> 0` | redirige al motor de ofertas |

> **Nota sobre `txtModelo`:** El campo en pantalla se llama "Modelo" pero filtra sobre `a.AROBSE` (campo observaciones/catálogo del artículo), **no** sobre `a.armode` (que es la línea). No confundir.

---

## 2. Carga de ComboBoxes

### 2.1 `cmbArmadora` — `CargarComboArmadora()`

- **Tabla origen:** `rubros` + JOIN `listas`
- **Filtro SQL:** `RBCODI > 100` (los rubros con código mayor a 100 son Armadoras)
- **SQL completo:**
```sql
SELECT DISTINCT R.rbcodi, R.rbdesc
FROM rubros AS R
INNER JOIN listas AS LT ON R.RBCODI = LT.ARRUB2
WHERE R.RBCODI > 100
ORDER BY R.rbdesc
```
- **Binding:** `Items.Add` — solo carga texto descriptivo, **sin** `SelectedValue`
- **Elemento por defecto:** `"TODOS"` insertado manualmente en índice 0
- **Estado inicial:** `SelectedIndex` no se fuerza a -1 explícitamente
- **Lectura del valor:** `cmbArmadora.Text`
- **⚠️ Bug conocido:** En el bloque `catch`, el mensaje usa `{ex}` sin `$`, por lo que no interpola el error real

---

### 2.2 `cmbLineas` — `CargarComboLineas(string _Armadora)`

- **Tabla origen:** `lineas` + JOIN `listas` (+ JOIN `rubros` si hay armadora activa)
- **Parámetro:** si `_Armadora` es `null`, vacío o `"0"` → carga todas las líneas
- **SQL general (sin filtro):**
```sql
SELECT DISTINCT LI.licodi, LI.lidesc
FROM (lineas AS LI INNER JOIN listas AS LT ON LI.lidesc = LT.ARMODE)
ORDER BY LI.lidesc
```
- **SQL filtrado por Armadora:**
```sql
SELECT DISTINCT LI.licodi, LI.lidesc
FROM (lineas AS LI INNER JOIN listas AS LT ON LI.lidesc = LT.ARMODE)
INNER JOIN RUBROS AS r ON LT.ARRUB2 = r.RBCODI
WHERE r.RBDESC = ?    -- parámetro posicional OleDb (seguro, usa Parameters.AddWithValue)
ORDER BY LI.lidesc
```
- **Binding:** `DataSource` con `DisplayMember = "lidesc"` y `ValueMember = "licodi"`
- **Elemento por defecto:** fila virtual `("0", "TODOS")` agregada antes del bind
- **Estado inicial:** `SelectedIndex = -1` forzado en bloque `finally`
- **Lectura del valor:**
  - `.Text` → devuelve el texto visible (`lidesc`)
  - `.SelectedValue` → devuelve el código (`licodi`, tipo `string`). El valor `"0"` indica "TODOS"
- **Disparador:** cuando el usuario elige una Armadora, se llama `CargarComboLineas(cmbArmadora.Text)` para repoblar con las líneas de esa armadora

> **Mejora pendiente:** El JOIN entre lineas y rubros actualmente usa texto (`r.RBDESC = ?`). Podría usar código: `lineas.LIRUBR = rubros.RBCODI = listas.ARRUB2` para mayor precisión y performance.

---

### 2.3 `cmbRubros` — `CargarComboRubros()`

- **Tabla origen:** `rubros` + JOIN `listas`
- **Filtro SQL:** `RBCODI < 100` (los rubros con código menor a 100 son Repuestos/Partes)
- **SQL completo:**
```sql
SELECT DISTINCT R.rbcodi, R.rbdesc
FROM rubros AS R
INNER JOIN listas AS LT ON R.RBCODI = LT.ARRUB2
WHERE R.RBCODI < 100
ORDER BY R.rbdesc
```
- **Binding:** `Items.Add` — solo carga texto descriptivo, **sin** `SelectedValue`
- **Elemento por defecto:** `"TODOS"` insertado manualmente en índice 0
- **Estado inicial:** `SelectedIndex` no se fuerza a -1 explícitamente
- **Lectura del valor:** `cmbRubros.Text`
- **Campo SQL de filtrado en grilla:** `a.armode = '{cmbRubros.Text}'`
- **⚠️ Bug conocido:** En el bloque `catch`, el mensaje usa `{ex}` sin `$`, por lo que no interpola el error real

---

### 2.4 `cmbMarcas` — `CargarComboMarcas()`

- **Tabla origen:** `proveedo`
- **Filtro SQL:** `prcrub = 'R'` (solo proveedores de tipo Repuesto)
- **SQL completo:**
```sql
SELECT DISTINCT prcodi, prmarc
FROM proveedo
WHERE prcrub = 'R'
ORDER BY prmarc
```
- **Binding:** `DataSource` con `DisplayMember = "prmarc"` y `ValueMember = "prcodi"`
- **Elemento por defecto:** fila virtual `(0, "TODOS")` agregada antes del bind
- **Estado inicial:** `SelectedIndex = -1` forzado explícitamente
- **Lectura del valor:**
  - `.Text` → nombre de marca visible (`prmarc`)
  - `.SelectedValue` → código del proveedor (`prcodi`, tipo `int`). El valor `0` indica "TODOS"
- **Campo SQL de filtrado en grilla:** `a.arprov = {prcodi}` (usa el entero, no texto)

---

### Tabla comparativa de binding

| Control | Binding | `SelectedValue` | Opción "TODOS" | `SelectedIndex = -1` forzado |
|---|---|---|---|---|
| `cmbArmadora` | `Items.Add` (solo texto) | ❌ No disponible | Como string en índice 0 | ❌ No |
| `cmbLineas` | `DataSource` | ✅ `licodi` (string) | Fila `("0","TODOS")` | ✅ Sí (finally) |
| `cmbRubros` | `Items.Add` (solo texto) | ❌ No disponible | Como string en índice 0 | ❌ No |
| `cmbMarcas` | `DataSource` | ✅ `prcodi` (int) | Fila `(0,"TODOS")` | ✅ Sí (explícito) |

---

## 3. Orden de carga al iniciar (`funcion_inicial(opc: 0)`)

```
funcion_inicial(0)
  ├── CargarComboArmadora()
  ├── CargarComboLineas("0")      ← carga todas las líneas
  ├── CargarComboRubros()
  └── CargarComboMarcas()
```

Cuando el usuario selecciona una Armadora:
```
cmbArmadora_SelectedIndexChanged
  └── CargarComboLineas(cmbArmadora.Text)   ← repuebla líneas filtradas
  └── consultas_Relacionales()
```

---

## 4. Motores de búsqueda

### 4.1 `consultas_Relacionales()` — Motor principal (`Module.cs`)

Motor de filtrado multi-criterio. Construye el `WHERE` dinámicamente con una lista de condiciones unidas por `AND`.

**SELECT base:**
```sql
SELECT
    a.arterc       AS Codigo,
    a.ardesc       AS Descripcion,
    p.PRMARC       AS Marca,
    a.armode       AS Linea,
    a.arinte       AS Posición,
    a.arpre2       AS Preciopublico,
    IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) AS aroo,
    a.arff         AS Stock,
    a.artt,
    IIF(a.arubic1 IS NOT NULL, a.arubic1, '') AS arubic1,
    IIF(a.arobse  IS NOT NULL, a.arobse,  '') AS arobse
FROM (listas AS a INNER JOIN proveedo AS p ON p.prcodi = a.arprov)
```

**Construcción del WHERE (condiciones acumuladas por AND):**

| Condición | Siempre | Campo fuente | SQL generado |
|---|---|---|---|
| Filtro de seguridad | ✅ Siempre | — | `a.arprim <> '000'` |
| Descripción | Si `txtDescripcion` no vacío | `a.ARDESC` | `(a.ARDESC LIKE '%p1%' AND a.ARDESC LIKE '%p2%')` |
| Modelo/Catálogo | Si `txtModelo` no vacío | `a.AROBSE` | `(a.AROBSE LIKE '%p1%' AND a.AROBSE LIKE '%p2%')` |
| Armadora | Si `cmbArmadora.Text` no vacío | `r.rbdesc` + JOIN RUBROS | `r.rbdesc = '<texto>'` |
| Marca | Si `cmbMarcas.SelectedValue` ≠ 0 | `a.arprov` | `a.arprov = {prcodi int}` |
| Rubro | Si `cmbRubros.Text` no vacío/TODOS | `a.armode` | `a.armode = '<texto>'` |

**JOIN adicional** (solo si Armadora activa):
```sql
INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI
```

**Ejemplo — Armadora + Marca + Descripción multi-palabra:**
```sql
FROM (listas AS a INNER JOIN proveedo AS p ON p.prcodi = a.arprov)
INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI
WHERE a.arprim <> '000'
  AND (a.ARDESC LIKE '%filtro%' AND a.ARDESC LIKE '%aceite%')
  AND r.rbdesc = 'FORD'
  AND a.arprov = 42
ORDER BY a.arterc
```

---

### 4.2 `consultas_Ofertas()` — Motor de ofertas (`Module.cs`)

Idéntico a `consultas_Relacionales` en estructura, con dos diferencias:

1. **Condición base adicional:** `a.ARTT <> 0` (solo artículos con precio de oferta)
2. **Campo precio:** usa `a.arpre5` (precio de oferta) en lugar de `a.arpre2`

```sql
-- Condiciones base siempre presentes:
a.ARTT <> 0
a.arprim <> '000'
-- + las mismas condiciones dinámicas de consultas_Relacionales
```

---

### 4.3 `Filtrar_x_linea(string _linea)` — Filtro exclusivo de `cmbLineas` (`frmPrincipal.cs`)

Motor independiente. **No lee** ningún otro filtro activo (`cmbArmadora`, `txtDescripcion`, etc.).

```sql
SELECT
    a.arterc       AS Codigo,
    a.ardesc       AS Descripcion,
    p.prrazo       AS Marca,        -- ← usa prrazo (razón social), NO p.PRMARC
    a.armode       AS Linea,
    a.arinte       AS Posición,
    IIF(a.artt IS NOT NULL AND a.artt <> 0, a.arpre5, a.arpre2) AS Preciopublico,
    IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) AS aroo,
    a.arff         AS Stock, a.artt, arubic1, arobse
FROM listas AS a
INNER JOIN proveedo AS p ON p.prcodi = a.arprov
WHERE a.armode = '<cmbLineas.Text>'   -- sin filtro arprim <> '000'
ORDER BY a.arcodi                      -- ordena por arcodi, no arterc
```

**Diferencias clave vs. `consultas_Relacionales`:**

| Aspecto | `Filtrar_x_linea` | `consultas_Relacionales` |
|---|---|---|
| Campo Marca | `p.prrazo` (razón social) | `p.PRMARC` |
| Filtro seguridad `arprim<>'000'` | ❌ No incluido | ✅ Siempre |
| Orden resultado | `a.arcodi` | `a.arterc` |
| Combina otros filtros | ❌ Ninguno | ✅ Todos |
| Precio | `IIF(arpre5/arpre2)` | `a.arpre2` fijo |

---

### 4.4 `FiltrarPorCodigo()` — Filtro exclusivo de `txtCodigo` (`frmPrincipal.cs`)

Motor independiente. Disparado por tecla `Enter` en `txtCodigo`. **No participa** en `consultas_Relacionales`.

```sql
-- Sin ckbOfertas:
WHERE a.arprim + a.arterc LIKE '%<txtCodigo.Text>%'
ORDER BY a.arcodi

-- Con ckbOfertas activo:
WHERE a.ARTT <> 0 AND a.arprim + a.arterc LIKE '%<txtCodigo.Text>%'
ORDER BY a.arcodi
```

- Usa `p.PRMARC` para Marca (igual que `consultas_Relacionales`)
- Precio: `IIF(artt, arpre5, arpre2)` adaptativo según oferta
- **Búsqueda aislada:** si hay texto en `txtCodigo` y además hay `cmbArmadora` activa, `FiltrarPorCodigo` ignora la armadora. Solo busca por código.

---

## 5. Reactividad y Debounce

### 5.1 Filtros con delay (timer2 — 1400 ms)

`txtDescripcion` y `txtModelo` usan patrón debounce manual. Cada pulsación:
1. `timer2.Stop()` — cancela la espera anterior
2. `timer2.Start()` — inicia nueva cuenta de 1400 ms

Cuando el temporizador expira (`timer2_Tick`):

| Condición de guardia | Acción |
|---|---|
| `_porLimpieza == true` | Aborta (evita filtrado durante limpieza manual) |
| Ambos campos vacíos | Aborta |
| `ckbOfertas == false` | `Module.consultas_Relacionales()` |
| `ckbOfertas == true` | `Module.consultas_Ofertas()` |

### 5.2 Filtros instantáneos (sin delay)

| Control | Evento | Motor invocado |
|---|---|---|
| `txtCodigo` | `KeyPress` (Enter) | `FiltrarPorCodigo()` |
| `txtEquivalencia` | `KeyPress` (Enter) | `consultas_Relacionales()` directo |
| `cmbArmadora` | `SelectedIndexChanged` | `consultas_Relacionales()` + `CargarComboLineas()` |
| `cmbLineas` | `SelectedIndexChanged` | `Filtrar_x_linea()` |
| `cmbRubros` | `SelectedIndexChanged` | `consultas_Relacionales()` / `consultas_Ofertas()` |
| `cmbMarcas` | `SelectedIndexChanged` | `consultas_Relacionales()` / `consultas_Ofertas()` |
| `ckbOfertas` | `Click` | `consultas_Ofertas()` / `consultas_Relacionales()` |

---

## 6. `RefrescarSegunFiltros()` — Re-aplicación de filtros activos (`frmPrincipal.cs`)

Método llamado al cerrar `frmConfFamilia` o `frmDatosUsuario` vía `btnSalir`, para reaplicar el filtro que estaba activo en lugar de hacer siempre una recarga total.

**Lógica de precedencia (primer true gana):**

```
1. cmbLineas.SelectedValue != 0          → Filtrar_x_linea(cmbLineas.Text)
2. txtCodigo no vacío                    → FiltrarPorCodigo()
3. ckbOfertas.Checked                   → consultas_Ofertas()
4. txtDescripcion/txtModelo no vacíos   → consultas_Relacionales()
   O cmbArmadora/cmbRubros/cmbMarcas activos (≠ vacío/TODOS)
5. Sin filtros activos                  → funcion_inicial(1)   (recarga total)
```

---

## 7. Algoritmo multi-palabra — `GenerarCondicionMultipalabra()`

Método helper en `Module.cs`. Recibe un campo SQL y el texto del usuario.

- Divide el texto por espacios
- Por cada palabra genera: `campo LIKE '%palabra%'`
- Une todas con `AND`

**Ejemplo:** input `"filtro aceite ford"` sobre campo `a.ARDESC`:
```sql
(a.ARDESC LIKE '%filtro%' AND a.ARDESC LIKE '%aceite%' AND a.ARDESC LIKE '%ford%')
```

Ventaja: encuentra resultados sin importar el orden de las palabras en la base.

---

## 8. Relaciones entre tablas (referencia rápida)

```
listas (artículos)
  ├── a.arprov   = proveedo.prcodi    (Marca / Proveedor)
  ├── a.ARRUB2   = rubros.RBCODI      (Rubro o Armadora)
  ├── a.armode                        (Línea — coincide con lineas.lidesc)
  └── a.arprim   = '000'              (registros de sistema, siempre excluidos)

rubros
  ├── RBCODI > 100  → Armadoras       (cmbArmadora)
  └── RBCODI < 100  → Rubros/Partes   (cmbRubros)

lineas
  ├── licodi  → ValueMember de cmbLineas
  ├── lidesc  → DisplayMember (= a.armode en listas)
  └── LIRUBR  → podría usarse para JOIN con rubros (mejora pendiente)
```

---

## 9. Columnas del SELECT y campos de precio

| Alias resultado | Campo origen | Notas |
|---|---|---|
| `Codigo` | `a.arterc` | Código de artículo visible |
| `Descripcion` | `a.ardesc` | |
| `Marca` | `p.PRMARC` | En `Filtrar_x_linea` usa `p.prrazo` (razón social) |
| `Linea` | `a.armode` | |
| `Posición` | `a.arinte` | |
| `Preciopublico` | `a.arpre2` (fijo) en `consultas_Relacionales` | `a.arpre5` en ofertas; `IIF(arpre5,arpre2)` en `FiltrarPorCodigo` y `Filtrar_x_linea` |
| `aroo` | `IIF(a.artt<>0, a.artt, a.aroo)` | |
| `Stock` | `a.arff` | |

---

## 10. Comportamientos importantes para QA

### 10.1 `txtCodigo` es búsqueda aislada
No se combina con ningún otro filtro activo. Si hay `cmbArmadora = "FORD"` y se busca por código `11758`, solo filtra por código. Para combinar, primero limpiar el código o viceversa.

### 10.2 `cmbLineas` es excluyente
`Filtrar_x_linea` ignora `cmbArmadora`, `cmbRubros`, `cmbMarcas`, `txtDescripcion` y `txtModelo`. Es una búsqueda directa por `a.armode`.

### 10.3 Precio en `consultas_Relacionales` es fijo (`arpre2`)
A diferencia de `FiltrarPorCodigo` y `Filtrar_x_linea`, en `consultas_Relacionales` el precio no usa IIF; siempre muestra `arpre2`. En `consultas_Ofertas` siempre muestra `arpre5`.

### 10.4 SQL Injection en `cmbArmadora` y `cmbRubros`
Usan binding `Items.Add` (sin parámetros), por lo que el texto se inserta directamente en el SQL: `r.rbdesc = '{cmbArmadora.Text}'`. Si el nombre del rubro contiene comillas simples, puede causar error. `cmbLineas` ya usa `Parameters.AddWithValue` (correcto).

### 10.5 Bug de interpolación en `catch` de `CargarComboArmadora` y `CargarComboRubros`
```csharp
// ⚠️ Sin el $, {ex} es texto literal, no el mensaje de error:
throw new Exception(", dentro de 'CargarComboArmadora': {ex}");
// Correcto sería:
throw new Exception($", dentro de 'CargarComboArmadora': {ex.Message}");
```

---

## 11. Diagrama de decisión al seleccionar un filtro

```
Usuario interactúa con filtro
          │
          ├── cmbArmadora cambia?
          │     ├── "TODOS" → cmbArmadora.Text = "" → funcion_inicial(0)
          │     └── otro → CargarComboLineas(texto) + consultas_Relacionales()
          │
          ├── cmbLineas cambia?
          │     ├── SelectedValue == 0 ("TODOS") → no hace nada
          │     └── otro → Filtrar_x_linea(cmbLineas.Text)
          │
          ├── cmbRubros cambia?
          │     ├── "TODOS" o vacío → no hace nada
          │     └── otro → consultas_Relacionales() / consultas_Ofertas()
          │
          ├── cmbMarcas cambia?
          │     ├── SelectedValue == 0 ("TODOS") → no hace nada
          │     └── otro → consultas_Relacionales() / consultas_Ofertas()
          │
          ├── txtDescripcion / txtModelo escribe?
          │     └── timer2.Stop + timer2.Start (debounce 1400ms)
          │           └── timer2_Tick → consultas_Relacionales() / consultas_Ofertas()
          │
          └── txtCodigo presiona Enter?
                └── FiltrarPorCodigo()
```

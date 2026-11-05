# Guía de Pruebas - Sistema de Equivalencias

## 📋 Archivos de Prueba Disponibles

Este directorio contiene 4 archivos relacionados con pruebas:

### 1. **TEST_PLAN_EQUIVALENCIAS.md** ⭐
**Qué es**: Plan de pruebas completo y detallado con 9 categorías de pruebas.

**Cuándo usarlo**: 
- Para pruebas formales antes de release
- Para validación completa del sistema
- Para documentar resultados oficiales

**Contenido**:
- 20+ tests individuales
- Pasos detallados para cada prueba
- Resultados esperados y criterios de fallo
- Checklist final de validación

---

### 2. **CASOS_PRUEBA_EJEMPLOS.md** 🎯
**Qué es**: Casos de prueba específicos con datos reales y ejemplos concretos.

**Cuándo usarlo**:
- Cuando necesitas ejemplos prácticos
- Para entender QUÉ buscar exactamente
- Para pruebas rápidas de validación

**Contenido**:
- 18 casos específicos con ejemplos
- Búsquedas recomendadas ("cazoleta strada", "11", etc.)
- Matriz de pruebas cruzadas
- Scripts SQL para preparar datos

---

### 3. **test-equivalencias.ps1** 🤖
**Qué es**: Script de PowerShell interactivo para asistir en las pruebas.

**Cuándo usarlo**:
- Para abrir rápidamente los planes de prueba
- Para generar plantillas de reportes
- Para verificar que los archivos críticos están actualizados

**Cómo usarlo**:
```powershell
# Menú interactivo
.\test-equivalencias.ps1

# Abrir directamente el plan de pruebas
.\test-equivalencias.ps1 -AbrirPlan

# Generar plantilla de reporte
.\test-equivalencias.ps1 -GenerarReporte
```

**Funciones**:
1. Abrir plan de pruebas en Notepad
2. Generar plantilla de reporte con fecha
3. Ver checklist rápido en consola
4. Verificar archivos críticos modificados

---

### 4. **README_TESTS.md** 📖
**Qué es**: Este archivo. Guía de inicio rápido.

---

## 🚀 Inicio Rápido - 3 Niveles de Pruebas

### Nivel 1: Prueba Rápida (5-10 minutos)
**Para**: Validación rápida después de cambios menores

**Pasos**:
1. Ejecutar `.\test-equivalencias.ps1`
2. Seleccionar opción 3: "Ver Checklist Rápido"
3. Probar manualmente cada item del checklist
4. Anotar cualquier problema

**Checklist Rápido incluye**:
- ✅ Equivalencias cargan correctamente
- ✅ "cazoleta strada" funciona
- ✅ Códigos clickeables funcionan
- ✅ Limpiar no causa problemas

---

### Nivel 2: Prueba Selectiva (15-20 minutos)
**Para**: Validación de funcionalidades específicas

**Pasos**:
1. Abrir `CASOS_PRUEBA_EJEMPLOS.md`
2. Seleccionar 5-7 casos críticos:
   - Caso 2: "amortiguador delantero fiat"
   - Caso 5: Código "11"
   - Caso 7: Artículo con equivalencias
   - Caso 10: Secuencia Limpiar → Filtrar → Interactuar
   - Caso 14: SelectionChanged vs Click
3. Ejecutar cada caso y anotar resultados

---

### Nivel 3: Prueba Completa (45-60 minutos)
**Para**: Validación formal antes de release o commit importante

**Pasos**:
1. Ejecutar `.\test-equivalencias.ps1 -GenerarReporte`
2. Abrir `TEST_PLAN_EQUIVALENCIAS.md`
3. Ejecutar TODOS los tests (secciones 1-8)
4. Llenar el reporte generado
5. Firmar y guardar el reporte

---

## 🎯 Qué Probar Según el Cambio Realizado

### Si cambios en: `Module.cs` → búsqueda multi-campo
**Probar**:
- Test 3.1, 3.2, 3.3, 3.4 (Búsqueda Multi-Campo)
- Caso 1, 2, 3, 4 (ejemplos específicos)

### Si cambios en: `frmPrincipal.cs` → SelectionChanged
**Probar**:
- Test 1.1, 1.2, 1.3 (Equivalencias - Carga Automática)
- Caso 7, 9, 14, 15 (SelectionChanged específicos)

### Si cambios en: Flags de limpieza
**Probar**:
- Test 5.1, 5.2, 5.3 (Control de Flags)
- Caso 10, 11 (secuencias de limpieza)

### Si cambios en: Códigos clickeables
**Probar**:
- Test 4.1, 4.2 (Códigos Clickeables)
- Caso 12, 13 (navegación por equivalencias)

---

## 📊 Cómo Reportar un Bug

Si encuentras un problema:

1. **Documenta el contexto**:
   - ¿Qué test estabas ejecutando?
   - ¿Qué pasos seguiste exactamente?

2. **Captura el resultado**:
   - ¿Qué esperabas que pasara?
   - ¿Qué pasó realmente?
   - ¿Hay mensaje de error? (captura de pantalla)

3. **Intenta reproducirlo**:
   - ¿Pasa siempre o solo a veces?
   - ¿Con qué datos específicos pasa?

4. **Usa el formato**:
```
Bug #: [número]
Test: [nombre del test]
Descripción: [qué salió mal]
Pasos para reproducir:
1. 
2. 
3. 
Resultado esperado: 
Resultado actual: 
Severidad: [Crítica/Alta/Media/Baja]
Reproducible: [Siempre/A veces/Una vez]
```

---

## ✅ Criterios de Aceptación

Para considerar que las pruebas PASARON:

### Criterios Mínimos (DEBE cumplir todos):
- ✅ La búsqueda "cazoleta strada" encuentra artículos con palabras en campos diferentes
- ✅ Las equivalencias se cargan al hacer clic Y al navegar con teclado
- ✅ Los códigos en equivalencias son clickeables y filtran correctamente
- ✅ Después de limpiar → filtrar → hacer clic en grilla NO se borra todo
- ✅ El código "11" encuentra resultados

### Criterios Deseables (importante pero no bloqueante):
- ✅ Todos los edge cases funcionan sin errores
- ✅ Tooltips y estilos visuales correctos
- ✅ Performance aceptable con cambios rápidos

### Criterios Críticos (un fallo = RECHAZAR):
- ❌ Las equivalencias muestran datos del artículo anterior (desfase)
- ❌ Se produce excepción/crash en uso normal
- ❌ La búsqueda multi-campo no funciona para nada
- ❌ Limpiar rompe la funcionalidad de filtrado

---

## 🔍 Verificación Rápida de Archivos Modificados

Ejecuta esto para verificar que los cambios están presentes:

```powershell
.\test-equivalencias.ps1
# Selecciona opción 4: "Verificar Archivos Críticos"
```

Esto verifica:
- ✅ Archivos críticos existen
- ✅ Fecha de última modificación
- ✅ Métodos clave presentes en código
- ✅ Cambios específicos aplicados

---

## 📝 Plantillas de Reportes

### Reporte Rápido (Nivel 1 y 2):
```
REPORTE DE PRUEBAS RÁPIDAS
Fecha: _______________
Probado por: _______________

RESULTADOS:
[ ] ✅ APROBADO - Todo funciona correctamente
[ ] ⚠️ APROBADO CON OBSERVACIONES - Ver notas
[ ] ❌ RECHAZADO - Problemas críticos encontrados

PRUEBAS EJECUTADAS:
1. _________________ : [ ] PASS [ ] FAIL
2. _________________ : [ ] PASS [ ] FAIL
3. _________________ : [ ] PASS [ ] FAIL

OBSERVACIONES:
_________________________________
_________________________________

BUGS ENCONTRADOS:
_________________________________
_________________________________
```

### Reporte Formal (Nivel 3):
Generado automáticamente por `.\test-equivalencias.ps1 -GenerarReporte`

---

## 🎓 Tips para Pruebas Efectivas

1. **Prepara datos de prueba**: Antes de empezar, identifica artículos específicos que usarás
2. **Prueba una cosa a la vez**: No combines múltiples cambios
3. **Documenta mientras pruebas**: No confíes en tu memoria
4. **Reproduce bugs**: Si algo falla, intenta hacerlo fallar de nuevo
5. **Usa el teclado**: No solo el mouse (prueba navegación con flechas)
6. **Prueba secuencias**: No solo acciones individuales
7. **Casos extremos importan**: Grilla vacía, sin equivalencias, caracteres especiales

---

## 📞 Contacto y Soporte

Si tienes dudas sobre las pruebas:
- Consulta el `TEST_PLAN_EQUIVALENCIAS.md` para detalles técnicos
- Revisa `CASOS_PRUEBA_EJEMPLOS.md` para ejemplos prácticos
- Usa el script PowerShell para generar reportes

---

## 🔄 Historial de Versiones de Tests

**v1.0** - 2026-08-22
- Plan de pruebas inicial
- Casos de ejemplo
- Script de asistencia
- Documentación completa

---

## 📚 Resumen de Commits Relacionados

**Commit anterior**: "fix: Mejorar búsqueda multi-campo y auto-actualización de equivalencias"

**Cambios clave probados**:
1. Búsqueda multi-palabra entre campos diferentes (txtDescripcion)
2. SelectionChanged actualiza equivalencias automáticamente
3. Flag PorLimpieza evita búsquedas durante limpieza
4. Module.NumEqui se actualiza antes de cargar equivalencias
5. Códigos clickeables en grilla de equivalencias

---

## 🎯 Próximos Pasos

1. Ejecutar al menos Nivel 1 (Prueba Rápida)
2. Si todo OK → commit con confianza
3. Si hay problemas → documentar en reporte
4. Para release formal → ejecutar Nivel 3 (Prueba Completa)

**¡Buenas pruebas!** 🚀

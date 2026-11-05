# Script de Ayuda para Pruebas de Equivalencias
# Este script NO ejecuta pruebas automáticas, sino que ayuda a documentar los resultados

param(
    [switch]$GenerarReporte,
    [switch]$AbrirPlan
)

$testPlanPath = ".\TEST_PLAN_EQUIVALENCIAS.md"
$reportPath = ".\TEST_REPORT_EQUIVALENCIAS_$(Get-Date -Format 'yyyyMMdd_HHmmss').md"

function Show-Menu {
    Clear-Host
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  ASISTENTE DE PRUEBAS - SISTEMA EQUIVALENCIAS" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Abrir Plan de Pruebas (TEST_PLAN_EQUIVALENCIAS.md)" -ForegroundColor Yellow
    Write-Host "2. Generar Plantilla de Reporte" -ForegroundColor Yellow
    Write-Host "3. Ver Checklist Rápido" -ForegroundColor Yellow
    Write-Host "4. Verificar Archivos Críticos" -ForegroundColor Yellow
    Write-Host "5. Salir" -ForegroundColor Yellow
    Write-Host ""
}

function Open-TestPlan {
    if (Test-Path $testPlanPath) {
        Write-Host "Abriendo plan de pruebas..." -ForegroundColor Green
        Start-Process notepad.exe $testPlanPath
    } else {
        Write-Host "ERROR: No se encuentra el archivo $testPlanPath" -ForegroundColor Red
    }
}

function New-TestReport {
    $content = @"
# Reporte de Pruebas - Sistema de Equivalencias
**Fecha de prueba**: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
**Ejecutado por**: [NOMBRE]
**Versión probada**: [COMMIT HASH o VERSIÓN]

---

## Resumen Ejecutivo

- **Total de pruebas**: 
- **Pruebas exitosas**: 
- **Pruebas fallidas**: 
- **Pruebas omitidas**: 
- **Estado general**: [APROBADO / APROBADO CON OBSERVACIONES / RECHAZADO]

---

## Resultados por Categoría

### 1. Equivalencias - Carga Automática
- Test 1.1 - Selección por Click: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 1.2 - Selección por Teclado: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 1.3 - Equivalencias después de Filtrar: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 

### 2. Limpieza - Estado Vacío
- Test 2.1 - Limpieza con btnLimpiar: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 2.2 - Filtro sin resultados: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 

### 3. Búsqueda Multi-Campo
- Test 3.1 - Búsqueda de una palabra: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 3.2 - Búsqueda multi-palabra en campos diferentes: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Búsqueda probada: "cazoleta strada"
  - Resultados encontrados: [SÍ/NO]
  - Notas: 
  
- Test 3.3 - Búsqueda 3 palabras: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 3.4 - Palabras parciales: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 

### 4. Códigos Clickeables
- Test 4.1 - Click en código: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 4.2 - Tooltip: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 

### 5. Control de Flags
- Test 5.1 - Limpieza no dispara búsquedas: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 5.2 - Escribir después de limpiar: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 5.3 - Interactuar después de limpiar: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 

### 6. Búsqueda por Código
- Test 6.1 - Código completo: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 6.2 - Código parcial "11": [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 

### 7. Flujo Integrado
- Test 7.1 - Flujo completo: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas detalladas por paso:
    - Paso 2 (búsqueda): 
    - Paso 4 (equivalencias): 
    - Paso 5 (navegación): 
    - Paso 6 (click equivalencia): 
    - Paso 8 (limpiar): 
    - Paso 10 (interacción): 

### 8. Edge Cases
- Test 8.1 - Artículo sin equivalencias: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 8.2 - Cambio rápido: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 
  
- Test 8.3 - Caracteres especiales: [ ] ✅ PASS  [ ] ❌ FAIL  [ ] ⏭️ SKIP
  - Notas: 

---

## Bugs Encontrados

### Bug #1
- **Test**: 
- **Descripción**: 
- **Pasos para reproducir**:
  1. 
  2. 
  3. 
- **Resultado esperado**: 
- **Resultado actual**: 
- **Severidad**: [Crítica/Alta/Media/Baja]

### Bug #2
[Si aplica]

---

## Observaciones Generales

[Comentarios adicionales, sugerencias, problemas no críticos, etc.]

---

## Conclusión y Recomendación

[Aprobar / Aprobar con observaciones / Rechazar]

**Razones**:
- 
- 
- 

---

## Firma

- **Probado por**: ________________
- **Revisado por**: ________________
- **Fecha**: $(Get-Date -Format 'yyyy-MM-dd')

"@

    $content | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "Plantilla de reporte generada: $reportPath" -ForegroundColor Green
    Start-Process notepad.exe $reportPath
}

function Show-QuickChecklist {
    Clear-Host
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  CHECKLIST RÁPIDO - Funcionalidades Críticas" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "EQUIVALENCIAS:" -ForegroundColor Yellow
    Write-Host "  [ ] Las equivalencias cargan al hacer clic en artículo"
    Write-Host "  [ ] Las equivalencias cargan al navegar con teclado"
    Write-Host "  [ ] Se limpian al presionar botón Limpiar"
    Write-Host "  [ ] Muestran datos del artículo CORRECTO (no el anterior)"
    Write-Host ""
    Write-Host "BÚSQUEDA MULTI-CAMPO:" -ForegroundColor Yellow
    Write-Host "  [ ] 'cazoleta strada' encuentra artículos con palabras en campos diferentes"
    Write-Host "  [ ] Busca en Descripción, Montadora y Línea simultáneamente"
    Write-Host "  [ ] Funciona con palabras parciales"
    Write-Host ""
    Write-Host "CÓDIGOS CLICKEABLES:" -ForegroundColor Yellow
    Write-Host "  [ ] Los códigos en grilla de equivalencias están en azul"
    Write-Host "  [ ] Tooltip 'VER' aparece al pasar mouse"
    Write-Host "  [ ] Al hacer clic, filtra por ese código"
    Write-Host ""
    Write-Host "LIMPIEZA:" -ForegroundColor Yellow
    Write-Host "  [ ] Después de limpiar, se puede filtrar normalmente"
    Write-Host "  [ ] NO se borra todo al hacer clic en grilla después de filtrar"
    Write-Host "  [ ] Limpiar no dispara búsquedas automáticas"
    Write-Host ""
    Write-Host "BÚSQUEDA POR CÓDIGO:" -ForegroundColor Yellow
    Write-Host "  [ ] '11' encuentra resultados"
    Write-Host "  [ ] '002' encuentra resultados"
    Write-Host ""
    Write-Host "Presiona cualquier tecla para volver al menú..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Test-CriticalFiles {
    Clear-Host
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  VERIFICACIÓN DE ARCHIVOS CRÍTICOS" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $files = @(
        @{Path=".\DistribuidoraPalomarNet\Forms\frmPrincipal.cs"; Name="frmPrincipal.cs"},
        @{Path=".\DistribuidoraPalomarNet\Forms\frmPrincipal.Designer.cs"; Name="frmPrincipal.Designer.cs"},
        @{Path=".\DistribuidoraPalomarNet\Module.cs"; Name="Module.cs"},
        @{Path=".\DistribuidoraPalomarNet\Forms\frmConfFamilia.cs"; Name="frmConfFamilia.cs"}
    )
    
    foreach ($file in $files) {
        if (Test-Path $file.Path) {
            Write-Host "[✓]" -ForegroundColor Green -NoNewline
            Write-Host " $($file.Name) existe"
            
            $lastWrite = (Get-Item $file.Path).LastWriteTime
            Write-Host "    Última modificación: $lastWrite" -ForegroundColor Gray
        } else {
            Write-Host "[✗]" -ForegroundColor Red -NoNewline
            Write-Host " $($file.Name) NO ENCONTRADO"
        }
    }
    
    Write-Host ""
    Write-Host "Verificando métodos clave en frmPrincipal.cs..." -ForegroundColor Yellow
    
    $frmPrincipalPath = ".\DistribuidoraPalomarNet\Forms\frmPrincipal.cs"
    if (Test-Path $frmPrincipalPath) {
        $content = Get-Content $frmPrincipalPath -Raw
        
        $methods = @(
            "DataGridView1_SelectionChanged",
            "CargarEquivalencias",
            "PorLimpieza",
            "DataGridView2_Click",
            "EnsureCodigoColumnLinkStyling"
        )
        
        foreach ($method in $methods) {
            if ($content -match $method) {
                Write-Host "  [✓] $method" -ForegroundColor Green
            } else {
                Write-Host "  [✗] $method NO ENCONTRADO" -ForegroundColor Red
            }
        }
    }
    
    Write-Host ""
    Write-Host "Verificando cambios en Module.cs..." -ForegroundColor Yellow
    
    $modulePath = ".\DistribuidoraPalomarNet\Module.cs"
    if (Test-Path $modulePath) {
        $content = Get-Content $modulePath -Raw
        
        if ($content -match "NUEVA LÓGICA: Cada palabra puede estar en cualquier campo") {
            Write-Host "  [✓] Búsqueda multi-campo actualizada" -ForegroundColor Green
        } else {
            Write-Host "  [✗] Búsqueda multi-campo NO actualizada" -ForegroundColor Red
        }
        
        if ($content -match "PorLimpieza") {
            Write-Host "  [✓] Control de limpieza implementado" -ForegroundColor Green
        } else {
            Write-Host "  [✗] Control de limpieza NO implementado" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "Presiona cualquier tecla para volver al menú..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Main Script Logic
if ($AbrirPlan) {
    Open-TestPlan
    exit
}

if ($GenerarReporte) {
    New-TestReport
    exit
}

# Interactive Menu
do {
    Show-Menu
    $choice = Read-Host "Selecciona una opción (1-5)"
    
    switch ($choice) {
        "1" { Open-TestPlan }
        "2" { New-TestReport }
        "3" { Show-QuickChecklist }
        "4" { Test-CriticalFiles }
        "5" { 
            Write-Host "¡Hasta luego!" -ForegroundColor Green
            exit 
        }
        default { 
            Write-Host "Opción inválida. Presiona Enter para continuar..." -ForegroundColor Red
            Read-Host
        }
    }
} while ($true)

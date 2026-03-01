# Guía: Generar tu primer instalador (MSI) con WiX v7 (HeatWave / Visual Studio)

Tenés 2 caminos.

- El más simple para “primer instalador” es con **HeatWave (Visual Studio)**.
- El otro es por **CLI**.

---

## Camino A) Visual Studio + HeatWave (más fácil)

### 1) Instalá HeatWave

En Visual Studio:

- Extensions -> Manage Extensions
- Buscá **HeatWave (FireGiant)** e instalalo.

### 2) Crear el proyecto de instalador

En la solución:

- Add -> New Project
- Elegí un template tipo:
  - **WiX MSI Package (WiX v4/v5/v6/v7)** (según te aparezca con HeatWave)
- Nombralo por ejemplo:
  - `DistribuidoraPalomarNet.Installer`

### 3) Configurar para tomar la salida de tu app

La idea es que el instalador copie lo que realmente corrés:

- `DistribuidoraPalomarNet.exe`
- DLLs del `bin\Release\net8.0-windows\win-x86\`
- Archivos que tu app necesita:
  - `Config.ini` (si tiene que ir a `C:\DistribuidoraPalomarNet`)
  - Base `.mdb` si aplica (según tu flujo)
  - Imágenes / recursos que uses “por ruta fija”

En HeatWave normalmente se hace con:

- Project Reference al `.csproj` principal, o
- Harvest del output folder (carpeta `bin\Release...`)

### 4) Definir carpeta de instalación

Elegí una:

- **Program Files** (recomendado para apps instalables):
  - `ProgramFilesFolder\DistribuidoraPalomarNet\`

Si tu app realmente necesita escribir en `C:\DistribuidoraPalomarNet` (por la base y descargas), entonces:

- Instalás el EXE en Program Files
- Y creás además `C:\DistribuidoraPalomarNet` para datos
  - (Mejor práctica: `ProgramDataFolder`, pero entiendo que hoy usás rutas fijas)

### 5) Compilar el MSI

- Poné la solución en **Release**
- Build del proyecto instalador
- Te genera un `.msi` en `bin\Release\...` del proyecto WiX

---

## Camino B) CLI (rápido si ya conocés WiX)

### 1) Instalar WiX (si no lo tenés)

`wix` como herramienta (depende de tu setup). Si ya te aparece el error `WIX7015`, probablemente ya está entrando por NuGet/SDK, pero igual para CLI suele ser:

- `dotnet tool install --global wix` (si aplica en tu entorno)
- o usar el SDK en un `.wixproj`

### 2) Crear `.wxs` + build

- Compilar tu app Release
- Harvest / declarar archivos
- `wix build ... -o setup.msi`

---

## Pregunta clave (para hacerlo bien sin romperte nada)

Antes de armar el instalador, necesitás definir **cómo querés instalar la carpeta de datos**:

### Modo “prolijo” (recomendado)

- EXE en: `C:\Program Files\DistribuidoraPalomarNet\`
- Datos en: `C:\ProgramData\DistribuidoraPalomarNet\` (o similar)

### Modo “como hoy” (rutas fijas)

- Crear/usar `C:\DistribuidoraPalomarNet` como carpeta principal (EXE + base + descargas)

Confirmar además si el instalador debe incluir:

- `BASEPALOMARNET.mdb` (sí/no)
- `Config.ini` inicial (sí/no)
- Crear acceso directo en escritorio (sí/no)

Con eso, se arma un proyecto WiX mínimo (MSI) y quedan generando instaladores “de verdad”.

1) Compilar la aplicación (Release x86)
powershell
dotnet build .\DistribuidoraPalomarNet\DistribuidoraPalomarNet.csproj -c Release -p:Platform=x86 -v minimal

2) Compilar el instalador WiX (Release)
powershell
dotnet build .\DistribuidoraPalomarNet.Installer\DistribuidoraPalomarNet.Installer.wixproj -c Release -v minimal

(Opcional) En una sola línea (secuencial)
powershell
dotnet build .\DistribuidoraPalomarNet\DistribuidoraPalomarNet.csproj -c Release -p:Platform=x86 -v minimal; `
dotnet build .\DistribuidoraPalomarNet.Installer\DistribuidoraPalomarNet.Installer.wixp
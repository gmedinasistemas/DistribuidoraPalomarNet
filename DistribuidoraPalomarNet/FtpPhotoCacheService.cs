using FluentFTP;
using System;
using System.IO;
using System.Threading.Tasks;

namespace DistribuidoraPalomarNet
{
    /// <summary>
    /// Servicio de caché local para fotos de artículos descargadas desde FTP.
    /// Implementa estrategia de descarga bajo demanda: solo descarga la foto cuando se solicita
    /// y la almacena en caché local para accesos futuros instantáneos.
    /// </summary>
    public class FtpPhotoCacheService
    {
        private readonly string _ftpHost;
        private readonly int _ftpPort;
        private readonly string _ftpUser;
        private readonly string _ftpPassword;
        private readonly string _cacheFolder;
        private readonly string _ftpRemoteFolder;

        /// <summary>
        /// Inicializa el servicio de caché de fotos FTP.
        /// </summary>
        /// <param name="host">Host del servidor FTP (ej: c1480632.ferozo.com)</param>
        /// <param name="port">Puerto FTP (generalmente 21)</param>
        /// <param name="user">Usuario FTP</param>
        /// <param name="password">Contraseña FTP</param>
        /// <param name="remoteFolder">Carpeta remota donde están las fotos (ej: /imagenes/FOTOS)</param>
        public FtpPhotoCacheService(string host, int port, string user, string password, string remoteFolder = "/imagenes/FOTOS")
        {
            _ftpHost = host;
            _ftpPort = port;
            _ftpUser = user;
            _ftpPassword = password;
            _ftpRemoteFolder = remoteFolder;

            // Carpeta de caché en %LocalAppData%\DistribuidoraPalomarNet\FotosCache
            _cacheFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DistribuidoraPalomarNet", "FotosCache");
            
            Directory.CreateDirectory(_cacheFolder);
        }

        /// <summary>
        /// Obtiene la ruta local de la foto del artículo buscando en múltiples extensiones (.jpg y .png).
        /// Si no existe en caché, intenta descargar desde FTP probando ambas extensiones.
        /// </summary>
        /// <param name="codigoArticulo">Código único del artículo</param>
        /// <returns>Ruta local del archivo o null si no se encuentra en ninguna extensión</returns>
        public async Task<string> ObtenerRutaLocalFotoAsync(string codigoArticulo)
        {
            if (string.IsNullOrWhiteSpace(codigoArticulo))
            {
                System.Diagnostics.Debug.WriteLine("[FTP Cache] código de artículo vacío");
                return null;
            }

            // Intentar con ambas extensiones: .jpg primero, luego .png
            string[] extensiones = { ".jpg", ".png" };

            foreach (var extension in extensiones)
            {
                string resultado = await ObtenerRutaLocalFotoConExtensionAsync(codigoArticulo, extension).ConfigureAwait(false);
                if (resultado != null)
                    return resultado;
            }

            System.Diagnostics.Debug.WriteLine($"[FTP Cache] No se encontró foto para {codigoArticulo} en ninguna extensión");
            return null;
        }

        /// <summary>
        /// Crea un cliente FTP configurado con los parámetros del servicio.
        /// </summary>
        private AsyncFtpClient CrearFtpClient()
        {
            var client = new AsyncFtpClient(_ftpHost, _ftpUser, _ftpPassword, _ftpPort);
            // El servidor requiere TLS/SSL en el canal de control (FTPS explícito)
            client.Config.EncryptionMode = FtpEncryptionMode.Explicit;
            client.Config.ValidateAnyCertificate = true;
            client.Config.ConnectTimeout = 10000;
            client.Config.DataConnectionType = FtpDataConnectionType.PASV;
            return client;
        }

        /// <summary>
        /// Obtiene la ruta local de la foto del artículo con una extensión específica.
        /// Si no existe en caché, la descarga desde FTP.
        /// Si existe en caché pero la versión remota es más nueva, la re-descarga.
        /// </summary>
        /// <param name="codigoArticulo">Código único del artículo</param>
        /// <param name="extension">Extensión del archivo (.jpg o .png)</param>
        /// <returns>Ruta local del archivo o null si falla la descarga</returns>
        private async Task<string> ObtenerRutaLocalFotoConExtensionAsync(string codigoArticulo, string extension)
        {
            if (string.IsNullOrWhiteSpace(codigoArticulo))
                return null;

            // Asegurar que la extensión tenga el punto
            if (!extension.StartsWith("."))
                extension = "." + extension;

            string nombreArchivo = $"{codigoArticulo}{extension}";
            string rutaLocal = Path.Combine(_cacheFolder, nombreArchivo);

            // Construir ruta remota completa
            string rutaRemota = _ftpRemoteFolder.TrimEnd('/') + "/" + nombreArchivo;

            // ✅ Ya está en caché — verificar si la versión remota es más nueva
            if (File.Exists(rutaLocal))
            {
                DateTime fechaLocal = File.GetLastWriteTime(rutaLocal);
                System.Diagnostics.Debug.WriteLine($"[FTP Cache] Foto en caché: {nombreArchivo}  local={fechaLocal:yyyy-MM-dd HH:mm:ss}");

                try
                {
                    using (var client = CrearFtpClient())
                    {
                        await client.Connect().ConfigureAwait(false);
                        if (!client.IsConnected)
                        {
                            System.Diagnostics.Debug.WriteLine("[FTP Cache] No se pudo conectar para comparar fechas");
                            return rutaLocal; // Usar local si no hay conexión
                        }

                        if (await client.FileExists(rutaRemota).ConfigureAwait(false))
                        {
                            var infoRemoto = await client.GetObjectInfo(rutaRemota).ConfigureAwait(false);
                            DateTime? fechaRemota = infoRemoto?.Modified;
                            if (fechaRemota.HasValue)
                            {
                                System.Diagnostics.Debug.WriteLine($"[FTP Cache] Fecha remota: {fechaRemota.Value:yyyy-MM-dd HH:mm:ss}");
                                if (fechaRemota.Value > fechaLocal)
                                {
                                    System.Diagnostics.Debug.WriteLine("[FTP Cache] Remoto es más nuevo — re-descargando...");
                                    var status = await client.DownloadFile(rutaLocal, rutaRemota, FtpLocalExists.Overwrite).ConfigureAwait(false);
                                    if (status == FtpStatus.Success)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[FTP Cache] Re-descarga exitosa: {nombreArchivo}");
                                    }
                                    else
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[FTP Cache] Re-descarga fallida: {nombreArchivo}  status={status}");
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine("[FTP Cache] Caché vigente — no se descarga");
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[FTP Cache] Archivo remoto no existe: {rutaRemota}");
                        }

                        await client.Disconnect().ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[FTP Cache] Error comparando fechas: {ex.GetType().Name}: {ex.Message}");
                }

                return rutaLocal;
            }

            // ❌ No existe localmente, descargar desde FTP
            try
            {
                System.Diagnostics.Debug.WriteLine($"[FTP Cache] Descargando desde FTP: {nombreArchivo}");
                System.Diagnostics.Debug.WriteLine($"[FTP Cache] Host={_ftpHost}:{_ftpPort}  User={_ftpUser}  Folder={_ftpRemoteFolder}");

                using (var client = CrearFtpClient())
                {
                    await client.Connect().ConfigureAwait(false);

                    if (!client.IsConnected)
                    {
                        System.Diagnostics.Debug.WriteLine("[FTP Cache] No se pudo conectar al servidor FTP");
                        return null;
                    }

                    System.Diagnostics.Debug.WriteLine("[FTP Cache] Conexión FTP exitosa");
                    System.Diagnostics.Debug.WriteLine($"[FTP Cache] Ruta remota: {rutaRemota}");

                    bool existe = await client.FileExists(rutaRemota).ConfigureAwait(false);
                    if (!existe)
                    {
                        System.Diagnostics.Debug.WriteLine($"[FTP Cache] Archivo no existe en FTP: {rutaRemota}");
                        await client.Disconnect().ConfigureAwait(false);
                        return null;
                    }

                    System.Diagnostics.Debug.WriteLine("[FTP Cache] Archivo existe en FTP, descargando...");

                    var status = await client.DownloadFile(rutaLocal, rutaRemota, FtpLocalExists.Overwrite).ConfigureAwait(false);
                    await client.Disconnect().ConfigureAwait(false);

                    if (status == FtpStatus.Success && File.Exists(rutaLocal))
                    {
                        System.Diagnostics.Debug.WriteLine($"[FTP Cache] Descarga exitosa: {nombreArchivo} -> {rutaLocal}");
                        return rutaLocal;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[FTP Cache] Descarga fallida: {nombreArchivo}  status={status}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FTP Cache] Error: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Limpia toda la caché (borra todas las fotos descargadas).
        /// Útil para liberar espacio o forzar re-descarga de fotos actualizadas.
        /// </summary>
        public void LimpiarCache()
        {
            try
            {
                if (Directory.Exists(_cacheFolder))
                {
                    foreach (var file in Directory.GetFiles(_cacheFolder))
                    {
                        try 
                        { 
                            File.Delete(file); 
                            System.Diagnostics.Debug.WriteLine($"[FTP Cache] Eliminado: {Path.GetFileName(file)}");
                        } 
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[FTP Cache] Error al eliminar {Path.GetFileName(file)}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FTP Cache] Error al limpiar caché: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene el tamaño total de la caché en bytes.
        /// </summary>
        public long ObtenerTamanoCache()
        {
            try
            {
                if (!Directory.Exists(_cacheFolder))
                    return 0;

                long tamano = 0;
                foreach (var file in Directory.GetFiles(_cacheFolder))
                {
                    tamano += new FileInfo(file).Length;
                }
                return tamano;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Obtiene el número de fotos en caché.
        /// </summary>
        public int ObtenerCantidadFotosEnCache()
        {
            try
            {
                if (!Directory.Exists(_cacheFolder))
                    return 0;

                return Directory.GetFiles(_cacheFolder).Length;
            }
            catch
            {
                return 0;
            }
        }
    }
}

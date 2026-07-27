using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DistribuidoraPalomarNet;

namespace ImageViewerApp
{
    public partial class frmImagenZoom : Form
    {
        private string[] imageFiles = Array.Empty<string>();
        private int currentIndex = -1;
        private string baseImagePath;
        private string cacheFtpPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DistribuidoraPalomarNet", "FotosCache");

        public frmImagenZoom()
        {
            InitializeComponent();
            baseImagePath = Module.ObtenerCarpetaFotosBasePublica();
        }

        private async void btnSearch_Click(object sender, EventArgs e)
        {
            string code = txtCode.Text.Trim();
            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("Por favor ingrese un código", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 1. Mostrar de inmediato lo que ya existe en disco (local + caché FTP),
            //    sin esperar las descargas: así la foto que se ve en el catálogo aparece al instante.
            imageFiles = BuscarImagenesEnDisco(code);
            MostrarPrimeraImagenODefecto();

            // 2. Descargar en segundo plano las variantes que falten desde el FTP.
            await EnsurePhotosDownloadedAsync(code).ConfigureAwait(true);

            // 3. Refrescar la lista con lo nuevo descargado (sin perder la imagen en pantalla).
            string imagenActual = (currentIndex >= 0 && currentIndex < imageFiles.Length)
                ? imageFiles[currentIndex]
                : null;

            imageFiles = BuscarImagenesEnDisco(code);

            if (imageFiles.Length == 0)
            {
                MostrarPrimeraImagenODefecto();
                return;
            }

            int nuevoIndice = imagenActual != null ? Array.IndexOf(imageFiles, imagenActual) : -1;
            if (nuevoIndice >= 0)
            {
                currentIndex = nuevoIndice;
                UpdateNavigationButtons();
            }
            else
            {
                MostrarPrimeraImagenODefecto();
            }
        }

        /// <summary>
        /// Busca en disco (carpeta local y caché FTP) la foto principal y las variantes _f1.._f5
        /// del código indicado, en .jpg y .png. No accede a la red.
        /// </summary>
        private string[] BuscarImagenesEnDisco(string code)
        {
            var imagePaths = new System.Collections.Generic.List<string>();
            string[] extensiones = { ".jpg", ".png" };

            // 1. Buscar imagen original en carpeta local (.jpg y .png)
            foreach (var ext in extensiones)
            {
                var localOriginal = Path.Combine(baseImagePath, $"{code}{ext}");
                if (File.Exists(localOriginal))
                {
                    imagePaths.Add(localOriginal);
                    break; // Solo agregar una vez (prioridad a .jpg)
                }
            }

            // 2. Buscar imagen original en caché FTP (si no está en local) - .jpg y .png
            if (imagePaths.Count == 0)
            {
                foreach (var ext in extensiones)
                {
                    var cacheOriginal = Path.Combine(cacheFtpPath, $"{code}{ext}");
                    if (File.Exists(cacheOriginal))
                    {
                        imagePaths.Add(cacheOriginal);
                        break; // Solo agregar una vez (prioridad a .jpg)
                    }
                }
            }

            // 3. Buscar imágenes con sufijos _f1 a _f5 en carpeta local (.jpg y .png)
            for (int i = 1; i <= 5; i++)
            {
                bool encontrada = false;
                foreach (var ext in extensiones)
                {
                    var localSuffix = Path.Combine(baseImagePath, $"{code}_f{i}{ext}");
                    if (File.Exists(localSuffix))
                    {
                        imagePaths.Add(localSuffix);
                        encontrada = true;
                        break; // Solo agregar una vez por sufijo (prioridad a .jpg)
                    }
                }

                // 4. Si no está en local, buscar la variante en caché FTP (.jpg y .png)
                if (!encontrada)
                {
                    foreach (var ext in extensiones)
                    {
                        var cacheSuffix = Path.Combine(cacheFtpPath, $"{code}_f{i}{ext}");
                        if (File.Exists(cacheSuffix))
                        {
                            imagePaths.Add(cacheSuffix);
                            break; // Solo agregar una vez por sufijo (prioridad a .jpg)
                        }
                    }
                }
            }

            return imagePaths.ToArray();
        }

        /// <summary>
        /// Muestra la primera imagen de la lista, o la imagen por defecto SINFOTO.png si no hay ninguna.
        /// </summary>
        private void MostrarPrimeraImagenODefecto()
        {
            if (imageFiles.Length > 0)
            {
                currentIndex = 0;
                DisplayCurrentImage();
            }
            else
            {
                currentIndex = -1;
                try
                {
                    if (pictureBox1.Image != null)
                        pictureBox1.Image.Dispose();

                    string rutaSinFoto = Module.ObtenerRutaImagenPorDefectoPublica();
                    pictureBox1.Image = File.Exists(rutaSinFoto) ? Image.FromFile(rutaSinFoto) : null;
                }
                catch
                {
                    pictureBox1.Image = null;
                }
            }

            UpdateNavigationButtons();
        }

        private static void TryCopyCacheToBase(string rutaFotoCache, string basePath)
        {
            if (string.IsNullOrWhiteSpace(rutaFotoCache) || !File.Exists(rutaFotoCache))
                return;

            try
            {
                Directory.CreateDirectory(basePath);
                string rutaFotoBase = Path.Combine(basePath, Path.GetFileName(rutaFotoCache));
                bool debeCopiar = true;
                if (File.Exists(rutaFotoBase))
                {
                    DateTime baseWrite = File.GetLastWriteTime(rutaFotoBase);
                    DateTime cacheWrite = File.GetLastWriteTime(rutaFotoCache);
                    debeCopiar = cacheWrite > baseWrite;
                }

                if (debeCopiar)
                    File.Copy(rutaFotoCache, rutaFotoBase, overwrite: true);
            }
            catch
            {
            }
        }

        private async Task EnsurePhotosDownloadedAsync(string code)
        {
            if (Module._fotoService == null)
                return;

            string basePath = baseImagePath;
            string[] variantes = new string[6];
            variantes[0] = code;
            for (int i = 1; i <= 5; i++)
                variantes[i] = $"{code}_f{i}";

            foreach (var variante in variantes)
            {
                try
                {
                    string rutaFotoCache = await Module._fotoService.ObtenerRutaLocalFotoAsync(variante).ConfigureAwait(true);
                    if (rutaFotoCache != null && File.Exists(rutaFotoCache))
                        TryCopyCacheToBase(rutaFotoCache, basePath);
                }
                catch
                {
                }
            }
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentIndex > 0)
                {
                    currentIndex--;
                    DisplayCurrentImage();
                    UpdateNavigationButtons();
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentIndex < imageFiles.Length - 1)
                {
                    currentIndex++;
                    DisplayCurrentImage();
                    UpdateNavigationButtons();
                }
            }
            catch (Exception ex)
            {
            }

        }

        private void DisplayCurrentImage()
        {
            if (currentIndex >= 0 && currentIndex < imageFiles.Length)
            {
                try
                {
                    if (pictureBox1.Image != null)
                    {
                        pictureBox1.Image.Dispose();
                    }
                    pictureBox1.Image = Image.FromFile(imageFiles[currentIndex]);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al cargar la imagen: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void UpdateNavigationButtons()
        {
            btnPrevious.Enabled = currentIndex > 0;
            btnNext.Enabled = currentIndex < imageFiles.Length - 1;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                pictureBox1.Image.Dispose();
            }
            base.OnFormClosing(e);
        }

        private void frmImagenZoom_Load(object sender, EventArgs e)
        {
            btnSearch_Click(sender, e);
        }
    }
}

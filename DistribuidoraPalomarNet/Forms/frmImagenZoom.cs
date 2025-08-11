using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ImageViewerApp
{
    public partial class frmImagenZoom : Form
    {
        private string[] imageFiles;
        private int currentIndex = -1;
        private string baseImagePath = @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\imagenes\FOTOS";

        public frmImagenZoom()
        {
            InitializeComponent();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string code = txtCode.Text.Trim();
            if (string.IsNullOrEmpty(code))
            {
                MessageBox.Show("Por favor ingrese un código", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Primero buscar la imagen original
            var originalImage = Path.Combine(baseImagePath, $"{code}.jpg");

            // Luego buscar las imágenes con sufijos _f1 a _f5
            var suffixImages = Enumerable.Range(1, 5)
                                .Select(i => Path.Combine(baseImagePath, $"{code}_f{i}.jpg"))
                                .Where(File.Exists)
                                .ToArray();

            // Combinar la imagen original (si existe) con las imágenes con sufijos
            imageFiles = (File.Exists(originalImage) ? new[] { originalImage } : Array.Empty<string>())
                            .Concat(suffixImages)
                            .ToArray();

            if (imageFiles.Length == 0)
            {
                MessageBox.Show("No se encontraron imágenes para este código", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                pictureBox1.Image = null;
                return;
            }

            currentIndex = 0;
            DisplayCurrentImage();
            UpdateNavigationButtons();
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

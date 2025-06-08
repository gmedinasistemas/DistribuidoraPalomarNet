namespace DistribuidoraPalomarNet
{
    public partial class frmImagen : Form
    {
        public frmImagen()
        {
            InitializeComponent();
        }

        private void frmImagen_Load(object sender, EventArgs e)
        {
            try
            {
                this.WindowState = FormWindowState.Maximized;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar la imagen: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

using System;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                var mainForm = new frmPrincipal();
                Module.SetMainForm(mainForm);
                Application.Run(mainForm);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al iniciar la aplicación: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet.Forms
{
    public partial class frmMensajePersonalizado : Form
    {
        public frmMensajePersonalizado()
        {
            InitializeComponent();
        }

        static int IdBoton;
        static frmMensajePersonalizado nuevoMessageBox;

        // Función para mostrar el mensaje con el parámetro solo de texto
        public static int Show(string texto)
        {
            nuevoMessageBox = new frmMensajePersonalizado();
            nuevoMessageBox.LblMensaje.Text = texto;
            nuevoMessageBox.ShowDialog();
            return IdBoton;
        }

        // Función para mostrar el mensaje con el parámetro solo de texto y parámetro de qué botón específico (1=aceptar, 2=cancelar)
        public static int Show(string texto, int boton)
        {
            nuevoMessageBox = new frmMensajePersonalizado();
            nuevoMessageBox.LblMensaje.Text = texto;
            if (boton == 1)
            {
                nuevoMessageBox.BtnOk.Location = new Point(313, 138);
                nuevoMessageBox.BtnCancelar.Visible = false;
            }
            else if (boton == 2)
            {
                nuevoMessageBox.BtnOk.Visible = false;
            }
            nuevoMessageBox.ShowDialog();
            return IdBoton;
        }

        // Función para mostrar el mensaje con el parámetro solo de texto, color de fondo y color de letra
        public static int Show(string texto, string caption, Color fondo, Color colorletra)
        {
            nuevoMessageBox = new frmMensajePersonalizado();
            nuevoMessageBox.lblcaption.Text = caption;
            nuevoMessageBox.LblMensaje.Text = texto;
            nuevoMessageBox.LblMensaje.ForeColor = colorletra;
            nuevoMessageBox.LblMensaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            nuevoMessageBox.BackColor = fondo;
            nuevoMessageBox.ShowDialog();
            return IdBoton;
        }

        // Función para mostrar el mensaje con el parámetro solo de texto, parámetro de qué botón específico (1=aceptar, 2=cancelar), color de fondo y color de letra
        public static int Show(string texto, int boton, Color fondo, Color colorletra)
        {
            nuevoMessageBox = new frmMensajePersonalizado();
            nuevoMessageBox.LblMensaje.Text = texto;
            if (boton == 1)
            {
                nuevoMessageBox.BtnOk.Location = new Point(313, 165);
                nuevoMessageBox.BtnCancelar.Visible = false;
            }
            else if (boton == 2)
            {
                nuevoMessageBox.BtnOk.Visible = false;
            }
            nuevoMessageBox.BackColor = fondo;
            nuevoMessageBox.ForeColor = colorletra;
            nuevoMessageBox.ShowDialog();
            return IdBoton;
        }

        // Método que va a devolver qué botón se hizo click y luego destruir el formulario-mensaje
        private void BtnOk_Click(object sender, EventArgs e)
        {
            IdBoton = 1;
            nuevoMessageBox.Dispose();
        }

        // Método que va a devolver qué botón se hizo click y luego destruir el formulario-mensaje
        private void BtnCancelar_Click(object sender, EventArgs e)
        {
            IdBoton = 2;
            nuevoMessageBox.Dispose();
        }
    }
}

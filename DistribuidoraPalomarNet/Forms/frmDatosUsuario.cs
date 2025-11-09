using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace DistribuidoraPalomarNet
{
    public partial class frmDatosUsuario : Form
    {
        public frmDatosUsuario()
        {
            InitializeComponent();
            Module.SetDatosUsuarioForm(this);
        }

        #region Metodos
        private void frmDatosUsuario_Load(object sender, EventArgs e)
        {
            try
            {
                if (ModuloUsuario.existeUsuario())
                {
                    Module.accionDatosUsu = 1;
                    Module.CargarDatos();
                }
                else
                {
                    Module.accionDatosUsu = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al iniciar el 'Formulario de Usuario': " + ex.Message.ToString(), "Problema de acceso", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void txtCodCliente_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter || e.KeyChar == (char)Keys.Tab)
                {
                    if (!string.IsNullOrEmpty(txtCodCliente.Text.Trim()) && decimal.TryParse(txtCodCliente.Text.Trim(), out _))
                    {
                        if (!Module.ValidarCodiCliente(txtCodCliente.Text))
                        {
                            txtCodCliente.BackColor = Color.DarkRed;
                            txtCodCliente.Focus();
                            txtCuitCliente.Enabled = false;
                        }
                        else
                        {
                            txtCodCliente.BackColor = Color.White;
                            txtCuitCliente.Enabled = true;
                            txtCuitCliente.Focus();
                        }
                    }
                }
                else
                {
                    if (char.IsNumber(e.KeyChar))
                    {
                        e.Handled = false; // habilita la escritura
                    }
                    else if (char.IsControl(e.KeyChar))
                    {
                        e.Handled = false;
                    }
                    else
                    {
                        e.Handled = true;
                    }
                }
            }
            catch (Exception)
            {
                txtCuitCliente.Enabled = false;
            }
        }

        private void txtCuitCliente_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == (char)Keys.Enter || e.KeyChar == (char)Keys.Tab)
                {
                    if (!string.IsNullOrEmpty(txtCodCliente.Text.Trim()) && 
                        decimal.TryParse(txtCodCliente.Text.Trim(), out _) && 
                        decimal.TryParse(txtCuitCliente.Text.Trim(), out _))
                    {
                        if (!Module.ValidarCuitCliente(txtCodCliente.Text.Trim(), txtCuitCliente.Text.Trim()))
                        {
                            txtCuitCliente.BackColor = Color.DarkRed;
                            txtCuitCliente.Focus();
                            txtEmail.Enabled = false;
                            btnGrabar.Enabled = false;
                        }
                        else
                        {
                            txtCuitCliente.BackColor = Color.White;
                            btnGrabar.Enabled = true;
                            txtEmail.Enabled = true;
                            txtEmail.Focus();
                        }
                    }
                }
                else
                {
                    if (char.IsNumber(e.KeyChar))
                    {
                        e.Handled = false; // habilita la escritura
                    }
                    else if (char.IsControl(e.KeyChar))
                    {
                        e.Handled = false;
                    }
                    else
                    {
                        e.Handled = true;
                    }
                }
            }
            catch (Exception)
            {
                txtEmail.Enabled = false;
                btnGrabar.Enabled = false;
            }
        }

        private void btnGrabar_Click(object sender, EventArgs e)
        {
            try
            {
                btnGrabar.Enabled = false;
                if (!string.IsNullOrEmpty(txtCodCliente.Text.Trim()))
                {
                    if (!string.IsNullOrEmpty(txtCuitCliente.Text.Trim()))
                    {
                        if (!string.IsNullOrEmpty(txtDesc1.Text.Trim()))
                        {
                            if (decimal.TryParse(txtDesc1.Text.Trim(), out _))
                            {
                                if (!string.IsNullOrEmpty(txtDesc2.Text.Trim()))
                                {
                                    if (decimal.TryParse(txtDesc2.Text.Trim(), out _))
                                    {
                                        if (!string.IsNullOrEmpty(txtBene.Text.Trim()))
                                        {
                                            if (decimal.TryParse(txtBene.Text.Trim(), out _))
                                            {
                                                // Grabo todos los datos los paso como variables a la funcion
                                                if (Module.accionDatosUsu == 1)
                                                {
                                                    // Para Actualizar
                                                    ModuloUsuario.UpdateUsuario(txtCodCliente.Text, txtCuitCliente.Text, txtEmail.Text, Convert.ToDecimal(txtDesc1.Text), Convert.ToDecimal(txtDesc2.Text), Module.US_CLRAZO, Module.US_CLCIVA.ToString(), Module.US_CLVEND.ToString(), Module.US_CLCPAG.ToString(), txtBene.Text);
                                                }
                                                else
                                                {
                                                    // por primera vez
                                                    ModuloUsuario.GrabarUsuario(txtCodCliente.Text, txtCuitCliente.Text, txtEmail.Text, Convert.ToDecimal(txtDesc1.Text), Convert.ToDecimal(txtDesc2.Text), Module.US_CLRAZO, Module.US_CLCIVA.ToString(), Module.US_CLVEND.ToString(), Module.US_CLCPAG.ToString(), txtBene.Text);
                                                }

                                                // +++++++++++++++++++++++COMENTO TEMPORARIAMENTE+++++++++++++++++++++++++++
                                                Module.Cargar_Usuario();
                                                Module.aplicar_desc_Cliente();
                                                Module.Calculadora();
                                                // aplicar_descuento3()
                                                // aplicar_beneficio()
                                                // Cargar_Usuario()
                                            }
                                            else
                                            {
                                                txtBene.BackColor = Color.Red;
                                                MessageBox.Show("Debe ingresar datos correctos", "Solo Numeros enteros", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                txtBene.BackColor = Color.White;
                                                txtBene.Focus();
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Debe ingresar el Beneficio", "falta de datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            txtBene.Focus();
                                        }
                                    }
                                    else
                                    {
                                        txtDesc2.BackColor = Color.Red;
                                        MessageBox.Show("Debe ingresar datos correctos", "Solo Numeros enteros", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        txtDesc2.BackColor = Color.White;
                                        txtDesc2.Focus();
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Debe ingresar el Descuento 2", "falta de datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    txtDesc2.Focus();
                                }
                            }
                            else
                            {
                                txtDesc1.BackColor = Color.Red;
                                MessageBox.Show("Debe ingresar datos correctos", "Solo Numeros enteros", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                txtDesc1.BackColor = Color.White;
                                txtDesc1.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Debe ingresar el Descuento 1", "falta de datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtDesc1.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Debe ingresar el Cuit de Cliente", "falta de datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtCuitCliente.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("Debe ingresar el Código de Cliente", "falta de datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtCodCliente.Focus();
                }
                btnGrabar.Enabled = true;
            }
            catch (Exception ex)
            {
                btnGrabar.Enabled = true;
                MessageBox.Show("Error dentro del 'btnGrabar': " + ex.Message.ToString(), "Inconveniente en la acción", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            if (ptbImgSi.Visible)
            {
                Module.aplicar_desc_Cliente(); // Cargo Descuentos 1 y 2
                Module.Update_Precios();
                Module.Update_Precios_Ofertas();
                //Module.Update_PrecioSug_Ofertas();
                Module.funcion_inicial(1);
            }
            Module.Cargar_Usuario();
            this.Close();
        }
        #endregion
    }
}

using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace DistribuidoraPalomarNet
{
    public partial class frmDatosUsuario : Form
    {
        // Inicializa los controles del formulario y registra esta instancia en Module
        // para que otros módulos puedan referenciarla (ej: CargarDatos, aplicar_desc_Cliente).
        public frmDatosUsuario()
        {
            InitializeComponent();
            Module.SetDatosUsuarioForm(this);
        }

        // Actualiza los indicadores visuales de estado del cliente en frmPrincipal
        // (ej: mensaje de cliente restringido, habilitación de secciones).
        // Se llama tras validar o modificar datos del cliente.
        private void RefreshEstadoEnPrincipal()
        {
            try
            {
                var fp = Application.OpenForms.OfType<frmPrincipal>().FirstOrDefault();
                if (fp != null)
                {
                    fp.ApplyEstadoClienteToUI();
                }
            }
            catch
            {
            }
        }

        #region Metodos
        // Al abrir el formulario: los campos de descuento quedan en solo lectura (se cargan desde la base del distribuidor).
        // Si ya existe un usuario registrado, carga sus datos y valida los descuentos asignados.
        // Según si existe (accionDatosUsu=1) o no (0), btnGrabar actuará como UPDATE o INSERT.
        private void frmDatosUsuario_Load(object sender, EventArgs e)
        {
            try
            {
                txtDesc1.ReadOnly = true;
                txtDesc2.ReadOnly = true;

                if (ModuloUsuario.existeUsuario())
                {
                    Module.accionDatosUsu = 1;
                    Module.CargarDatos();

                    if (!string.IsNullOrWhiteSpace(txtCodCliente.Text) && !string.IsNullOrWhiteSpace(txtCuitCliente.Text))
                    {
                        if (Module.ValidarCuitCliente(txtCodCliente.Text.Trim(), txtCuitCliente.Text.Trim()))
                        {
                            txtDesc1.Text = Module.ClienteDescuento1.ToString();
                            txtDesc2.Text = Module.ClienteDescuento2.ToString();
                        }
                    }
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

        // Controla la entrada en el campo Código de Cliente:
        // - Solo acepta dígitos (bloquea letras y caracteres especiales).
        // - Al presionar Enter o Tab valida el código contra la base del distribuidor:
        //   válido → habilita txtCuitCliente y pasa el foco; inválido → fondo rojo y bloquea el CUIT.
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
                            // Código no encontrado en la base del distribuidor
                            txtCodCliente.BackColor = Color.DarkRed;
                            txtCodCliente.Focus();
                            txtCuitCliente.Enabled = false;
                        }
                        else
                        {
                            // Código válido: habilitar siguiente campo
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
                        e.Handled = true; // bloquea caracteres no numéricos
                    }
                }
            }
            catch (Exception)
            {
                txtCuitCliente.Enabled = false;
            }
        }

        // Controla la entrada en el campo CUIT del cliente:
        // - Solo acepta dígitos.
        // - Al presionar Enter o Tab cruza el par Código+CUIT contra la base del distribuidor:
        //   válido → carga descuentos asignados, habilita email y btnGrabar;
        //   inválido → fondo rojo, bloquea email y btnGrabar.
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
                            // CUIT no coincide con el código en la base del distribuidor
                            txtCuitCliente.BackColor = Color.DarkRed;
                            txtCuitCliente.Focus();
                            txtEmail.Enabled = false;
                            btnGrabar.Enabled = false;
                        }
                        else
                        {
                            // Par código+CUIT válido: cargar descuentos y habilitar formulario
                            txtCuitCliente.BackColor = Color.White;
                            btnGrabar.Enabled = true;
                            txtEmail.Enabled = true;
                            txtEmail.Focus();

                            // Descuentos asignados al cliente en la base del distribuidor (solo lectura)
                            txtDesc1.Text = Module.ClienteDescuento1.ToString();
                            txtDesc2.Text = Module.ClienteDescuento2.ToString();

                            if (Module.ClienteRestringido && !string.IsNullOrWhiteSpace(Module.ClienteMensajeValidacion))
                            {
                                MessageBox.Show(Module.ClienteMensajeValidacion, "Acceso restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                            RefreshEstadoEnPrincipal();
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

        // Graba o actualiza los datos del usuario tras validar todos los campos.
        // Valida de forma anidada: código → CUIT → Desc1 → Desc2 → Beneficio.
        // Tras guardar: recarga el usuario en Module, aplica descuentos y recalcula precios en frmPrincipal.
        private void btnGrabar_Click(object sender, EventArgs e)
        {
            try
            {
                btnGrabar.Enabled = false;
                if (!string.IsNullOrEmpty(txtCodCliente.Text.Trim()))
                {
                    if (!string.IsNullOrEmpty(txtCuitCliente.Text.Trim()))
                    {
                        if (!Module.ValidarCuitCliente(txtCodCliente.Text.Trim(), txtCuitCliente.Text.Trim()))
                        {
                            MessageBox.Show("Datos inválidos", "Acceso restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            btnGrabar.Enabled = true;
                            return;
                        }

                        txtDesc1.Text = Module.ClienteDescuento1.ToString();
                        txtDesc2.Text = Module.ClienteDescuento2.ToString();

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
                                                // accionDatosUsu: 1 = UPDATE (usuario ya registrado), 0 = INSERT (primera vez)
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

                                                Module.ValidarCuitCliente(txtCodCliente.Text.Trim(), txtCuitCliente.Text.Trim());
                                                RefreshEstadoEnPrincipal();
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

        // Cierra el formulario. Si hubo un guardado exitoso (ptbImgSi visible):
        // 1. Recarga los descuentos del cliente.
        // 2. Recalcula los precios generales y de ofertas.
        // 3. Reaplica el filtro activo en frmPrincipal (preserva la búsqueda que el usuario tenía).
        private void btnSalir_Click(object sender, EventArgs e)
        {
            if (ptbImgSi.Visible)
            {
                Module.aplicar_desc_Cliente(); // Cargo Descuentos 1 y 2
                Module.Update_Precios();
                Module.Update_Precios_Ofertas();

                //Module.Update_PrecioSug_Ofertas();

                // Solo llamar a Module.funcion_inicial(1) si existe el frmPrincipal
                // y tiene datos visibles en la grilla
                var fp = Application.OpenForms.OfType<frmPrincipal>().FirstOrDefault();

                if (fp != null && fp.DataGridView1 != null)
                {
                    // Hay al menos una fila real (no la fila nueva vacía)
                    bool hayFilasConDatos = fp.DataGridView1.Rows
                        .Cast<DataGridViewRow>()
                        .Any(r => !r.IsNewRow);

                    if (hayFilasConDatos)
                    {
                        fp.RefrescarSegunFiltros();
                    }
                }
            }

            Module.ValidarCuitCliente(txtCodCliente.Text.Trim(), txtCuitCliente.Text.Trim());
            RefreshEstadoEnPrincipal();
            Module.Cargar_Usuario(false);
            this.Close();
        }
        #endregion
    }
}

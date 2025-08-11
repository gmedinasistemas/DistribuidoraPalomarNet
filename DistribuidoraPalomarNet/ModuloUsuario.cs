using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    public static class ModuloUsuario
    {
        #region Funciones
        public static frmDatosUsuario _frmDatosUsuario;

        /// <summary>
        /// Valida si existe el código de cliente en la base de datos y actualiza los datos globales.
        /// </summary>
        public static bool ValidarCodiCliente(string codi)
        {
            OleDbCommand cmaccess = new OleDbCommand();
            OleDbDataReader draccess = null;
            try
            {
                if (_frmDatosUsuario == null) throw new InvalidOperationException("Formulario de datos de usuario no inicializado");

                Module.ConectarDBArti();
                cmaccess.Connection = Module.cnn;
                cmaccess.CommandType = CommandType.Text;
                cmaccess.CommandText = $"select clcodi,clrazo,clciva,clvend,clcpag from clientes WHERE clcodi={codi}";
                draccess = cmaccess.ExecuteReader();

                if (draccess.Read())
                {
                    _frmDatosUsuario.lblUsuario.Text = draccess["clrazo"].ToString();
                    Module.US_CLRAZO = draccess["clrazo"].ToString();
                    Module.US_CLCIVA = !string.IsNullOrEmpty(draccess["clciva"].ToString()) ? Convert.ToInt32(draccess["clciva"]) : 0;
                    Module.US_CLVEND = !string.IsNullOrEmpty(draccess["clvend"].ToString()) ? Convert.ToInt32(draccess["clvend"]) : 0;
                    Module.US_CLCPAG = !string.IsNullOrEmpty(draccess["clcpag"].ToString()) ? Convert.ToInt32(draccess["clcpag"]) : 0;
                    Module.ExisteCodi_Cliente = true;
                }
                else
                {
                    Module.US_CLRAZO = string.Empty;
                    Module.US_CLCIVA = 0;
                    Module.US_CLVEND = 0;
                    Module.US_CLCPAG = 0;
                    Module.ExisteCodi_Cliente = false;
                }

                draccess.Close();
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();

                return Module.ExisteCodi_Cliente;
            }
            catch (Exception)
            {
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
                return false;
            }
            finally
            {
                if (draccess != null && !draccess.IsClosed)
                    draccess.Close();
                cmaccess.Dispose();
            }
        }

        /// <summary>
        /// Valida si existe el cliente por código y CUIT en la base de datos y actualiza los datos globales.
        /// </summary>
        public static bool ValidarCuitCliente(string codi, string cuit)
        {
            OleDbCommand cmaccess = new OleDbCommand();
            OleDbDataReader draccess = null;
            try
            {
                Module.ConectarDBArti();
                cmaccess.Connection = Module.cnn;
                cmaccess.CommandType = CommandType.Text;
                cmaccess.CommandText = $"select clcodi,clrazo,clciva,clvend,clcpag from clientes WHERE clcodi={codi} and clcuit={cuit}";
                draccess = cmaccess.ExecuteReader();

                if (draccess.Read())
                {
                    Module.US_CLRAZO = draccess["clrazo"].ToString();
                    Module.US_CLCIVA = !string.IsNullOrEmpty(draccess["clciva"].ToString()) ? Convert.ToInt32(draccess["clciva"]) : 0;
                    Module.US_CLVEND = !string.IsNullOrEmpty(draccess["clvend"].ToString()) ? Convert.ToInt32(draccess["clvend"]) : 0;
                    Module.US_CLCPAG = !string.IsNullOrEmpty(draccess["clcpag"].ToString()) ? Convert.ToInt32(draccess["clcpag"]) : 0;
                    Module.ExisteCuit_Cliente = true;
                }
                else
                {
                    Module.US_CLRAZO = string.Empty;
                    Module.US_CLCIVA = 0;
                    Module.US_CLVEND = 0;
                    Module.US_CLCPAG = 0;
                    Module.ExisteCuit_Cliente = false;
                }

                draccess.Close();
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();

                return Module.ExisteCuit_Cliente;
            }
            catch (Exception)
            {
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
                return false;
            }
            finally
            {
                if (draccess != null && !draccess.IsClosed)
                    draccess.Close();
                cmaccess.Dispose();
            }
        }


        public static void SetDatosUsuarioForm(frmDatosUsuario form)
        {
            _frmDatosUsuario = form;
        }

        public static void GrabarUsuario(string p1, string p2, string p3, decimal p4, decimal p5, string p6, string p7, string p8, string p9, string bene)
        {
            using var cmd = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmd.Connection = Module.cnnU;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"INSERT INTO USUARIO 
                    (US_CODI, US_CUIT, US_MAIL, US_DESCUENTO1, US_DESCUENTO2, US_CLRAZO, US_CLCIVA, US_CLVEND, US_CLCPAG, US_BENEFICIO) 
                    VALUES (@codi, @cuit, @mail, @desc1, @desc2, @clrazo, @clciva, @clvend, @clcpag, @bene)";

                cmd.Parameters.AddWithValue("@codi", p1);
                cmd.Parameters.AddWithValue("@cuit", p2);
                cmd.Parameters.AddWithValue("@mail", p3);
                cmd.Parameters.AddWithValue("@desc1", p4);
                cmd.Parameters.AddWithValue("@desc2", p5);
                cmd.Parameters.AddWithValue("@clrazo", p6);
                cmd.Parameters.AddWithValue("@clciva", p7);
                cmd.Parameters.AddWithValue("@clvend", p8);
                cmd.Parameters.AddWithValue("@clcpag", p9);
                cmd.Parameters.AddWithValue("@bene", bene);

                cmd.ExecuteNonQuery();
                Module.US_CODI = p1;

                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = true;
                _frmDatosUsuario.ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = false;
                _frmDatosUsuario.ptbImgNo.Visible = true;
                throw new Exception($"Error en GrabarUsuario: {ex.Message}");
            }
        }

        public static bool existeUsuario()
        {
            var cmUsu = new OleDbCommand();
            OleDbDataReader drUsu = null;
            var datos = false;

            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = "SELECT US_CODI, US_CUIT, US_MAIL from USUARIO";
                drUsu = cmUsu.ExecuteReader();
                if (drUsu.Read())
                {
                    datos = true;
                }
                drUsu.Close();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                return datos;
            }
            catch (Exception)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                return false;
            }
            finally
            {
                if (drUsu != null && !drUsu.IsClosed)
                    drUsu.Close();
                cmUsu.Dispose();
            }
        }

        public static void UpdateUsuario(string p1, string p2, string p3, decimal p4, decimal p5, string p6, string p7, string p8, string p9, string bene)
        {
            using var cmd = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmd.Connection = Module.cnnU;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"UPDATE USUARIO SET 
                    US_CODI = @codi,
                    US_CUIT = @cuit,
                    US_MAIL = @mail,
                    US_DESCUENTO1 = @desc1,
                    US_DESCUENTO2 = @desc2,
                    US_USUARIO = @usuario,
                    US_CLCIVA = @clciva,
                    US_CLVEND = @clvend,
                    US_CLCPAG = @clcpag,
                    US_BENEFICIO = @bene";

                cmd.Parameters.AddWithValue("@codi", p1);
                cmd.Parameters.AddWithValue("@cuit", p2);
                cmd.Parameters.AddWithValue("@mail", p3);
                cmd.Parameters.AddWithValue("@desc1", p4);
                cmd.Parameters.AddWithValue("@desc2", p5);
                cmd.Parameters.AddWithValue("@usuario", p6);
                cmd.Parameters.AddWithValue("@clciva", p7);
                cmd.Parameters.AddWithValue("@clvend", p8);
                cmd.Parameters.AddWithValue("@clcpag", p9);
                cmd.Parameters.AddWithValue("@bene", bene);

                cmd.ExecuteNonQuery();
                Module.US_CODI = p1;

                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = true;
                _frmDatosUsuario.ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = false;
                _frmDatosUsuario.ptbImgNo.Visible = true;
                throw new Exception($"Error en UpdateUsuario: {ex.Message}");
            }
        }

        public static void CargarDatos()
        {
            using var cmd = new OleDbCommand();
            try
            {
                if (_frmDatosUsuario == null) throw new InvalidOperationException("Formulario de datos de usuario no inicializado");

                Module.ConectarDBUsuario();
                cmd.Connection = Module.cnnU;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT US_CODI, US_CUIT, US_MAIL, US_DESCUENTO1, US_DESCUENTO2, US_BENEFICIO from USUARIO";
                
                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    _frmDatosUsuario.txtCodCliente.Text = reader["US_CODI"].ToString();
                    _frmDatosUsuario.txtCuitCliente.Text = reader["US_CUIT"].ToString();
                    _frmDatosUsuario.txtEmail.Text = reader["US_MAIL"].ToString();
                    _frmDatosUsuario.txtDesc1.Text = reader["US_DESCUENTO1"].ToString();
                    _frmDatosUsuario.txtDesc2.Text = reader["US_DESCUENTO2"].ToString();
                    _frmDatosUsuario.txtBene.Text = reader["US_BENEFICIO"].ToString();
                    ValidarCodiCliente(_frmDatosUsuario.txtCodCliente.Text);
                }

                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                throw new Exception($"Error en CargarDatos: {ex.Message}");
            }
        }

        public static void CargarImagenes()
        {
            try
            {
                if (_frmDatosUsuario == null) throw new InvalidOperationException("Formulario de datos de usuario no inicializado");

                string imgSiPath = System.IO.Path.Combine(Application.StartupPath, "imagenes", "si.jpg");
                string imgNoPath = System.IO.Path.Combine(Application.StartupPath, "imagenes", "no.jpg");

                if (System.IO.File.Exists(imgSiPath))
                {
                    _frmDatosUsuario.ptbImgSi.Image = Image.FromFile(imgSiPath);
                }
                if (System.IO.File.Exists(imgNoPath))
                {
                    _frmDatosUsuario.ptbImgNo.Image = Image.FromFile(imgNoPath);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en CargarImagenes: {ex.Message}");
            }
        }

        public static void ValidarDatosUsuario()
        {
            try
            {
                if (!string.IsNullOrEmpty(_frmDatosUsuario.txtCodCliente.Text) &&
                    !string.IsNullOrEmpty(_frmDatosUsuario.txtCuitCliente.Text))
                {
                    _frmDatosUsuario.ptbImgNo.Visible = false;
                    _frmDatosUsuario.ptbImgSi.Visible = true;
                    _frmDatosUsuario.txtEmail.Enabled = true;
                    _frmDatosUsuario.btnGrabar.Enabled = true;
                }
                else
                {
                    _frmDatosUsuario.ptbImgNo.Visible = true;
                    _frmDatosUsuario.ptbImgSi.Visible = false;
                    _frmDatosUsuario.txtEmail.Enabled = false;
                    _frmDatosUsuario.btnGrabar.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al validar los datos del usuario: {ex.Message}",
                    "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void ValidarDescuentos()
        {
            try
            {
                if (!string.IsNullOrEmpty(_frmDatosUsuario.txtDesc1.Text) &&
                    !string.IsNullOrEmpty(_frmDatosUsuario.txtDesc2.Text))
                {
                    if (decimal.TryParse(_frmDatosUsuario.txtDesc1.Text, out decimal desc1) &&
                        decimal.TryParse(_frmDatosUsuario.txtDesc2.Text, out decimal desc2))
                    {
                        if (desc1 >= 0 && desc1 <= 100 && desc2 >= 0 && desc2 <= 100)
                        {
                            _frmDatosUsuario.ptbImgNo.Visible = false;
                            _frmDatosUsuario.ptbImgSi.Visible = true;
                            _frmDatosUsuario.btnGrabar.Enabled = true;
                            return;
                        }
                    }
                }

                _frmDatosUsuario.ptbImgNo.Visible = true;
                _frmDatosUsuario.ptbImgSi.Visible = false;
                _frmDatosUsuario.btnGrabar.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al validar los descuentos: {ex.Message}",
                    "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}

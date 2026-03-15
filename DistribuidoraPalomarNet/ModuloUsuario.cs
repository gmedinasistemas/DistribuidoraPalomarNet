using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    /// <summary>
    /// Módulo estático que centraliza la autenticación, registro y gestión de datos del usuario-cliente.
    /// </summary>
    /// <remarks>
    /// Opera sobre dos bases de datos distintas:
    /// <para><b>Base del distribuidor</b> (Module.cnn / ini_ruta_base):<br/>
    /// Tabla <c>clientes</c> — valida código de cliente, CUIT y lee los descuentos asignados.</para>
    /// <para><b>Base del usuario local</b> (Module.cnnU / ini_ruta_usu):<br/>
    /// Tabla <c>USUARIO</c> — almacena el registro del cliente que opera con la aplicación.</para>
    /// </remarks>
    public static class ModuloUsuario
    {
        #region Funciones

        /// <summary>
        /// Referencia a la instancia activa de <c>frmDatosUsuario</c>.
        /// Se asigna desde el constructor del formulario mediante <see cref="SetDatosUsuarioForm"/>.
        /// Todos los métodos que actualizan controles del formulario dependen de esta referencia.
        /// </summary>
        public static frmDatosUsuario _frmDatosUsuario;

        /// <summary>
        /// Busca un cliente por código en la base del distribuidor (tabla <c>clientes</c>).
        /// Si existe, actualiza las variables globales de Module con los datos del cliente
        /// (razón social, condición IVA, vendedor, condición de pago) y marca <c>ExisteCodi_Cliente = true</c>.
        /// Si no existe, limpia esas variables.
        /// </summary>
        /// <param name="codi">Código numérico del cliente a buscar.</param>
        /// <returns>
        /// <c>true</c> si el código existe en la base del distribuidor;
        /// <c>false</c> si no se encuentra o ante cualquier error de base de datos.
        /// </returns>
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

                // Cliente encontrado: poblar el label del formulario y actualizar variables globales de Module
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
                    // Código no encontrado: limpiar variables globales y marcar cliente como inexistente
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
        /// Valida el par código+CUIT del cliente contra la base del distribuidor (tabla <c>clientes</c>).
        /// Si es válido, carga los descuentos asignados (<c>cldes1</c>, <c>cldes2</c>) en las variables
        /// globales de Module y determina si el cliente está autorizado o restringido
        /// (ambos descuentos = 0 indica que el distribuidor aún no le asignó precios).
        /// </summary>
        /// <param name="codi">Código numérico del cliente.</param>
        /// <param name="cuit">CUIT del cliente (se convierte internamente a <c>double</c> para la consulta).</param>
        /// <returns>
        /// <c>true</c> si el par código+CUIT es válido en la base del distribuidor;
        /// <c>false</c> si no se encuentra o ante cualquier error de base de datos.
        /// </returns>
        public static bool ValidarCuitCliente(string codi, string cuit)
        {
            OleDbCommand cmaccess = new OleDbCommand();
            OleDbDataReader draccess = null;
            try
            {
                Module.ConectarDBArti();
                cmaccess.Connection = Module.cnn;
                cmaccess.CommandType = CommandType.Text;
                cmaccess.CommandText = "SELECT clcodi, clrazo, clciva, clvend, clcpag, cldes1, cldes2 FROM clientes WHERE clcodi=@codi AND clcuit=@cuit";

                // Convertir los parámetros string a sus tipos nativos para la consulta parametrizada
                // (evita SQL injection y garantiza la comparación correcta en la base de datos)
                if (!int.TryParse((codi ?? string.Empty).Trim(), out int cod))
                    cod = 0;
                if (!double.TryParse((cuit ?? string.Empty).Trim(), out double cuitNum))
                    cuitNum = 0;

                cmaccess.Parameters.Clear();
                cmaccess.Parameters.AddWithValue("@codi", cod);
                cmaccess.Parameters.AddWithValue("@cuit", cuitNum);
                draccess = cmaccess.ExecuteReader();

                // Par código+CUIT válido: cargar todos los datos y descuentos del cliente
                if (draccess.Read())
                {
                    Module.US_CLRAZO = draccess["clrazo"].ToString();
                    Module.US_CLCIVA = !string.IsNullOrEmpty(draccess["clciva"].ToString()) ? Convert.ToInt32(draccess["clciva"]) : 0;
                    Module.US_CLVEND = !string.IsNullOrEmpty(draccess["clvend"].ToString()) ? Convert.ToInt32(draccess["clvend"]) : 0;
                    Module.US_CLCPAG = !string.IsNullOrEmpty(draccess["clcpag"].ToString()) ? Convert.ToInt32(draccess["clcpag"]) : 0;

                    Module.ClienteDescuento1 = draccess["cldes1"] != DBNull.Value ? Convert.ToDecimal(draccess["cldes1"]) : 0;
                    Module.ClienteDescuento2 = draccess["cldes2"] != DBNull.Value ? Convert.ToDecimal(draccess["cldes2"]) : 0;
                    // Un cliente con ambos descuentos en 0 está "restringido": puede ver el catálogo
                    // pero el distribuidor aún no le asignó precios de venta
                    Module.ClienteRestringido = Module.ClienteDescuento1 == 0 && Module.ClienteDescuento2 == 0;
                    Module.ClienteAutorizado = !Module.ClienteRestringido;
                    Module.ClienteMensajeValidacion = Module.ClienteRestringido
                        ? "Notifique a la Distribuidora que no tiene descuentos aplicados"
                        : string.Empty;

                    Module.ExisteCuit_Cliente = true;
                }
                else
                {
                    // Par inválido: limpiar todas las variables globales del cliente
                    Module.US_CLRAZO = string.Empty;
                    Module.US_CLCIVA = 0;
                    Module.US_CLVEND = 0;
                    Module.US_CLCPAG = 0;
                    Module.ExisteCuit_Cliente = false;

                    Module.ClienteDescuento1 = 0;
                    Module.ClienteDescuento2 = 0;
                    Module.ClienteRestringido = false;
                    Module.ClienteAutorizado = false;
                    Module.ClienteMensajeValidacion = "Datos inválidos";
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


        /// <summary>
        /// Registra la instancia activa de <c>frmDatosUsuario</c> en este módulo estático.
        /// Debe llamarse desde el constructor del formulario para que todos los métodos
        /// de <c>ModuloUsuario</c> puedan acceder y actualizar sus controles directamente.
        /// </summary>
        /// <param name="form">Instancia del formulario de datos de usuario recién creada.</param>
        public static void SetDatosUsuarioForm(frmDatosUsuario form)
        {
            _frmDatosUsuario = form;
        }

        /// <summary>
        /// Inserta un nuevo registro en la tabla <c>USUARIO</c> de la base local
        /// (primera vez que el cliente se registra en la aplicación).
        /// </summary>
        /// <remarks>
        /// <b>Importante:</b> los parámetros <paramref name="p4"/> y <paramref name="p5"/> (descuentos)
        /// son recibidos por firma de método pero <b>no se usan</b>. Los descuentos reales se obtienen
        /// en tiempo real de la base del distribuidor mediante <see cref="ValidarCuitCliente"/>,
        /// garantizando que siempre sean los valores asignados por el distribuidor.
        /// </remarks>
        /// <param name="p1">Código de cliente (US_CODI).</param>
        /// <param name="p2">CUIT del cliente (US_CUIT).</param>
        /// <param name="p3">Email de contacto (US_MAIL).</param>
        /// <param name="p4">Descuento 1 — <i>ignorado</i>; se consulta a la base del distribuidor.</param>
        /// <param name="p5">Descuento 2 — <i>ignorado</i>; se consulta a la base del distribuidor.</param>
        /// <param name="p6">Razón social del cliente (US_CLRAZO).</param>
        /// <param name="p7">Código de condición IVA (US_CLCIVA).</param>
        /// <param name="p8">Código de vendedor asignado (US_CLVEND).</param>
        /// <param name="p9">Código de condición de pago (US_CLCPAG).</param>
        /// <param name="bene">Beneficio adicional del cliente (US_BENEFICIO).</param>
        public static void GrabarUsuario(string p1, string p2, string p3, decimal p4, decimal p5, string p6, string p7, string p8, string p9, string bene)
        {
            using var cmd = new OleDbCommand();
            try
            {
                // Los descuentos siempre provienen de la base del distribuidor, no del parámetro.
                // p4/p5 son aceptados en la firma por compatibilidad pero sus valores son descartados.
                decimal desc1 = 0;
                decimal desc2 = 0;
                if (ValidarCuitCliente(p1, p2))
                {
                    desc1 = Module.ClienteDescuento1;
                    desc2 = Module.ClienteDescuento2;
                }

                Module.ConectarDBUsuario();
                cmd.Connection = Module.cnnU;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = @"INSERT INTO USUARIO 
                    (US_CODI, US_CUIT, US_MAIL, US_DESCUENTO1, US_DESCUENTO2, US_CLRAZO, US_CLCIVA, US_CLVEND, US_CLCPAG, US_BENEFICIO) 
                    VALUES (@codi, @cuit, @mail, @desc1, @desc2, @clrazo, @clciva, @clvend, @clcpag, @bene)";

                // Mapeo de parámetros → columnas USUARIO:
                // p1=US_CODI, p2=US_CUIT, p3=US_MAIL, desc1/2=de la base del distribuidor,
                // p6=US_CLRAZO, p7=US_CLCIVA, p8=US_CLVEND, p9=US_CLCPAG, bene=US_BENEFICIO
                cmd.Parameters.AddWithValue("@codi", p1);
                cmd.Parameters.AddWithValue("@cuit", p2);
                cmd.Parameters.AddWithValue("@mail", p3);
                cmd.Parameters.AddWithValue("@desc1", desc1);
                cmd.Parameters.AddWithValue("@desc2", desc2);
                cmd.Parameters.AddWithValue("@clrazo", p6);
                cmd.Parameters.AddWithValue("@clciva", p7);
                cmd.Parameters.AddWithValue("@clvend", p8);
                cmd.Parameters.AddWithValue("@clcpag", p9);
                cmd.Parameters.AddWithValue("@bene", bene);

                cmd.ExecuteNonQuery();
                Module.US_CODI = p1; // actualizar la variable global con el código del cliente registrado

                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = true;  // ícono verde: INSERT exitoso
                _frmDatosUsuario.ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = false;
                _frmDatosUsuario.ptbImgNo.Visible = true;  // ícono rojo: error en el INSERT
                throw new Exception($"Error en GrabarUsuario: {ex.Message}");
            }
        }

        /// <summary>
        /// Verifica si existe al menos un registro en la tabla <c>USUARIO</c> de la base local.
        /// </summary>
        /// <remarks>
        /// Se usa al abrir <c>frmDatosUsuario</c> para determinar el flujo de acción:
        /// si retorna <c>true</c> → modo edición (UPDATE); si retorna <c>false</c> → modo alta (INSERT).
        /// Esto se refleja en <c>Module.accionDatosUsu</c>: 1=UPDATE, 0=INSERT.
        /// </remarks>
        /// <returns>
        /// <c>true</c> si hay un registro en USUARIO; <c>false</c> si la tabla está vacía o ante cualquier error.
        /// </returns>
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

        /// <summary>
        /// Actualiza el registro existente en la tabla <c>USUARIO</c> de la base local.
        /// </summary>
        /// <remarks>
        /// Idéntico a <see cref="GrabarUsuario"/> en comportamiento: los descuentos se obtienen
        /// de la base del distribuidor (<see cref="ValidarCuitCliente"/>), ignorando
        /// los parámetros <paramref name="p4"/> y <paramref name="p5"/>.
        /// </remarks>
        /// <param name="p1">Código de cliente (US_CODI).</param>
        /// <param name="p2">CUIT del cliente (US_CUIT).</param>
        /// <param name="p3">Email de contacto (US_MAIL).</param>
        /// <param name="p4">Descuento 1 — <i>ignorado</i>; se consulta a la base del distribuidor.</param>
        /// <param name="p5">Descuento 2 — <i>ignorado</i>; se consulta a la base del distribuidor.</param>
        /// <param name="p6">Razón social del cliente (US_CLRAZO).</param>
        /// <param name="p7">Código de condición IVA (US_CLCIVA).</param>
        /// <param name="p8">Código de vendedor asignado (US_CLVEND).</param>
        /// <param name="p9">Código de condición de pago (US_CLCPAG).</param>
        /// <param name="bene">Beneficio adicional del cliente (US_BENEFICIO).</param>
        public static void UpdateUsuario(string p1, string p2, string p3, decimal p4, decimal p5, string p6, string p7, string p8, string p9, string bene)
        {
            using var cmd = new OleDbCommand();
            try
            {
                // Los descuentos siempre provienen de la base del distribuidor, no del parámetro.
                decimal desc1 = 0;
                decimal desc2 = 0;
                if (ValidarCuitCliente(p1, p2))
                {
                    desc1 = Module.ClienteDescuento1;
                    desc2 = Module.ClienteDescuento2;
                }

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
                cmd.Parameters.AddWithValue("@desc1", desc1);
                cmd.Parameters.AddWithValue("@desc2", desc2);
                cmd.Parameters.AddWithValue("@usuario", p6);
                cmd.Parameters.AddWithValue("@clciva", p7);
                cmd.Parameters.AddWithValue("@clvend", p8);
                cmd.Parameters.AddWithValue("@clcpag", p9);
                cmd.Parameters.AddWithValue("@bene", bene);

                cmd.ExecuteNonQuery();
                Module.US_CODI = p1; // actualizar variable global con el código del cliente

                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = true;  // ícono verde: UPDATE exitoso
                _frmDatosUsuario.ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _frmDatosUsuario.ptbImgSi.Visible = false;
                _frmDatosUsuario.ptbImgNo.Visible = true;  // ícono rojo: error en el UPDATE
                throw new Exception($"Error en UpdateUsuario: {ex.Message}");
            }
        }

        /// <summary>
        /// Lee el registro de la tabla <c>USUARIO</c> en la base local y puebla los controles
        /// de <c>frmDatosUsuario</c> con el código, CUIT, email, descuentos y beneficio guardados.
        /// Al finalizar llama a <see cref="ValidarCodiCliente"/> para sincronizar las variables
        /// globales de Module con los datos del cliente cargado.
        /// </summary>
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
                    // Sincronizar variables globales de Module (US_CLRAZO, ExisteCodi_Cliente, etc.)
                    // con los datos del cliente que se acaba de cargar en el formulario
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

        /// <summary>
        /// Carga las imágenes de estado visual en los PictureBox del formulario:
        /// <c>si.jpg</c> → <c>ptbImgSi</c> (ícono verde de operación exitosa) y
        /// <c>no.jpg</c> → <c>ptbImgNo</c> (ícono rojo de error o datos incompletos).
        /// Las imágenes se buscan en la subcarpeta <c>imagenes</c> del directorio de la aplicación.
        /// Si algún archivo no existe, el PictureBox correspondiente queda sin imagen (no lanza error).
        /// </summary>
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

        /// <summary>
        /// Verifica que los campos obligatorios del formulario (código y CUIT) no estén vacíos.
        /// Según el resultado, habilita o deshabilita <c>btnGrabar</c> y el campo email,
        /// y actualiza el ícono visual (verde = datos completos / rojo = faltan datos).
        /// </summary>
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

        /// <summary>
        /// Valida que los campos de descuento (Desc1 y Desc2) sean numéricos y estén en el rango 0–100.
        /// Si son válidos, habilita <c>btnGrabar</c> y muestra el ícono verde;
        /// si no, lo deshabilita y muestra el ícono rojo.
        /// </summary>
        /// <remarks>
        /// En el flujo habitual, los descuentos son de solo lectura (provienen de la base del distribuidor),
        /// por lo que este método actúa como guardia de seguridad ante cualquier modificación inesperada.
        /// </remarks>
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

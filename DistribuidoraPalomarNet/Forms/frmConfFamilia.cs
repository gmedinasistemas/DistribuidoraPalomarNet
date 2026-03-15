using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    public partial class frmConfFamilia : Form
    {
        private static bool _familiasInicializadas = false;
        private Label _lblCargando;

        // Intenta convertir una cadena a decimal aceptando tanto '.' como ',' como separador decimal,
        // independientemente de la cultura del sistema. Útil para campos editados a mano por el usuario.
        private static bool TryParseDecimalFlexible(string input, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(input))
                return false;

            var s = input.Trim();

            if (decimal.TryParse(s, NumberStyles.Number, CultureInfo.CurrentCulture, out value))
                return true;
            if (decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out value))
                return true;

            var swapped = s.Contains(',')
                ? s.Replace(",", ".")
                : s.Replace(".", ",");

            if (decimal.TryParse(swapped, NumberStyles.Number, CultureInfo.CurrentCulture, out value))
                return true;
            if (decimal.TryParse(swapped, NumberStyles.Number, CultureInfo.InvariantCulture, out value))
                return true;

            return false;
        }

        // Filtro de teclado reutilizable para TextBox numéricos.
        // Permite: dígitos, teclas de control (Backspace, Delete) y un único separador decimal ('.' o ',').
        // Bloquea cualquier otro carácter.
        private static void NumericTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar) || char.IsDigit(e.KeyChar))
                return;

            if (e.KeyChar == '.' || e.KeyChar == ',')
            {
                if (sender is TextBox tb)
                {
                    if (tb.Text.Contains(",") || tb.Text.Contains("."))
                    {
                        e.Handled = true;
                        return;
                    }
                }
                return;
            }

            e.Handled = true;
        }

        // Al entrar en modo edición en la grilla, aplica el filtro numérico solo a las columnas
        // que deben contener valores decimales: FADES1 (descuento), Iva y Beneficio.
        // Se desuscribe primero para evitar suscripciones duplicadas en cada edición.
        private void DataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is TextBox tb)
            {
                tb.KeyPress -= NumericTextBox_KeyPress;

                string colName = DataGridView1.CurrentCell?.OwningColumn?.Name;
                if (string.Equals(colName, "FADES1", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(colName, "Iva", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(colName, "Beneficio", StringComparison.OrdinalIgnoreCase))
                {
                    tb.KeyPress += NumericTextBox_KeyPress;
                }
            }
        }

        // Al perder el foco, formatea el valor del TextBox a dos decimales (N2) si es numérico válido.
        // Normaliza la visualización independientemente de cómo el usuario ingresó el número.
        private void NumericTextBox_Leave(object sender, EventArgs e)
        {
            if (sender is TextBox tb)
            {
                if (TryParseDecimalFlexible(tb.Text, out var v))
                    tb.Text = v.ToString("N2", CultureInfo.CurrentCulture);
            }
        }

        // Inicializa los controles del formulario, suscribe todos los eventos de botones, checkboxes
        // y grilla, y aplica el filtro numérico a los campos de entrada (txtbgral, txtiva, txtDescFlia).
        public frmConfFamilia()
        {
            InitializeComponent();
            // Wire events similares al patrón del proyecto
            Load += frmConfFamilia_Load;
            Shown += frmConfFamilia_Shown;
            btnGrabar.Click += btnGrabar_Click;
            btnSalir.Click += btnSalir_Click;
            CheckBox1.CheckedChanged += CheckBox1_CheckedChanged;
            btnRestaurar.Click += btnRestaurar_Click;

            DataGridView1.EditingControlShowing += DataGridView1_EditingControlShowing;
            txtbgral.KeyPress += NumericTextBox_KeyPress;
            txtiva.KeyPress += NumericTextBox_KeyPress;
            txtDescFlia.KeyPress += NumericTextBox_KeyPress;
            txtbgral.Leave += NumericTextBox_Leave;
            txtiva.Leave += NumericTextBox_Leave;
            txtDescFlia.Leave += NumericTextBox_Leave;
        }

        // Evento Load: intencionalmente vacío. La lógica de inicialización se difiere al evento Shown
        // para que el formulario se renderice antes de ejecutar las consultas a la base de datos.
        private void frmConfFamilia_Load(object sender, EventArgs e)
        {
            // La carga pesada se realiza en el evento Shown para que el formulario se pinte primero
        }

        // Al mostrarse el formulario: ajusta tamaño y posición en pantalla, muestra un indicador de carga
        // y ejecuta la sincronización de familias (Inicializacion) solo la primera vez que se abre.
        // Siempre recarga la grilla y la configura al final.
        private void frmConfFamilia_Shown(object sender, EventArgs e)
        {
            Cursor previousCursor = Cursor.Current;
            string previousText = this.Text;
            try
            {
                AutoScroll = true;
                var wa = Screen.FromControl(this).WorkingArea;
                if (Width > wa.Width || Height > wa.Height)
                {
                    Size = new Size(Math.Min(Width, wa.Width), Math.Min(Height, wa.Height));
                }
                Location = new Point(
                    wa.Left + Math.Max(0, (wa.Width - Width) / 2),
                    wa.Top + Math.Max(0, (wa.Height - Height) / 2));

                // Mostrar indicador visual mientras se inicializa
                Cursor.Current = Cursors.WaitCursor;
                this.Text = string.IsNullOrEmpty(previousText)
                    ? "Cargando familias..."
                    : previousText + " - Cargando familias...";
                // Crear label de estado si aún no existe
                if (_lblCargando == null)
                {
                    _lblCargando = new Label();
                    _lblCargando.Text = "Cargando familias...";
                    _lblCargando.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
                    _lblCargando.ForeColor = System.Drawing.Color.DarkSlateBlue;
                    _lblCargando.BackColor = System.Drawing.Color.Transparent;
                    _lblCargando.Dock = DockStyle.Fill;
                    _lblCargando.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                    this.Controls.Add(_lblCargando);
                }
                _lblCargando.BringToFront();
                _lblCargando.Visible = true;
                this.Refresh();

                // Ejecutar Inicializacion solo la primera vez (usa conexiones compartidas en Module)
                if (!_familiasInicializadas)
                {
                    Inicializacion();
                    _familiasInicializadas = true;
                }

                CargarGrilla();
                ConfigurarGrilla();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (_lblCargando != null)
                {
                    _lblCargando.Visible = false;
                }
                Cursor.Current = previousCursor;
                this.Text = previousText;
            }
        }

        // Configura la apariencia y comportamiento del DataGridView:
        // impide agregar/eliminar filas, define selección por fila completa, activa edición al entrar,
        // y formatea las columnas numéricas (Descuento, Iva, Beneficio) con dos decimales.
        private void ConfigurarGrilla()
        {
            DataGridView1.AutoGenerateColumns = true;
            DataGridView1.AllowUserToAddRows = false;
            DataGridView1.AllowUserToDeleteRows = false;
            DataGridView1.ReadOnly = false;
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridView1.MultiSelect = false;
            DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;

            if (DataGridView1.Columns.Contains("Familia"))
            {
                DataGridView1.Columns["Familia"].ReadOnly = true;
                DataGridView1.Columns["Familia"].Width = 200;
            }
            if (DataGridView1.Columns.Contains("FADES1"))
            {
                DataGridView1.Columns["FADES1"].HeaderText = "Descuento Familia";
                DataGridView1.Columns["FADES1"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridView1.Columns["FADES1"].DefaultCellStyle.Format = "N2";
            }
            if (DataGridView1.Columns.Contains("Iva"))
            {
                DataGridView1.Columns["Iva"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridView1.Columns["Iva"].DefaultCellStyle.Format = "N2";
            }
            if (DataGridView1.Columns.Contains("Beneficio"))
            {
                DataGridView1.Columns["Beneficio"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                DataGridView1.Columns["Beneficio"].DefaultCellStyle.Format = "N2";
            }

        }

        // Habilita o deshabilita los campos de carga general (txtbgral, txtiva, txtDescFlia)
        // según el estado del CheckBox. Cuando está tildado se puede ingresar un valor único
        // para aplicar a todas las familias a la vez.
        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool enabled = CheckBox1.Checked;
            lblbgral.Enabled = enabled;
            txtbgral.Enabled = enabled;
            lbliva.Enabled = enabled;
            txtiva.Enabled = enabled;
            lblDescFlia.Enabled = enabled;
            txtDescFlia.Enabled = enabled;
        }

        // Graba los descuentos según el modo activo:
        // - CheckBox1 desmarcado: recorre la grilla fila por fila y actualiza FABEN1, FADES1 y FAIVA por familia.
        // - CheckBox1 marcado: aplica los valores generales (txtbgral/txtiva/txtDescFlia) a TODAS las familias.
        // Al finalizar recarga la grilla para reflejar los valores guardados.
        private void btnGrabar_Click(object sender, EventArgs e)
        {
            try
            {
                btnGrabar.Enabled = false;
                if (CheckBox1.Checked == false)
                {
                    // Carga por grilla: recorre filas y actualiza FABEN1 y FAIVA por familia
                    if (DataGridView1.RowCount > 1)
                    {
                        //destilo la carga general.
                        CheckBox1_CheckedChanged(CheckBox1, EventArgs.Empty);
                        int s = 0;
                        while (s < DataGridView1.Rows.Count - 1)
                        {
                            Application.DoEvents();
                            var row = DataGridView1.Rows[s];
                            string familia = Convert.ToString(row.Cells[0].Value);
                            string ben1 = Convert.ToString(row.Cells[3].Value);
                            string fades1 = Convert.ToString(row.Cells[1].Value);
                            string iva = Convert.ToString(row.Cells[2].Value);

                            TryParseDecimalFlexible((ben1 ?? string.Empty).Trim(), out var ben1Val);
                            TryParseDecimalFlexible((fades1 ?? string.Empty).Trim(), out var fades1Val);
                            TryParseDecimalFlexible((iva ?? string.Empty).Trim(), out var ivaVal);

                            UpdateFamilias(familia, ben1Val, fades1Val, ivaVal);
                            s = s + 1;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No hay registros en la grilla", "Sin Datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // Carga general
                    CargaGral();
                }

                CargarGrilla();
                btnGrabar.Enabled = true;
            }
            catch (Exception ex)
            {
                btnGrabar.Enabled = true;
                MessageBox.Show($"Error dentro de 'btnGrabar_Click': {ex.Message}", "Inconveniente durante accion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        // Cierra el formulario. Si hubo cambios guardados (ptbImgSi visible):
        // 1. Muestra una barra de progreso mientras recalcula.
        // 2. Aplica los descuentos del cliente (aplicar_desc_Cliente).
        // 3. Recalcula precios generales y de ofertas (Update_Precios / Update_Precios_Ofertas).
        // 4. Reaplica el filtro activo en frmPrincipal intentando mantener la fila seleccionada.
        private void btnSalir_Click(object sender, EventArgs e)
        {
            Form frmProgreso = null;
            Label lblProgreso = null;
            ProgressBar pbProgreso = null;

            try
            {
                if (ptbImgSi.Visible)
                {
                    Cursor previousCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    frmProgreso = new Form();
                    frmProgreso.StartPosition = FormStartPosition.CenterScreen;
                    frmProgreso.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frmProgreso.ControlBox = false;
                    frmProgreso.MinimizeBox = false;
                    frmProgreso.MaximizeBox = false;
                    frmProgreso.ShowInTaskbar = false;
                    frmProgreso.TopMost = true;
                    frmProgreso.Width = 520;
                    frmProgreso.Height = 160;
                    frmProgreso.Text = "Cerrando configuracion";

                    lblProgreso = new Label();
                    lblProgreso.AutoSize = false;
                    lblProgreso.Left = 16;
                    lblProgreso.Top = 16;
                    lblProgreso.Width = frmProgreso.ClientSize.Width - 32;
                    lblProgreso.Height = 50;
                    lblProgreso.Text = "Aplicando cambios...";

                    pbProgreso = new ProgressBar();
                    pbProgreso.Left = 16;
                    pbProgreso.Top = 78;
                    pbProgreso.Width = frmProgreso.ClientSize.Width - 32;
                    pbProgreso.Height = 22;
                    pbProgreso.Style = ProgressBarStyle.Marquee;
                    pbProgreso.MarqueeAnimationSpeed = 30;

                    frmProgreso.Controls.Add(lblProgreso);
                    frmProgreso.Controls.Add(pbProgreso);
                    frmProgreso.Show();
                    Application.DoEvents();

                    lblProgreso.Text = "Aplicando descuentos del cliente...";
                    Application.DoEvents();
                    Module.aplicar_desc_Cliente(); // Cargo Descuentos 1 y 2

                    lblProgreso.Text = "Recalculando precios...";
                    Application.DoEvents();
                    Module.Update_Precios();

                    lblProgreso.Text = "Recalculando precios de ofertas...";
                    Application.DoEvents();
                    Module.Update_Precios_Ofertas();

                    lblProgreso.Text = "Actualizando catalogo...";
                    Application.DoEvents();

                    // Re-aplicar el filtro activo (si hay) o hacer carga completa
                    var fp = Application.OpenForms.OfType<frmPrincipal>().FirstOrDefault();

                    if (fp != null && fp.DataGridView1 != null)
                    {
                        int prevRowIndex = 0;
                        try
                        {
                            prevRowIndex = fp.DataGridView1.CurrentCell?.RowIndex ?? 0;
                        }
                        catch
                        {
                            prevRowIndex = 0;
                        }

                        bool hayFilasConDatos = fp.DataGridView1.Rows
                            .Cast<DataGridViewRow>()
                            .Any(r => !r.IsNewRow);

                        if (hayFilasConDatos)
                        {
                            fp.RefrescarSegunFiltros();

                            try
                            {
                                if (fp.DataGridView1.Rows.Count > 0)
                                {
                                    int targetRowIndex = Math.Max(0, Math.Min(prevRowIndex, fp.DataGridView1.Rows.Count - 1));
                                    fp.DataGridView1.CurrentCell = fp.DataGridView1.Rows[targetRowIndex].Cells[0];
                                    fp.DataGridView1.Rows[targetRowIndex].Selected = true;
                                    Module.AccionGrilla(fp.DataGridView1);
                                }
                            }
                            catch
                            {
                            }
                        }
                    }

                    Cursor.Current = previousCursor;
                }
            }
            catch
            {
            }
            finally
            {
                if (frmProgreso != null)
                {
                    try
                    {
                        frmProgreso.Close();
                        frmProgreso.Dispose();
                    }
                    catch
                    {
                    }
                }
            }
            this.Close();
        }

        // Reinicia las familias desde cero: borra todos los registros de FAMILIAS en la base de usuario
        // y los recrea sincronizando desde la base de artículos. Útil si los datos quedaron inconsistentes.
        private void btnRestaurar_Click(object sender, EventArgs e)
        {
            try
            {
                btnRestaurar.Enabled = false;
                BorrarFamilias();
                Inicializacion();
                CargarGrilla();
                btnRestaurar.Enabled = true;
            }
            catch (Exception ex)
            {
                btnRestaurar.Enabled = true;
                MessageBox.Show($"Error dentro de 'btnRestaurar_Click': {ex.Message}", "Inconveniente durante accion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        // Alternativa no usada en el flujo principal: recorre la grilla y actualiza el beneficio de cada
        // familia con validación de tipo numérico. Reservada para uso o extensión futura.
        // (El flujo activo usa btnGrabar_Click con UpdateFamilias directo.)
        private void ActualizarFilasDesdeGrilla()
        {
            if (DataGridView1.RowCount <= 0) return;
            for (int i = 0; i < DataGridView1.Rows.Count; i++)
            {
                var row = DataGridView1.Rows[i];
                if (row.IsNewRow) continue;
                string familia = Convert.ToString(row.Cells["Familia"].Value);
                string bene = Convert.ToString(row.Cells["Beneficio"].Value);
                if (string.IsNullOrWhiteSpace(familia)) continue;
                if (!string.IsNullOrWhiteSpace(bene) && !float.TryParse(bene.Trim(), out _))
                    throw new Exception("El beneficio debe ser numérico");
                UpdateFamilias(familia, bene);
            }
        }

        // Consulta la tabla FAMILIAS de la base de usuario y carga los resultados en el DataGridView.
        // Columnas: Familia (solo lectura), Descuento Familia (FADES1), Iva (FAIVA), Beneficio (FABEN1).
        private void CargarGrilla()
        {
            OleDbDataAdapter da = null;
            try
            {
                Module.ConectarDBUsuario();
                da = new OleDbDataAdapter("SELECT DISTINCT FADESC as Familia, FADES1, FAIVA AS Iva, FABEN1 as Beneficio from FAMILIAS", Module.cnnU);
                var ds = new DataSet();
                da.Fill(ds);
                DataGridView1.DataSource = ds.Tables[0];
                DataGridView1.Refresh();
                DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                if (DataGridView1.Columns.Contains("Familia"))
                    DataGridView1.Columns["Familia"].ReadOnly = true;
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                throw new Exception(", dentro de la accion 'CargarGrilla': " + ex.Message);
            }
        }

        // Sincroniza la tabla FAMILIAS entre la base de artículos (cnn) y la base de usuario (cnnU):
        // 1. Carga en memoria las familias ya existentes en cnnU.
        // 2. Recorre las familias de cnn: inserta las que no están en cnnU, actualiza las que cambiaron de descripción.
        // 3. Elimina de cnnU los registros que ya no existen en cnn (DepurarFamiliasUsuario).
        // Se ejecuta una sola vez por sesión (_familiasInicializadas).
        private void Inicializacion()
        {
            OleDbCommand cmFamiliasArt = new OleDbCommand();
            OleDbDataReader drFamiliasArt = null;
            try
            {
                if (string.IsNullOrWhiteSpace(Module.US_CODI))
                    return;

                // 1) Cargar en memoria las familias existentes en la base de usuario (cnnU)
                var familiasUsuario = new Dictionary<string, string>(); // FACODI -> FADESC
                OleDbCommand cmFamiliasUsu = new OleDbCommand();
                OleDbDataReader drFamiliasUsu = null;
                try
                {
                    Module.ConectarDBUsuario();
                    cmFamiliasUsu.Connection = Module.cnnU;
                    cmFamiliasUsu.CommandType = CommandType.Text;
                    cmFamiliasUsu.CommandText = "SELECT FACODI, FADESC FROM FAMILIAS";
                    drFamiliasUsu = cmFamiliasUsu.ExecuteReader();
                    while (drFamiliasUsu.Read())
                    {
                        string facodiU = Convert.ToString(drFamiliasUsu["FACODI"]);
                        string fadescU = Convert.ToString(drFamiliasUsu["FADESC"]);
                        if (!familiasUsuario.ContainsKey(facodiU))
                            familiasUsuario.Add(facodiU, fadescU);
                        else
                            familiasUsuario[facodiU] = fadescU;
                    }
                }
                finally
                {
                    if (drFamiliasUsu != null && !drFamiliasUsu.IsClosed) drFamiliasUsu.Close();
                    if (Module.cnnU.State == ConnectionState.Open) Module.cnnU.Close();
                    cmFamiliasUsu.Dispose();
                }

                // 2) Recorrer familias de la base de articulos y sincronizar con la base de usuario
                Module.ConectarDBArti();
                cmFamiliasArt.Connection = Module.cnn;
                cmFamiliasArt.CommandType = CommandType.Text;
                cmFamiliasArt.CommandText = "SELECT DISTINCT f.facodi, f.fadesc from familias as f " +
                                            //" inner join descuent as d on d.desflia=f.facodi " +
                                            //" where d.desclie=" + Module.US_CODI +
                                            " order by f.fadesc";
                drFamiliasArt = cmFamiliasArt.ExecuteReader();
                Application.DoEvents();
                while (drFamiliasArt.Read())
                {
                    string facodi = drFamiliasArt["FACODI"].ToString();
                    string fadesc = drFamiliasArt["FADESC"].ToString();

                    if (!familiasUsuario.TryGetValue(facodi, out var descActual))
                    {
                        // No existe en la base de usuario, insertar
                        InsertarFamilia(facodi, fadesc);
                        familiasUsuario[facodi] = fadesc;
                    }
                    else if (!string.Equals(descActual, fadesc, StringComparison.Ordinal))
                    {
                        // Existe pero con descripcion distinta, actualizar
                        UpdateFamilia(facodi, fadesc);
                        familiasUsuario[facodi] = fadesc;
                    }
                }
                drFamiliasArt.Close();

                // 3) Borrar de cnnU los FACODI que ya no existan en la base del conector cnn
                DepurarFamiliasUsuario();

                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
                throw new Exception(", dentro de la funcion 'Inicializacion': " + ex.Message);
            }
            finally
            {
                cmFamiliasArt.Dispose();
            }
        }

        // Verifica si una familia existe en la base de usuario y la inserta si no está.
        // Se usa como fallback de sincronización por descripción.
        private void ActualizarFAMIVA(string facodi, string fadesc)
        {
            try
            {
                if (!FamiliaExiste(fadesc))
                {
                    InsertarFamilia(facodi, fadesc);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(", dentro de 'ActualizarFAMIVA': " + ex.Message);
            }
        }

        // Limpia la base de usuario: recorre todos los FACODI en FAMILIAS (cnnU) y elimina
        // aquellos que ya no existen en la base de artículos (cnn). Previene registros huérfanos.
        private void DepurarFamiliasUsuario()
        {
            OleDbCommand cmdUsu = new OleDbCommand();
            OleDbDataReader drUsu = null;
            try
            {
                Module.ConectarDBUsuario();
                cmdUsu.Connection = Module.cnnU;
                cmdUsu.CommandType = CommandType.Text;
                cmdUsu.CommandText = "SELECT FACODI FROM FAMILIAS";
                drUsu = cmdUsu.ExecuteReader();
                var facodisAEliminar = new System.Collections.Generic.List<string>();
                while (drUsu.Read())
                {
                    string facodiU = Convert.ToString(drUsu["FACODI"]);
                    if (!FacodiExisteEnArticulos(facodiU))
                    {
                        facodisAEliminar.Add(facodiU);
                    }
                }
                drUsu.Close();

                foreach (var facodi in facodisAEliminar)
                {
                    using (var del = new OleDbCommand("DELETE FROM FAMILIAS WHERE FACODI=" + facodi, Module.cnnU))
                    {
                        del.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(", dentro de 'DepurarFamiliasUsuario': " + ex.Message);
            }
            finally
            {
                if (drUsu != null && !drUsu.IsClosed) drUsu.Close();
                if (Module.cnnU.State == ConnectionState.Open) Module.cnnU.Close();
                cmdUsu.Dispose();
            }
        }

        // Consulta la base de artículos para verificar si un FACODI existe en la tabla FAMILIAS.
        // Retorna true si existe, false si no o si hay error (seguro para uso en depuración).
        private bool FacodiExisteEnArticulos(string facodi)
        {
            OleDbCommand cmdArt = new OleDbCommand();
            OleDbDataReader dr = null;
            try
            {
                Module.ConectarDBArti();
                cmdArt.Connection = Module.cnn;
                cmdArt.CommandType = CommandType.Text;
                cmdArt.CommandText = "SELECT facodi FROM familias WHERE facodi=" + facodi;
                dr = cmdArt.ExecuteReader();
                bool existe = dr.Read();
                dr.Close();
                return existe;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (Module.cnn.State == ConnectionState.Open) Module.cnn.Close();
                cmdArt.Dispose();
            }
        }

        // Busca una familia en la base de usuario por su descripción (FADESC).
        // Retorna true si ya existe un registro con ese nombre.
        private bool FamiliaExiste(string Familia)
        {
            OleDbCommand cmFamilias = new OleDbCommand();
            OleDbDataReader drFamilias = null;
            bool existe = false;
            try
            {
                Module.ConectarDBUsuario();
                cmFamilias.Connection = Module.cnnU;
                cmFamilias.CommandType = CommandType.Text;
                cmFamilias.CommandText = "select FADESC from FAMILIAS where FADESC='" + (Familia ?? string.Empty).Replace("'", "''") + "'";
                drFamilias = cmFamilias.ExecuteReader();
                if (drFamilias.Read()) existe = true;
                drFamilias.Close();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                return existe;
            }
            catch (Exception)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                return existe;
            }
            finally
            {
                cmFamilias.Dispose();
            }
        }

        // Inserta una nueva familia en la base de usuario con todos sus valores numéricos en 0.
        // Se llama durante la sincronización cuando una familia de artículos no existe en el usuario.
        private void InsertarFamilia(string CodFamilia, string Familia)
        {
            OleDbCommand cmFamilias = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmFamilias.Connection = Module.cnnU;
                cmFamilias.CommandType = CommandType.Text;
                cmFamilias.CommandText = "insert into FAMILIAS (FACODI,FADESC,FABEN1,FADES1,FAIVA) values ('" + (CodFamilia ?? string.Empty).Replace("'", "''") + "','" + (Familia ?? string.Empty).Replace("'", "''") + "',0,0,0)";
                Application.DoEvents();
                cmFamilias.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                throw new Exception(", dentro de la accion 'InsertarFamilia': " + ex.Message);
            }
        }

        // Sobrecarga 1: actualiza solo FABEN1 (beneficio) de una familia identificada por FADESC.
        // Muestra ptbImgSi (verde) en éxito o ptbImgNo (rojo) en error.
        private void UpdateFamilias(string familia, string ben1)
        {
            OleDbCommand cmUsu = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = "UPDATE FAMILIAS SET FABEN1=" + ben1.ToString() + " WHERE FADESC='" + familia.Replace("'", "''") + "'";
                cmUsu.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = true;
                ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = false;
                ptbImgNo.Visible = true;
                throw new Exception(",Error dentro de 'UpdateFamilias' " + ex.Message);
            }
            finally
            {
                cmUsu.Dispose();
            }
        }

        // Sobrecarga 2 (principal): actualiza FABEN1, FADES1 y FAIVA de una familia por FADESC.
        // Usa parámetros posicionales (Parameters.AddWithValue) para evitar SQL injection y manejar decimales correctamente.
        private void UpdateFamilias(string familia, decimal ben1, decimal fades1, decimal iva)
        {
            using var cmUsu = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = "UPDATE FAMILIAS SET FABEN1 = ?, FADES1 = ?, FAIVA = ? WHERE FADESC = ?";
                cmUsu.Parameters.AddWithValue("@p1", ben1);
                cmUsu.Parameters.AddWithValue("@p2", fades1);
                cmUsu.Parameters.AddWithValue("@p3", iva);
                cmUsu.Parameters.AddWithValue("@p4", familia ?? string.Empty);
                cmUsu.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = true;
                ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = false;
                ptbImgNo.Visible = true;
                throw new Exception(",Error dentro de 'UpdateFamilias' " + ex.Message);
            }
        }

        // Sobrecarga 3: ejecuta un UPDATE con un query string libre sobre FAMILIAS.
        // Reservada para casos especiales; no se usa en el flujo principal actual.
        private void UpdateFamilias(string stringQuery)
        {
            OleDbCommand cmUsu = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = stringQuery;
                Application.DoEvents();
                cmUsu.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = true;
                ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = false;
                ptbImgNo.Visible = true;
                throw new Exception(",Error dentro de 'UpdateFamilias' " + ex.Message);
            }
            finally
            {
                cmUsu.Dispose();
            }
        }

        // Aplica los valores de los campos generales (txtbgral=beneficio, txtiva=IVA, txtDescFlia=descuento)
        // a TODAS las familias de la tabla FAMILIAS en la base de usuario.
        // Si los tres campos tienen valor los actualiza en una sola llamada; si solo algunos tienen valor,
        // los actualiza de forma independiente. Valida que los valores sean numéricos antes de ejecutar.
        private void CargaGral()
        {
            string UQuery;
            try
            {
                if (DataGridView1.RowCount > 1)
                {
                    bool hasBene = !string.IsNullOrWhiteSpace(txtbgral.Text);
                    bool hasIva = !string.IsNullOrWhiteSpace(txtiva.Text);
                    bool hasDesc = !string.IsNullOrWhiteSpace(txtDescFlia.Text);

                    if (hasBene && hasIva && hasDesc)
                    {
                        if (TryParseDecimalFlexible(txtbgral.Text.Trim(), out var benVal)
                            && TryParseDecimalFlexible(txtiva.Text.Trim(), out var ivaVal)
                            && TryParseDecimalFlexible(txtDescFlia.Text.Trim(), out var descVal))
                        {
                            UpdateFamiliasGeneral(benVal, ivaVal, descVal);
                        }
                        else
                        {
                            if (!TryParseDecimalFlexible(txtbgral.Text.Trim(), out _))
                                MessageBox.Show("El beneficio debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (!TryParseDecimalFlexible(txtiva.Text.Trim(), out _))
                                MessageBox.Show("El iva debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (!TryParseDecimalFlexible(txtDescFlia.Text.Trim(), out _))
                                MessageBox.Show("El descuento debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        if (hasBene)
                        {
                            if (TryParseDecimalFlexible(txtbgral.Text.Trim(), out var benVal))
                            {
                                UpdateFamiliasGeneral(benVal, null, null);
                            }
                            else
                            {
                                MessageBox.Show("El beneficio debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        if (hasIva)
                        {
                            if (TryParseDecimalFlexible(txtiva.Text.Trim(), out var ivaVal))
                            {
                                UpdateFamiliasGeneral(null, ivaVal, null);
                            }
                            else
                            {
                                MessageBox.Show("El iva debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }

                        if (hasDesc)
                        {
                            if (TryParseDecimalFlexible(txtDescFlia.Text.Trim(), out var descVal))
                            {
                                UpdateFamiliasGeneral(null, null, descVal);
                            }
                            else
                            {
                                MessageBox.Show("El descuento debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No hay registros para actualizar", "Sin Familias", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(", error dentro de 'CargaGral': " + ex.Message);
            }
        }

        // Construye y ejecuta un UPDATE general sobre FAMILIAS sin cláusula WHERE (afecta todos los registros).
        // Los tres parámetros son nullable: solo se incluyen en el SET los campos que tengan valor.
        // Muestra ptbImgSi (verde) en éxito o ptbImgNo (rojo) en error.
        private void UpdateFamiliasGeneral(decimal? ben1, decimal? iva, decimal? desc)
        {
            using var cmUsu = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;

                var sets = new List<string>();

                if (ben1.HasValue)
                {
                    sets.Add("FABEN1 = ?");
                    cmUsu.Parameters.AddWithValue("@p_ben", ben1.Value);
                }
                if (desc.HasValue)
                {
                    sets.Add("FADES1 = ?");
                    cmUsu.Parameters.AddWithValue("@p_desc", desc.Value);
                }
                if (iva.HasValue)
                {
                    sets.Add("FAIVA = ?");
                    cmUsu.Parameters.AddWithValue("@p_iva", iva.Value);
                }

                if (sets.Count == 0)
                    return;

                cmUsu.CommandText = "UPDATE FAMILIAS SET " + string.Join(", ", sets);

                cmUsu.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = true;
                ptbImgNo.Visible = false;
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                ptbImgSi.Visible = false;
                ptbImgNo.Visible = true;
                throw new Exception(",Error dentro de 'UpdateFamiliasGeneral' " + ex.Message);
            }
        }

        // Elimina TODOS los registros de la tabla FAMILIAS en la base de usuario.
        // Se usa antes de reinicializar (btnRestaurar) para partir desde un estado limpio.
        private void BorrarFamilias()
        {
            OleDbCommand cmFamilias = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmFamilias.Connection = Module.cnnU;
                cmFamilias.CommandType = CommandType.Text;
                cmFamilias.CommandText = "delete from FAMILIAS";
                cmFamilias.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                throw new Exception(", dentro de la acción 'BorrarFamilias': " + ex.Message);
            }
        }

        // Verifica si un FACODI existe en la tabla FAMILIAS de la base de usuario, buscando por código numérico.
        // Retorna true si existe, false si no o si hay error.
        private bool IDFamiliaExiste(string idF)
        {
            OleDbCommand cmFamilias = new OleDbCommand();
            OleDbDataReader drFamilias = null;
            bool existe = false;
            try
            {
                Module.ConectarDBUsuario();
                cmFamilias.Connection = Module.cnnU;
                cmFamilias.CommandType = CommandType.Text;
                cmFamilias.CommandText = "select FACODI from FAMILIAS where FACODI=" + idF;
                Application.DoEvents();
                drFamilias = cmFamilias.ExecuteReader();
                if (drFamilias.Read()) existe = true;
                drFamilias.Close();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                return existe;
            }
            catch
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                return existe;
            }
            finally
            {
                cmFamilias.Dispose();
            }
        }

        // Actualiza la descripción (FADESC) de una familia en la base de usuario, buscando por su FACODI.
        // Se usa durante la sincronización cuando la descripción en la base de artículos cambió.
        private void UpdateFamilia(string CodFamilia, string Familia)
        {
            OleDbCommand cmFamilias = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmFamilias.Connection = Module.cnnU;
                cmFamilias.CommandType = CommandType.Text;
                cmFamilias.CommandText = "update FAMILIAS set FADESC='" + (Familia ?? string.Empty).Replace("'", "''") + "' where FACODI=" + CodFamilia;
                Application.DoEvents();
                cmFamilias.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
                throw new Exception(", dentro de la acción 'UpdateFamilia': " + ex.Message);
            }
        }
    }
}

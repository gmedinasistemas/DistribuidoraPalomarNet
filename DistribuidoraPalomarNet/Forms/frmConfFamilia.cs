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

        // Separador de decimales fijo para parseo sin ambigüedad.
        private static readonly CultureInfo _esAR = CultureInfo.GetCultureInfo("es-AR");

        // Intenta convertir una cadena a decimal aceptando tanto '.' como ',' como separador decimal,
        // independientemente de la cultura del sistema.
        // Usa NumberStyles SIN AllowThousands para evitar que "10.5" sea interpretado como 105
        // en culturas donde '.' es separador de miles (es-AR, es-ES, etc.), y que "10,5" sea
        // interpretado como 105 en culturas donde ',' es separador de miles (en-US, etc.).
        private static bool TryParseDecimalFlexible(string input, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(input))
                return false;

            var s = input.Trim();

            // Sin AllowThousands: el separador sólo puede ser decimal, nunca de miles.
            const NumberStyles noThousands =
                NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite |
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint;

            // 1) InvariantCulture: punto como decimal  → "10.5" → 10,5
            if (decimal.TryParse(s, noThousands, CultureInfo.InvariantCulture, out value)) return true;

            // 2) es-AR: coma como decimal              → "10,5" → 10,5
            if (decimal.TryParse(s, noThousands, _esAR, out value)) return true;

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

        // Al terminar de editar una celda numérica en la grilla, convierte el valor ingresado
        // (tanto "10.5" como "10,5") al tipo del DataTable (Double desde Access) para que
        // DataGridView pueda almacenarlo sin errores y lo muestre formateado con N2.
        private void DataGridView1_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0) return;
            string colName = DataGridView1.Columns[e.ColumnIndex].Name;
            if (!string.Equals(colName, "FADES1", StringComparison.OrdinalIgnoreCase)
             && !string.Equals(colName, "Iva", StringComparison.OrdinalIgnoreCase)
             && !string.Equals(colName, "Beneficio", StringComparison.OrdinalIgnoreCase))
                return;

            if (e.Value is string raw && TryParseDecimalFlexible(raw, out decimal parsed))
            {
                // Convertir al tipo que espera el DataTable (típicamente Double desde Access)
                Type targetType = e.DesiredType ?? typeof(double);
                try { e.Value = Convert.ChangeType(parsed, targetType); }
                catch { e.Value = (double)parsed; }
                e.ParsingApplied = true;
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
            // btnRestaurar.Click ya está registrado en Designer.cs

            DataGridView1.EditingControlShowing += DataGridView1_EditingControlShowing;
            DataGridView1.CellParsing += DataGridView1_CellParsing;
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
                DataGridView1.Columns["Familia"].DisplayIndex = 0;
            }
            if (DataGridView1.Columns.Contains("Marca"))
            {
                DataGridView1.Columns["Marca"].ReadOnly = true;
                DataGridView1.Columns["Marca"].Width = 120;
                DataGridView1.Columns["Marca"].DisplayIndex = 1;
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
                            string facodi = Convert.ToString(row.Cells["FACODI"].Value);
                            string ben1 = Convert.ToString(row.Cells["Beneficio"].Value);
                            string fades1 = Convert.ToString(row.Cells["FADES1"].Value);
                            string iva = Convert.ToString(row.Cells["Iva"].Value);

                            TryParseDecimalFlexible((ben1 ?? string.Empty).Trim(), out var ben1Val);
                            TryParseDecimalFlexible((fades1 ?? string.Empty).Trim(), out var fades1Val);
                            TryParseDecimalFlexible((iva ?? string.Empty).Trim(), out var ivaVal);

                            UpdateFamiliasPorCodigo(facodi, ben1Val, fades1Val, ivaVal);
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
                ConfigurarGrilla();
                btnGrabar.Enabled = true;
            }
            catch (Exception ex)
            {
                btnGrabar.Enabled = true;
                MessageBox.Show($"Error dentro de 'btnGrabar_Click': {ex.Message}", "Inconveniente durante accion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            finally
            {
                CheckBox1.Checked = false;
            }
        }

        // Cierra el formulario. Si hubo cambios guardados (ptbImgSi visible):
        // 1. Muestra frmProgresoRecalculo (WebView2) con 4 pasos animados.
        // 2. Valida que DESDESC1 > 0; si no, notifica en webView22 de frmPrincipal y sale.
        // 3. Recalcula precios generales y de ofertas.
        // 4. Reaplica el filtro activo en frmPrincipal intentando mantener la fila seleccionada.
        private async void btnSalir_Click(object sender, EventArgs e)
        {
            frmProgresoRecalculo frmProgreso = null;

            try
            {
                if (ptbImgSi.Visible)
                {
                    Cursor previousCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    frmProgreso = new frmProgresoRecalculo();
                    frmProgreso.Show();
                    frmProgreso.WaitReady();

                    // Paso 0: Aplicar descuentos del cliente
                    frmProgreso.SetStep(0, "active");
                    frmProgreso.SetMensaje("Aplicando descuentos del cliente...");
                    Module.aplicar_desc_Cliente();
                    frmProgreso.SetStep(0, "done");

                    if (!Module.DescuentosConfigurados())
                    {
                        frmProgreso.SetError("Descuentos no configurados (DESDESC1 = 0)");
                        System.Threading.Thread.Sleep(1200);
                        frmProgreso.Close();
                        frmPrincipal.Instance?.MostrarNotificacionDescuento();
                        MessageBox.Show(
                            "Descuentos no configurados (DESDESC1 = 0).\nPida sus descuentos a la distribuidora.",
                            "Sin Descuentos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Cursor.Current = previousCursor;
                        this.Close();
                        return;
                    }

                    // Paso 1: Recalcular precios generales
                    frmProgreso.SetStep(1, "active");
                    frmProgreso.SetMensaje("Recalculando precios generales...");
                    Module.Update_Precios();
                    frmProgreso.SetStep(1, "done");

                    // Paso 2: Recalcular precios de ofertas
                    frmProgreso.SetStep(2, "active");
                    frmProgreso.SetMensaje("Recalculando precios de ofertas...");
                    Module.Update_Precios_Ofertas();
                    frmProgreso.SetStep(2, "done");

                    // Paso 3: Actualizar catálogo
                    frmProgreso.SetStep(3, "active");
                    frmProgreso.SetMensaje("Actualizando catálogo...");

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
                                    await Module.AccionGrilla(fp.DataGridView1);
                                }
                            }
                            catch
                            {
                            }
                        }
                    }

                    frmProgreso.SetStep(3, "done");
                    frmProgreso.SetFinalizado();
                    frmProgreso.SetMensaje("¡Proceso completado!");
                    System.Threading.Thread.Sleep(800);

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
                // PRIMER mensaje de confirmación
                DialogResult resultado = MessageBox.Show(
                    "Va a BORRAR y PERDER TODOS los Valores cargados.\n\n" +
                    "Esta acción actualizará las familias desde la base de datos principal,\n" +
                    "eliminando todos los descuentos, IVA y beneficios personalizados.\n\n" +
                    "¿DESEA SEGUIR ADELANTE?",
                    "⚠️ Confirmar Restauración de Familias",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                // Si el usuario elige NO, cancelar inmediatamente
                if (resultado == DialogResult.No)
                {
                    return;
                }

                // SEGUNDA confirmación: Si eligió SÍ, preguntar de nuevo
                DialogResult segundaConfirmacion = MessageBox.Show(
                    "¿Está seguro?",
                    "⚠️ Confirmación Final",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);

                // Si en la segunda confirmación elige NO, cancelar
                if (segundaConfirmacion != DialogResult.Yes)
                {
                    return;
                }

                // Proceder con la restauración
                btnRestaurar.Enabled = false;
                BorrarFamilias();
                Inicializacion();
                CargarGrilla();
                btnRestaurar.Enabled = true;

                MessageBox.Show(
                    "Las familias han sido restauradas exitosamente.\n\n" +
                    "Todos los valores han sido reiniciados desde la base de datos principal.",
                    "✓ Restauración Completada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
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

        // Consulta la base de artículos (cnn) la tabla proveedo y devuelve un diccionario PRCODI -> PRMARC.
        // Se usa para resolver la Marca de cada familia a partir de su FAPROV (ya disponible en cnnU).
        private Dictionary<string, string> ObtenerMarcasPorProveedor()
        {
            var dict = new Dictionary<string, string>();
            OleDbDataAdapter da = null;
            try
            {
                Module.ConectarDBArti();
                da = new OleDbDataAdapter("SELECT PRCODI, PRMARC FROM proveedo", Module.cnn);
                var ds = new DataSet();
                da.Fill(ds);
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    string prcodi = Convert.ToString(row["PRCODI"]);
                    if (string.IsNullOrEmpty(prcodi) || dict.ContainsKey(prcodi)) continue;
                    dict[prcodi] = row["PRMARC"] == DBNull.Value ? string.Empty : Convert.ToString(row["PRMARC"]);
                }
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
            }
            catch
            {
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
            }
            finally
            {
                da?.Dispose();
            }
            return dict;
        }

        // Consulta la tabla FAMILIAS de la base de usuario y carga los resultados en el DataGridView.
        // Cada fila queda identificada por FACODI (columna oculta), ya que una misma descripción de
        // familia (FADESC) puede repetirse con distinto proveedor (FAPROV).
        // Columnas visibles: Familia, Marca, Descuento Familia (FADES1), Iva (FAIVA), Beneficio (FABEN1).
        private void CargarGrilla()
        {
            OleDbDataAdapter da = null;
            try
            {
                Module.ConectarDBUsuario();
                da = new OleDbDataAdapter("SELECT FACODI, FADESC as Familia, FAPROV, FADES1, FAIVA AS Iva, FABEN1 as Beneficio from FAMILIAS order by FADESC", Module.cnnU);
                var ds = new DataSet();
                da.Fill(ds);

                // La Marca (PRMARC) se resuelve por FAPROV (ya presente en cnnU) contra proveedo (cnn).
                var marcasPorProveedor = ObtenerMarcasPorProveedor();
                ds.Tables[0].Columns.Add("Marca", typeof(string));
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    string faprov = row["FAPROV"] == DBNull.Value ? string.Empty : Convert.ToString(row["FAPROV"]);
                    row["Marca"] = (!string.IsNullOrEmpty(faprov) && marcasPorProveedor.TryGetValue(faprov, out var marca))
                        ? marca
                        : string.Empty;
                }

                DataGridView1.DataSource = ds.Tables[0];
                DataGridView1.Refresh();
                DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                if (DataGridView1.Columns.Contains("Familia"))
                    DataGridView1.Columns["Familia"].ReadOnly = true;
                if (DataGridView1.Columns.Contains("FACODI"))
                    DataGridView1.Columns["FACODI"].Visible = false;
                if (DataGridView1.Columns.Contains("FAPROV"))
                    DataGridView1.Columns["FAPROV"].Visible = false;
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
                var familiasUsuario = new Dictionary<string, (string FADESC, string FAPROV)>(); // FACODI -> (FADESC, FAPROV)
                OleDbCommand cmFamiliasUsu = new OleDbCommand();
                OleDbDataReader drFamiliasUsu = null;
                try
                {
                    Module.ConectarDBUsuario();
                    cmFamiliasUsu.Connection = Module.cnnU;
                    cmFamiliasUsu.CommandType = CommandType.Text;
                    cmFamiliasUsu.CommandText = "SELECT FACODI, FADESC, FAPROV FROM FAMILIAS";
                    drFamiliasUsu = cmFamiliasUsu.ExecuteReader();
                    while (drFamiliasUsu.Read())
                    {
                        string facodiU = Convert.ToString(drFamiliasUsu["FACODI"]);
                        string fadescU = Convert.ToString(drFamiliasUsu["FADESC"]);
                        string faprovU = drFamiliasUsu["FAPROV"] == DBNull.Value ? string.Empty : Convert.ToString(drFamiliasUsu["FAPROV"]);
                        familiasUsuario[facodiU] = (fadescU, faprovU);
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
                cmFamiliasArt.CommandText = "SELECT DISTINCT f.facodi, f.fadesc, f.faprov from familias as f " +
                                            //" inner join descuent as d on d.desflia=f.facodi " +
                                            //" where d.desclie=" + Module.US_CODI +
                                            " order by f.fadesc";
                drFamiliasArt = cmFamiliasArt.ExecuteReader();
                Application.DoEvents();
                while (drFamiliasArt.Read())
                {
                    string facodi = drFamiliasArt["FACODI"].ToString();
                    string fadesc = drFamiliasArt["FADESC"].ToString();
                    string faprov = drFamiliasArt["FAPROV"] == DBNull.Value ? string.Empty : drFamiliasArt["FAPROV"].ToString();

                    if (!familiasUsuario.TryGetValue(facodi, out var datosActuales))
                    {
                        // No existe en la base de usuario, insertar
                        InsertarFamilia(facodi, fadesc, faprov);
                        familiasUsuario[facodi] = (fadesc, faprov);
                    }
                    else if (!string.Equals(datosActuales.FADESC, fadesc, StringComparison.Ordinal)
                             || !string.Equals(datosActuales.FAPROV, faprov, StringComparison.Ordinal))
                    {
                        // Existe pero con descripcion o proveedor distinto, actualizar
                        UpdateFamilia(facodi, fadesc, faprov);
                        familiasUsuario[facodi] = (fadesc, faprov);
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
        // Se usa como fallback de sincronización por código de familia (FACODI).
        private void ActualizarFAMIVA(string facodi, string fadesc, string faprov = "")
        {
            try
            {
                if (!IDFamiliaExiste(facodi))
                {
                    InsertarFamilia(facodi, fadesc, faprov);
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
        // NOTA: puede haber varias familias con la misma descripción (distinto FAPROV);
        // para identificar una fila específica usar IDFamiliaExiste (FACODI).
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
        private void InsertarFamilia(string CodFamilia, string Familia, string faprov = "")
        {
            OleDbCommand cmFamilias = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmFamilias.Connection = Module.cnnU;
                cmFamilias.CommandType = CommandType.Text;
                string faprovVal = string.IsNullOrWhiteSpace(faprov) ? "NULL" : faprov;
                cmFamilias.CommandText = "insert into FAMILIAS (FACODI,FADESC,FABEN1,FADES1,FAIVA,FAPROV) values ('" + (CodFamilia ?? string.Empty).Replace("'", "''") + "','" + (Familia ?? string.Empty).Replace("'", "''") + "',0,0,0," + faprovVal + ")";
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
                cmUsu.CommandText = "UPDATE FAMILIAS SET FABEN1=" + ben1.ToString().Replace(".", ",") + " WHERE FADESC='" + familia.Replace("'", "''") + "'";
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

        // Sobrecarga 2 (legacy, no usada en el flujo activo): actualiza FABEN1, FADES1 y FAIVA de una
        // familia por FADESC. Se mantiene solo por compatibilidad con ActualizarFilasDesdeGrilla.
        // ADVERTENCIA: puede afectar más de una fila si hay familias con igual FADESC y distinto FAPROV.
        private void UpdateFamilias(string familia, decimal ben1, decimal fades1, decimal iva)
        {
            using var cmUsu = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = "UPDATE FAMILIAS SET FABEN1 = ?, FADES1 = ?, FAIVA = ? WHERE FADESC = ?";
                cmUsu.Parameters.AddWithValue("@p1", ben1.ToString().Replace(".", ","));
                cmUsu.Parameters.AddWithValue("@p2", fades1.ToString().Replace(".", ","));
                cmUsu.Parameters.AddWithValue("@p3", iva.ToString().Replace(".", ","));
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

        // Actualiza FABEN1, FADES1 y FAIVA de una familia identificada por su FACODI (clave única real).
        // Reemplaza a UpdateFamilias(string, decimal, decimal, decimal) en el flujo activo (btnGrabar_Click),
        // ya que FADESC puede repetirse entre familias con distinto proveedor (FAPROV).
        private void UpdateFamiliasPorCodigo(string facodi, decimal ben1, decimal fades1, decimal iva)
        {
            using var cmUsu = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = "UPDATE FAMILIAS SET FABEN1 = ?, FADES1 = ?, FAIVA = ? WHERE FACODI = ?";
                cmUsu.Parameters.AddWithValue("@p1", ben1.ToString().Replace(".", ","));
                cmUsu.Parameters.AddWithValue("@p2", fades1.ToString().Replace(".", ","));
                cmUsu.Parameters.AddWithValue("@p3", iva.ToString().Replace(".", ","));
                cmUsu.Parameters.AddWithValue("@p4", facodi ?? string.Empty);
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
                throw new Exception(",Error dentro de 'UpdateFamiliasPorCodigo' " + ex.Message);
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
                    cmUsu.Parameters.AddWithValue("@p_ben", ben1.Value.ToString().Replace(".", ","));
                }
                if (desc.HasValue)
                {
                    sets.Add("FADES1 = ?");
                    cmUsu.Parameters.AddWithValue("@p_desc", desc.Value.ToString().Replace(".", ","));
                }
                if (iva.HasValue)
                {
                    sets.Add("FAIVA = ?");
                    cmUsu.Parameters.AddWithValue("@p_iva", iva.Value.ToString().Replace(".", ","));
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

        // Actualiza la descripción (FADESC) y el proveedor (FAPROV) de una familia en la base de usuario,
        // buscando por su FACODI. Se usa durante la sincronización cuando alguno de esos datos cambió
        // en la base de artículos.
        private void UpdateFamilia(string CodFamilia, string Familia, string faprov = "")
        {
            OleDbCommand cmFamilias = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmFamilias.Connection = Module.cnnU;
                cmFamilias.CommandType = CommandType.Text;
                string faprovVal = string.IsNullOrWhiteSpace(faprov) ? "NULL" : faprov;
                cmFamilias.CommandText = "update FAMILIAS set FADESC='" + (Familia ?? string.Empty).Replace("'", "''") + "', FAPROV=" + faprovVal + " where FACODI=" + CodFamilia;
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

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    public partial class frmConfFamilia : Form
    {
        private static bool _familiasInicializadas = false;
        private Label _lblCargando;

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
        }

        private void frmConfFamilia_Load(object sender, EventArgs e)
        {
            // La carga pesada se realiza en el evento Shown para que el formulario se pinte primero
        }

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
                        int s = 0;
                        while (s < DataGridView1.Rows.Count - 1)
                        {
                            Application.DoEvents();
                            var row = DataGridView1.Rows[s];
                            string familia = Convert.ToString(row.Cells[0].Value);
                            string ben1 = Convert.ToString(row.Cells[3].Value);
                            string fades1 = Convert.ToString(row.Cells[1].Value);
                            string iva = Convert.ToString(row.Cells[2].Value);

                            decimal.TryParse((ben1 ?? string.Empty).Trim(), out var ben1Val);
                            decimal.TryParse((fades1 ?? string.Empty).Trim(), out var fades1Val);
                            decimal.TryParse((iva ?? string.Empty).Trim(), out var ivaVal);

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

                    // Solo llamar a Module.funcion_inicial(1) si existe el frmPrincipal
                    // y tiene datos visibles en la grilla
                    var fp = Application.OpenForms.OfType<frmPrincipal>().FirstOrDefault();

                    if (fp != null && fp.DataGridView1 != null)
                    {
                        bool hayFilasConDatos = fp.DataGridView1.Rows
                            .Cast<DataGridViewRow>()
                            .Any(r => !r.IsNewRow);

                        if (hayFilasConDatos)
                        {
                            Module.funcion_inicial(1);
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

        // Mantengo este método por si se quiere usar validación por fila, pero el flujo principal sigue el VB
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
                        if (float.TryParse(txtbgral.Text.Trim(), out _) && float.TryParse(txtiva.Text.Trim(), out _) && float.TryParse(txtDescFlia.Text.Trim(), out _))
                        {
                            decimal.TryParse(txtbgral.Text.Trim(), out var benVal);
                            decimal.TryParse(txtiva.Text.Trim(), out var ivaVal);
                            decimal.TryParse(txtDescFlia.Text.Trim(), out var descVal);
                            UpdateFamiliasGeneral(benVal, ivaVal, descVal);
                        }
                        else
                        {
                            if (!float.TryParse(txtbgral.Text.Trim(), out _))
                                MessageBox.Show("El beneficio debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (!float.TryParse(txtiva.Text.Trim(), out _))
                                MessageBox.Show("El iva debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (!float.TryParse(txtDescFlia.Text.Trim(), out _))
                                MessageBox.Show("El descuento debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        if (hasBene)
                        {
                            if (float.TryParse(txtbgral.Text.Trim(), out _))
                            {
                                decimal.TryParse(txtbgral.Text.Trim(), out var benVal);
                                UpdateFamiliasGeneral(benVal, null, null);
                            }
                            else
                            {
                                MessageBox.Show("El beneficio debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        if (hasIva)
                        {
                            if (float.TryParse(txtiva.Text.Trim(), out _))
                            {
                                decimal.TryParse(txtiva.Text.Trim(), out var ivaVal);
                                UpdateFamiliasGeneral(null, ivaVal, null);
                            }
                            else
                            {
                                MessageBox.Show("El iva debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }

                        if (hasDesc)
                        {
                            if (float.TryParse(txtDescFlia.Text.Trim(), out _))
                            {
                                decimal.TryParse(txtDescFlia.Text.Trim(), out var descVal);
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

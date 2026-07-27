using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    public partial class frmDescMarcas : Form
    {
        public frmDescMarcas()
        {
            InitializeComponent();
            ConfigureDataGridView();
            DataGridView1.CellEndEdit += DataGridView1_CellEndEdit;
            DataGridView1.DataError += DataGridView1_DataError;
            btnBuscar.Click += btnBuscar_Click;
            txtBuscarMarca.KeyDown += txtBuscarMarca_KeyDown;
        }

        private void ConfigureDataGridView()
        {
            DataGridView1.AutoGenerateColumns = false;
            DataGridView1.AllowUserToAddRows = false;
            DataGridView1.AllowUserToDeleteRows = false;
            DataGridView1.ReadOnly = false;
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridView1.MultiSelect = false;
            DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;

            // Agregar columnas
            var colMarca = new DataGridViewTextBoxColumn
            {
                Name = "Marca",
                HeaderText = "Marca",
                DataPropertyName = "Marca",
                Width = 150,
                ReadOnly = true
            };
            DataGridView1.Columns.Add(colMarca);

            var colDesc3 = new DataGridViewTextBoxColumn
            {
                Name = "Descuento3",
                HeaderText = "Descuento 3 %",
                DataPropertyName = "Descuento3",
                Width = 100,
                ReadOnly = false
            };
            DataGridView1.Columns.Add(colDesc3);

            var colBene = new DataGridViewTextBoxColumn
            {
                Name = "Beneficio",
                HeaderText = "Beneficio %",
                DataPropertyName = "Beneficio",
                Width = 100,
                ReadOnly = false
            };
            DataGridView1.Columns.Add(colBene);

            var colIva = new DataGridViewTextBoxColumn
            {
                Name = "Iva",
                HeaderText = "IVA %",
                DataPropertyName = "Iva",
                Width = 80,
                ReadOnly = false
            };
            DataGridView1.Columns.Add(colIva);

            // Formatos
            DataGridView1.Columns["Descuento3"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridView1.Columns["Descuento3"].DefaultCellStyle.Format = "N2";
            DataGridView1.Columns["Beneficio"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridView1.Columns["Beneficio"].DefaultCellStyle.Format = "N2";
            DataGridView1.Columns["Iva"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridView1.Columns["Iva"].DefaultCellStyle.Format = "N2";
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool enabled = CheckBox1.Checked;
            lbld3gral.Enabled = enabled;
            lblbgral.Enabled = enabled;
            txtdesc3g.Enabled = enabled;
            txtbgral.Enabled = enabled;
            // Usar controles de IVA del VB
            lbliva.Enabled = enabled;
            txtiva.Enabled = enabled;
            // Mantener los controles antiguos deshabilitados
            label1.Enabled = false;
            txtIvaGral.Enabled = false;
        }

        private void btnGrabar_Click(object sender, EventArgs e)
        {
            try
            {
                if (ValidarDatos())
                {
                    if (CheckBox1.Checked == false)
                    {
                        // Carga por grilla (recorre filas)
                        ActualizarFilasDesdeGrilla();
                    }
                    else
                    {
                        // Carga general
                        CargaGral();
                    }

                    // Actualizar la grilla
                    ActualizarGrilla();

                    MessageBox.Show("Datos actualizados correctamente",
                        "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ptbImgSi.Visible = true;
                    ptbImgNo.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al grabar los datos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ptbImgSi.Visible = false;
                ptbImgNo.Visible = true;
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            if (ptbImgSi.Visible)
            {
                if (!Module.DescuentosConfigurados())
                {
                    frmPrincipal.Instance?.MostrarNotificacionDescuento();
                    MessageBox.Show(
                        "Descuentos no configurados (DESDESC1 = 0).\nPida sus descuentos a la distribuidora.",
                        "Sin Descuentos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    Module.Update_Precios();
                    Module.Update_Precios_Ofertas();
                }
            }
            this.Close();
        }

        private bool ValidarDatos()
        {
            if (CheckBox1.Checked)
            {
                if (!string.IsNullOrWhiteSpace(txtdesc3g.Text) && !ValidarCampoNumerico(txtdesc3g.Text))
                {
                    MessageBox.Show("El descuento general debe ser un número válido",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                if (!string.IsNullOrWhiteSpace(txtbgral.Text) && !ValidarCampoNumerico(txtbgral.Text))
                {
                    MessageBox.Show("El beneficio general debe ser un número válido",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                if (!string.IsNullOrWhiteSpace(txtiva.Text) && !ValidarCampoNumerico(txtiva.Text))
                {
                    MessageBox.Show("El IVA general debe ser un número válido",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            return true;
        }

        private bool ValidarCampoNumerico(string valor)
        {
            return !string.IsNullOrWhiteSpace(valor) && float.TryParse(valor.Trim(), out _);
        }

        private void ActualizarFilasDesdeGrilla()
        {
            try
            {
                if (DataGridView1.RowCount > 0)
                {
                    for (int i = 0; i < DataGridView1.Rows.Count; i++)
                    {
                        var row = DataGridView1.Rows[i];
                        if (row.IsNewRow) continue;

                        string marca = Convert.ToString(row.Cells["Marca"].Value);
                        string desc3 = Convert.ToString(row.Cells["Descuento3"].Value);
                        string bene = Convert.ToString(row.Cells["Beneficio"].Value);
                        string iva = Convert.ToString(row.Cells["Iva"].Value);
                        if (string.IsNullOrWhiteSpace(marca)) continue;

                        UpdateProveedorAllFields(marca, desc3, bene, iva);
                    }
                }
                else
                {
                    MessageBox.Show("No hay registros en la grilla", "Sin Datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al actualizar filas: {ex.Message}");
            }
        }

        private void ActualizarGrilla()
        {
            try
            {
                Module.ConectarDBUsuario();
                using (var da = new OleDbDataAdapter("SELECT DISTINCT PRRAZO as Marca, PRDES3 as Descuento3, PRBEN1 as Beneficio, PRIVA as Iva FROM PROVEEDOR", Module.cnnU))
                {
                    var dt = new DataTable();
                    da.Fill(dt);
                    DataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al actualizar la grilla: {ex.Message}");
            }
        }

        private void ActualizarGrilla(string filtroMarca)
        {
            try
            {
                Module.ConectarDBUsuario();
                string sql = "SELECT DISTINCT PRRAZO as Marca, PRDES3 as Descuento3, PRBEN1 as Beneficio, PRIVA as Iva FROM PROVEEDOR";
                if (!string.IsNullOrWhiteSpace(filtroMarca))
                {
                    sql += " WHERE PRRAZO LIKE ?";
                }

                using (var cmd = new OleDbCommand(sql, Module.cnnU))
                {
                    if (!string.IsNullOrWhiteSpace(filtroMarca))
                    {
                        cmd.Parameters.AddWithValue("@p1", "%" + filtroMarca.Trim() + "%");
                    }
                    using (var da = new OleDbDataAdapter(cmd))
                    {
                        var dt = new DataTable();
                        da.Fill(dt);
                        DataGridView1.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al actualizar la grilla: {ex.Message}");
            }
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            try
            {
                ActualizarGrilla(txtBuscarMarca.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al buscar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtBuscarMarca_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnBuscar_Click(sender, EventArgs.Empty);
            }
        }

        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Valor inválido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.ThrowException = false;
        }

        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var grid = DataGridView1;
            var colName = grid.Columns[e.ColumnIndex].Name;
            if (colName != "Descuento3" && colName != "Beneficio" && colName != "Iva") return;

            var row = grid.Rows[e.RowIndex];
            var marcaObj = row.Cells["Marca"].Value;
            var valorObj = row.Cells[e.ColumnIndex].Value;
            if (marcaObj == null) return;

            string marca = Convert.ToString(marcaObj)?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(marca)) return;

            if (valorObj == null || string.IsNullOrWhiteSpace(Convert.ToString(valorObj)))
            {
                MessageBox.Show("El valor no puede estar vacío.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ActualizarGrilla();
                return;
            }

            if (!float.TryParse(Convert.ToString(valorObj).Trim(), out float valor) || valor < 0)
            {
                MessageBox.Show("Ingrese un número válido mayor o igual a 0.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ActualizarGrilla();
                return;
            }

            try
            {
                ActualizarCampoMarca(marca, colName, valor);
                // Opcional: refrescar para asegurar consistencia visual y formateos
                ActualizarGrilla();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al actualizar la marca: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ActualizarGrilla();
            }
        }

        private void ActualizarCampoMarca(string marca, string columna, float valor)
        {
            try
            {
                Module.ConectarDBUsuario();
                using (var cmd = new OleDbCommand())
                {
                    cmd.Connection = Module.cnnU;
                    if (columna == "Descuento3")
                        cmd.CommandText = "UPDATE PROVEEDOR SET PRDES3 = @v WHERE PRRAZO = @m";
                    else if (columna == "Beneficio")
                        cmd.CommandText = "UPDATE PROVEEDOR SET PRBEN1 = @v WHERE PRRAZO = @m";
                    else if (columna == "Iva")
                        cmd.CommandText = "UPDATE PROVEEDOR SET PRIVA = @v WHERE PRRAZO = @m";
                    else
                        throw new InvalidOperationException("Columna no soportada");

                    cmd.Parameters.AddWithValue("@v", valor);
                    cmd.Parameters.AddWithValue("@m", marca);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al actualizar campo por marca: {ex.Message}");
            }
        }

        private void frmDescMarcas_Load(object sender, EventArgs e)
        {
            try
            {
                Inicializacion();
                ActualizarGrilla();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar datos: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Métodos portados del VB
        private void Inicializacion()
        {
            OleDbCommand cmMarcas = new OleDbCommand();
            OleDbDataReader drMarcas = null;
            try
            {
                Module.ConectarDBArti();
                cmMarcas.Connection = Module.cnn;
                cmMarcas.CommandType = CommandType.Text;
                //cmMarcas.CommandText = "SELECT DISTINCT prmarc from proveedo where prcrub = 'R' order by prmarc";
                cmMarcas.CommandText = "SELECT DISTINCT FACODI, FADESC, FAMIVA from FAMILIAS order by FACODI";
                drMarcas = cmMarcas.ExecuteReader();
                while (drMarcas.Read())
                {
                    string marca = drMarcas["FADESC"].ToString();//antes prmarc
                    if (!MarcaExiste(marca))
                    {
                        InsertarMarca(marca);
                    }
                }
                drMarcas.Close();
                if (Module.cnn.State == ConnectionState.Open) Module.cnn.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnn.State == ConnectionState.Open) Module.cnn.Close();
                throw new Exception(", dentro de la funcion 'Inicializacion': " + ex.Message);
            }
            finally
            {
                cmMarcas.Dispose();
            }
        }

        private bool MarcaExiste(string marca)
        {
            OleDbCommand cmMarca = new OleDbCommand();
            OleDbDataReader drMarca = null;
            bool existe = false;
            try
            {
                Module.ConectarDBUsuario();
                cmMarca.Connection = Module.cnnU;
                cmMarca.CommandType = CommandType.Text;
                cmMarca.CommandText = "select PRRAZO from PROVEEDOR where PRRAZO=?";
                cmMarca.Parameters.AddWithValue("@p1", marca);
                drMarca = cmMarca.ExecuteReader();
                if (drMarca.Read()) existe = true;
                drMarca.Close();
                if (Module.cnnU.State == ConnectionState.Open) Module.cnnU.Close();
                return existe;
            }
            catch
            {
                if (Module.cnnU.State == ConnectionState.Open) Module.cnnU.Close();
                return existe;
            }
            finally
            {
                cmMarca.Dispose();
            }
        }

        private void InsertarMarca(string marca)
        {
            OleDbCommand cmMarca = new OleDbCommand();
            try
            {
                Module.ConectarDBUsuario();
                cmMarca.Connection = Module.cnnU;
                cmMarca.CommandType = CommandType.Text;
                cmMarca.CommandText = "insert into PROVEEDOR (PRRAZO,PRDES3,PRBEN1,PRIVA) values (?,?,?,?)";
                cmMarca.Parameters.AddWithValue("@p1", marca);
                cmMarca.Parameters.AddWithValue("@p2", 0);
                cmMarca.Parameters.AddWithValue("@p3", 30);
                cmMarca.Parameters.AddWithValue("@p4", 21);
                cmMarca.ExecuteNonQuery();
                if (Module.cnnU.State == ConnectionState.Open) Module.cnnU.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open) Module.cnnU.Close();
                throw new Exception(", dentro de la accion 'InsertarMarca': " + ex.Message);
            }
        }

        private void UpdateProveedorAllFields(string marca, string desc3, string ben1, string iva)
        {
            try
            {
                if (!float.TryParse(desc3?.Trim(), out var vDesc3)) vDesc3 = 0;
                if (!float.TryParse(ben1?.Trim(), out var vBen1)) vBen1 = 0;
                if (!float.TryParse(iva?.Trim(), out var vIva)) vIva = 0;

                Module.ConectarDBUsuario();
                using (var cmd = new OleDbCommand("UPDATE PROVEEDOR SET PRDES3=?, PRBEN1=?, PRIVA=? WHERE PRRAZO=?", Module.cnnU))
                {
                    cmd.Parameters.AddWithValue("@p1", vDesc3);
                    cmd.Parameters.AddWithValue("@p2", vBen1);
                    cmd.Parameters.AddWithValue("@p3", vIva);
                    cmd.Parameters.AddWithValue("@p4", marca);
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(",Error dentro de 'UpdateProveedorAllFields' " + ex.Message);
            }
            finally
            {
                if (Module.cnnU.State == ConnectionState.Open) Module.cnnU.Close();
            }
        }

        private void CargaGral()
        {
            try
            {
                if (DataGridView1.RowCount > 0)
                {
                    bool hasDesc = !string.IsNullOrWhiteSpace(txtdesc3g.Text);
                    bool hasBene = !string.IsNullOrWhiteSpace(txtbgral.Text);
                    bool hasIva = !string.IsNullOrWhiteSpace(txtiva.Text);

                    if (hasDesc && hasBene && hasIva)
                    {
                        if (!ValidarCampoNumerico(txtdesc3g.Text)) { MessageBox.Show("El descuento debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                        if (!ValidarCampoNumerico(txtbgral.Text)) { MessageBox.Show("El beneficio debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                        if (!ValidarCampoNumerico(txtiva.Text)) { MessageBox.Show("El iva debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                        using (var cmd = new OleDbCommand("UPDATE PROVEEDOR SET PRDES3=?, PRBEN1=?, PRIVA=?", Module.cnnU))
                        {
                            Module.ConectarDBUsuario();
                            cmd.Parameters.AddWithValue("@p1", float.Parse(txtdesc3g.Text.Trim()));
                            cmd.Parameters.AddWithValue("@p2", float.Parse(txtbgral.Text.Trim()));
                            cmd.Parameters.AddWithValue("@p3", float.Parse(txtiva.Text.Trim()));
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        if (hasDesc)
                        {
                            if (!ValidarCampoNumerico(txtdesc3g.Text)) { MessageBox.Show("El descuento debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                            using (var cmd = new OleDbCommand("UPDATE PROVEEDOR SET PRDES3=?", Module.cnnU))
                            {
                                Module.ConectarDBUsuario();
                                cmd.Parameters.AddWithValue("@p1", float.Parse(txtdesc3g.Text.Trim()));
                                cmd.ExecuteNonQuery();
                            }
                        }
                        if (hasBene)
                        {
                            if (!ValidarCampoNumerico(txtbgral.Text)) { MessageBox.Show("El beneficio debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                            using (var cmd = new OleDbCommand("UPDATE PROVEEDOR SET PRBEN1=?", Module.cnnU))
                            {
                                Module.ConectarDBUsuario();
                                cmd.Parameters.AddWithValue("@p1", float.Parse(txtbgral.Text.Trim()));
                                cmd.ExecuteNonQuery();
                            }
                        }
                        if (hasIva)
                        {
                            if (!ValidarCampoNumerico(txtiva.Text)) { MessageBox.Show("El iva debe ser numerico", "Valor incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                            using (var cmd = new OleDbCommand("UPDATE PROVEEDOR SET PRIVA=?", Module.cnnU))
                            {
                                Module.ConectarDBUsuario();
                                cmd.Parameters.AddWithValue("@p1", float.Parse(txtiva.Text.Trim()));
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No hay registros para actualizar", "Sin Marcas", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(", error dentro de 'CargaGral': " + ex.Message);
            }
        }
    }
}

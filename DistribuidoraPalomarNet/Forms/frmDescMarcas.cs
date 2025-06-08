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
        }

        private void ConfigureDataGridView()
        {
            DataGridView1.AutoGenerateColumns = false;
            DataGridView1.AllowUserToAddRows = false;
            DataGridView1.AllowUserToDeleteRows = false;
            DataGridView1.ReadOnly = true;
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridView1.MultiSelect = false;

            // Agregar columnas
            DataGridView1.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Marca",
                HeaderText = "Marca",
                DataPropertyName = "Marca",
                Width = 150
            });

            DataGridView1.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Descuento3",
                HeaderText = "Descuento 3 %",
                DataPropertyName = "Descuento3",
                Width = 100
            });

            DataGridView1.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Beneficio",
                HeaderText = "Beneficio %",
                DataPropertyName = "Beneficio",
                Width = 100
            });
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool enabled = CheckBox1.Checked;
            lbld3gral.Enabled = enabled;
            lblbgral.Enabled = enabled;
            txtdesc3g.Enabled = enabled;
            txtbgral.Enabled = enabled;
        }

        private void btnGrabar_Click(object sender, EventArgs e)
        {
            try
            {
                if (ValidarDatos())
                {
                    if (CheckBox1.Checked)
                    {
                        // Actualizar valores generales
                        ActualizarValoresGenerales();
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
                Module.Update_Precios();
                Module.Update_Precios_Ofertas();
            }
            this.Close();
        }

        private bool ValidarDatos()
        {
            if (CheckBox1.Checked)
            {
                if (!ValidarCampoNumerico(txtdesc3g.Text))
                {
                    MessageBox.Show("El descuento general debe ser un número válido",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                if (!ValidarCampoNumerico(txtbgral.Text))
                {
                    MessageBox.Show("El beneficio general debe ser un número válido",
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

        private void ActualizarValoresGenerales()
        {
            try
            {
                Module.ConectarDBUsuario();
                using (var cmd = new OleDbCommand())
                {
                    cmd.Connection = Module.cnnU;
                    cmd.CommandText = @"UPDATE Marcas 
                                      SET Descuento3 = @desc3, Beneficio = @bene";
                    
                    cmd.Parameters.AddWithValue("@desc3", float.Parse(txtdesc3g.Text.Trim()));
                    cmd.Parameters.AddWithValue("@bene", float.Parse(txtbgral.Text.Trim()));
                    
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al actualizar valores generales: {ex.Message}");
            }
        }

        private void ActualizarGrilla()
        {
            try
            {
                Module.ConectarDBUsuario();
                using (var da = new OleDbDataAdapter("SELECT Marca, Descuento3, Beneficio FROM Marcas", Module.cnnU))
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
    }
}

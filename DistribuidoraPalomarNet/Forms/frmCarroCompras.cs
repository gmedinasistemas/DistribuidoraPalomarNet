using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;

namespace DistribuidoraPalomarNet
{
    public partial class frmCarroCompras : Form
    {
        public frmCarroCompras()
        {
            InitializeComponent();
            Module.SetCarroComprasForm(this);
        }

        private void frmCarroCompras_Load(object sender, EventArgs e)
        {
            try
            {
                Module.funcion_inicial_carrito();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error dentro del Load Carro Compras: " + ex.Message);
            }
        }

        private void btnEliminarItem_Click(object sender, EventArgs e)
        {
            try
            {
                int numFila = this.gridPedidos.CurrentCell.RowIndex;
                int item = Convert.ToInt32(this.gridPedidos[0, numFila].Value);
                if (item > 0)
                {
                    if (MessageBox.Show($"¿Desea Eliminar el item nº: {this.gridPedidos[0, numFila].Value}, Codigo: {this.gridPedidos[1, numFila].Value}?",
                        "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        if (ModuloPreparac.EliminarItemCarro(item))
                        {
                            // Eliminado correctamente
                        }
                    }
                    Module.funcion_inicial_carrito();
                }
                btnEliminarItem.Enabled = false;
            }
            catch (Exception ex)
            {
                btnEliminarItem.Enabled = false;
                MessageBox.Show("Error dentro de Eliminar Item Click" + ex.Message, "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVaciarCarro_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridPedidos.RowCount > 0)
                {
                    if (MessageBox.Show("¿Desea Eliminar todos los registros?", "Confirmar acción",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        if (ModuloPreparac.EliminarItemsCarro())
                        {
                            MessageBox.Show("Se eliminaron Correctamente los registros", "Acción Correcta",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("No se eliminaron los registros", "Acción Incompleta",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No hay registros", "Peticion denegada",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                Module.funcion_inicial_carrito();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error dentro de Vaciar Carro Click" + ex.Message, "Alerta",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportarCarro_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridPedidos.RowCount > 0)
                {
                    btnExportarCarro.Enabled = false;
                    Label2.Visible = true;
                    Label2.Refresh();

                    var doc = new Document(PageSize.A4, 10, 10, 10, 10);
                    string filename = $@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\pedidos\Preimpreso {Module._frmPrincipal.lblusuario.Text} {DateTime.Now.ToString().Replace("/", "-").Replace(":", "")}.pdf";
                    var file = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                    PdfWriter.GetInstance(doc, file);
                    doc.Open();
                    ExportarDatosPDF(doc);
                    doc.Close();
                    Process.Start(filename);
                    Label2.Visible = false;
                    Label2.Refresh();
                    btnExportarCarro.Enabled = true;
                }
                else
                {
                    MessageBox.Show("No hay registros", "Peticion denegada",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Label2.Visible = false;
                Label2.Refresh();
                btnExportarCarro.Enabled = true;
                MessageBox.Show("No se puede generar el documento PDF.: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnPedirCarro_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridPedidos.RowCount > 0)
                {
                    btnPedirCarro.Enabled = false;
                    PictureBox1.Visible = true;
                    Label2.Visible = true;
                    Label2.Refresh();
                    PictureBox1.Refresh();

                    Module.ProcesoCorreo();

                    if (VerificarConexionURL("http://www.sistemavs.com.ar/webs/WSPalomar.asmx"))
                    {
                        ModuloPreparac.ProcesoWSAsync();
                    }

                    ModuloPreparac.EliminarItemsCarro();
                    lblTotal.Text = "";
                    lblTotal.Refresh();
                    Module.funcion_inicial_carrito();
                    PictureBox1.Visible = false;
                    Label2.Visible = false;
                    Label2.Refresh();
                    PictureBox1.Refresh();
                    btnPedirCarro.Enabled = true;
                }
                else
                {
                    MessageBox.Show("No hay registros", "Petición denegada",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                PictureBox1.Visible = false;
                Label2.Visible = false;
                Label2.Refresh();
                PictureBox1.Refresh();
                btnPedirCarro.Enabled = true;
                MessageBox.Show("Dentro de la accion 'Pasar Pedido': " + ex.Message,
                    "Error inesperado", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnmulti_Click(object sender, EventArgs e)
        {
            try
            {
                Module.funcion_inicial_carrito();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error dentro de btnmulti_Click: " + ex.Message,
                    "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridPedidos_Click(object sender, EventArgs e)
        {
            btnEliminarItem.Enabled = true;
        }

        private void gridPedidos_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                int numFila = gridPedidos.CurrentCell.RowIndex;
                if (MessageBox.Show("¿Desea actualizar la  Cantidad?",
                    $"Modificar pieza: {gridPedidos[1, numFila].Value}",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string cantidad = Microsoft.VisualBasic.Interaction.InputBox(
                        "Ingrese la cantidad", "Modificacion", "1");
                    
                    if (double.TryParse(cantidad, out double cantidadNum) && cantidadNum > 0)
                    {
                        long item = Convert.ToInt64(gridPedidos[0, numFila].Value);
                        ModuloPreparac.Update_item(item, cantidad);
                    }
                    else
                    {
                        MessageBox.Show("Debe ingresar valores numericos",
                            "Dato no Valido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error dentro de la accion 'Doble Click de la grilla': " + ex.Message,
                    "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region Métodos Auxiliares
        private void ExportarDatosPDF(Document doc)
        {
            try
            {
                // Implementar la lógica para exportar a PDF
                var table = new PdfPTable(gridPedidos.Columns.Count);
                table.WidthPercentage = 100;

                // Agregar encabezados
                foreach (DataGridViewColumn column in gridPedidos.Columns)
                {
                    table.AddCell(new Phrase(column.HeaderText));
                }

                // Agregar datos
                foreach (DataGridViewRow row in gridPedidos.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        table.AddCell(new Phrase(cell.Value?.ToString() ?? ""));
                    }
                }

                doc.Add(table);
            }
            catch (Exception ex)
            {
                throw new Exception("Error en ExportarDatosPDF: " + ex.Message);
            }
        }

        private bool VerificarConexionURL(string url)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.Timeout = 5000;
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    return response.StatusCode == HttpStatusCode.OK;
                }
            }
            catch
            {
                return false;
            }
        }
        #endregion
    }
}

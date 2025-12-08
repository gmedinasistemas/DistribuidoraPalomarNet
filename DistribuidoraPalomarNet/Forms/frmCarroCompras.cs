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
                Module.Cargar_Usuario(false);
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
                    string filename = $@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\pedidos\Preimpreso {Module.US_CODI} {DateTime.Now.ToString().Replace("/", "-").Replace(":", "")}.pdf";
                    var file = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                    PdfWriter.GetInstance(doc, file);
                    doc.Open();
                    Module.ExportarDatosPDF(doc);
                    doc.Close();
                    if (File.Exists(filename))
                    {
                        Process.Start(new ProcessStartInfo(filename) { UseShellExecute = true });
                    }
                    else
                    {
                        Console.WriteLine("El archivo no existe en la ruta especificada.");
                    }
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
                    //comento el envio de mail, ya esta probado y funciona ok.
                    Module.ProcesoCorreo();

                    if (VerificarConexionURL("http://www.sistemavs.com.ar/webs/WSPalomar.asmx"))
                    {
                        //ModuloPreparac.ProcesoWSAsync();
                        ModuloPreparac.ProcesoWS();
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
        //private void ExportarDatosPDF(Document doc)
        //{
        //    try
        //    {
        //        // Implementar la lógica para exportar a PDF
        //        var table = new PdfPTable(gridPedidos.Columns.Count);
        //        table.WidthPercentage = 100;

        //        // Agregar encabezados
        //        foreach (DataGridViewColumn column in gridPedidos.Columns)
        //        {
        //            table.AddCell(new Phrase(column.HeaderText));
        //        }

        //        // Agregar datos
        //        foreach (DataGridViewRow row in gridPedidos.Rows)
        //        {
        //            foreach (DataGridViewCell cell in row.Cells)
        //            {
        //                table.AddCell(new Phrase(cell.Value?.ToString() ?? ""));
        //            }
        //        }

        //        doc.Add(table);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("Error en ExportarDatosPDF: " + ex.Message);
        //    }
        //}
        //public void ExportarDatosPDF(Document document)
        //{
        //    try
        //    {
        //        // Se crea un objeto PDFTable con el numero de columnas del DataGridView.
        //        PdfPTable datatable = new PdfPTable((gridPedidos.ColumnCount - 2)); //reemplazo -1 por -3 para sacar los precios#lo agregue gmedina

        //        // Se asignan algunas propiedades para el diseño del PDF.
        //        datatable.DefaultCell.Padding = 1;
        //        float[] headerwidths = GetColumnasSize(gridPedidos);
        //        datatable.SetWidths(headerwidths);

        //        datatable.WidthPercentage = 100;
        //        datatable.DefaultCell.BorderWidth = 1;
        //        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;

        //        // se crea el logo del pdf
        //        iTextSharp.text.Image imagendemo; // Declaracion de una imagen
        //        imagendemo = iTextSharp.text.Image.GetInstance(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\imagenes\LogoImpresion.png"); // Dirreccion a la imagen que se hace referencia
        //        imagendemo.ScaleAbsoluteWidth(150); // Ancho de la imagen
        //        imagendemo.ScaleAbsoluteHeight(45); // Altura de la imagen
        //                                            // Phrase texto2 = new Phrase("Catalogo Pedido", new Font(Font.Name = "Tahoma", 20, Font.Bold));

        //        // Se crea el encabezado en el PDF.
        //        Paragraph encabezado = new Paragraph("Reporte de Pedido. Cliente " + Module.US_CODI, FontFactory.GetFont(FontFactory.COURIER, 14, iTextSharp.text.Font.BOLD));
        //        // Se crea el texto abajo del encabezado.
        //        Phrase texto = new Phrase("Realizado la fecha : " + DateTime.Now, FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.BOLD));

        //        // Se capturan los nombres de las columnas del DataGridView.
        //        for (int i = 0; i < gridPedidos.ColumnCount; i++)
        //        {
        //            if (i != 0 && i != 4) // saco columnas de id y precios#lo agregue gmedina
        //            {
        //                datatable.AddCell(new PdfPCell(new Paragraph(gridPedidos.Columns[i].HeaderText, FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.ITALIC))));
        //            }
        //        }

        //        datatable.HeaderRows = 1;
        //        datatable.DefaultCell.BorderWidth = 1;

        //        // Se generan las columnas del DataGridView.
        //        for (int i = 0; i < gridPedidos.RowCount; i++)
        //        {
        //            for (int j = 0; j < gridPedidos.ColumnCount; j++)
        //            {
        //                if (j != 0 && j != 4) // saco columnas de id y precios#lo agregue gmedina
        //                {
        //                    datatable.AddCell(new PdfPCell(new Paragraph(gridPedidos[j, i].Value.ToString(), FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.ITALIC))));
        //                }
        //                PictureBox1.Refresh();
        //            }
        //            datatable.CompleteRow();
        //        }

        //        // Se agrega el PDFTable al documento.
        //        document.Add(imagendemo);
        //        // document.Add(texto2);
        //        document.Add(encabezado);
        //        document.Add(texto);
        //        document.Add(datatable);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(", dentro de la acción 'ExportarDatosPDF': " + ex.Message.ToString());
        //    }
        //}
        //public float[] GetColumnasSize(DataGridView dg)
        //{
        //    // funcion para obtener el tamaño de la columnas del datagridview
        //    float[] values = new float[dg.ColumnCount - 2]; // reemplazo -1 por -4#lo agregue gmedina

        //    // For loop commented out as in the original VB code
        //    //for (int i = 0; i <= dg.ColumnCount - 4; i++) // reemplazo -1 por -4#lo agregue gmedina
        //    //{
        //    //    values[i] = (float)dg.Columns[i + 1].Width;
        //    //}

        //    // Hardcoded values as in the original VB code
        //    values[0] = (float)105;
        //    values[1] = (float)95;
        //    values[2] = (float)420;
        //    values[3] = (float)215;
        //    values[4] = (float)125;
        //    values[5] = (float)175;

        //    return values;
        //}
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

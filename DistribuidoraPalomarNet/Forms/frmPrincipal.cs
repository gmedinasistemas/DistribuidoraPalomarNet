using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Net;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Image = System.Drawing.Image;
using System.Threading.Tasks;
using DistribuidoraPalomarNet.Forms;
using System.Drawing.Imaging;
using ImageViewerApp;

namespace DistribuidoraPalomarNet
{
    public partial class frmPrincipal : Form
    {
        private void DataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Application.DoEvents();
            //DataGridView2.DefaultCellStyle.SelectionBackColor = Color.Transparent;
            //DataGridView2.DefaultCellStyle.SelectionForeColor = Color.Transparent;
            ColorearCelda(DataGridView2);
        }


        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                /*
                if (llamaSimulaClick)
                {
                    Application.DoEvents();
                    if (DataGridView1.DefaultCellStyle.SelectionBackColor != Color.Transparent)
                    {
                        DataGridView1.DefaultCellStyle.SelectionBackColor = Color.Transparent;
                        DataGridView1.DefaultCellStyle.SelectionForeColor = Color.Transparent;
                    }
                    llamaSimulaClick = false;
                }
                */
                ColorearCelda(DataGridView1);
            }
            catch (Exception ex)
            {
                // Exception handling
            }
            finally
            {
                //llamaSimulaClick = false;
            }
        }


        private void ColorearCelda(DataGridView dataGridView)
        {
            // Para semáforo de Stock
            Application.DoEvents();

            // Configuración de estilo por defecto
            dataGridView.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Transparent;
            dataGridView.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Transparent;

            // Código comentado migrado a C#
            /*
            int w = dataGridView1.CurrentCell.RowIndex;
            if (dataGridView1["stock", w].Value != null)
            {
                if ((int)dataGridView1["stock", w].Value >= ini_cond_verde)
                {
                    dataGridView1["stock", w].Style.BackColor = Color.Green;
                    dataGridView1["stock", w].Style.ForeColor = Color.Green;
                    dataGridView1.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Green;
                    dataGridView1.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Green;
                }
                else if ((int)dataGridView1["stock", w].Value < ini_cond_verde && (int)dataGridView1["stock", w].Value > ini_cond_rojo)
                {
                    dataGridView1["stock", w].Style.BackColor = Color.Yellow;
                    dataGridView1["stock", w].Style.ForeColor = Color.Yellow;
                    dataGridView1.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Yellow;
                    dataGridView1.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Yellow;
                }
                else if ((int)dataGridView1["stock", w].Value <= ini_cond_rojo)
                {
                    dataGridView1["stock", w].Style.BackColor = Color.Red;
                    dataGridView1["stock", w].Style.ForeColor = Color.Red;
                    dataGridView1.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Red;
                    dataGridView1.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Red;
                }
            }
            else
            {
                dataGridView1["stock", w].Style.BackColor = Color.Red;
                dataGridView1["stock", w].Style.ForeColor = Color.White;
            }
            */
        }

        private static bool terminoTadControlSelecting = false;

        private static bool colorearGrillaenEjecucion = false;
        private static bool llamaSimulaClick = false;


        private static frmCarroCompras _frmCarroCompras = new frmCarroCompras();
        private static bool Migra = false;

        private static frmPrincipal _instance;
        public static frmPrincipal Instance
        {
            get
            {
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new frmPrincipal();
                }
                return _instance;
            }
        }

        public frmPrincipal()
        {
            InitializeComponent();
            DataGridView1.CellClick += DataGridView1_CellClick;
            DataGridView2.CellClick += DataGridView2_CellClick;
            InitializeMenuEvents();

            // Asignar manejadores de eventos para botones
            btnGenerarTxt.Click += BtnGenerarTxt_Click;
            btnEquivalencia.Click += BtnEquivalencia_Click;
            btnConfUsu.Click += BtnConfUsu_Click;
            btncerrar.Click += Btncerrar_Click;
            btnVerCarrito.Click += BtnVerCarrito_Click;
            btnLimpiar.Click += BtnLimpiar_Click;
            btnnovedades.Click += Btnnovedades_Click;
        }
        private void Btnnovedades_Click(object sender, EventArgs e)
        {
            try
            {
                Module.Funcion_Novedades();
                TimerVerOfertasEquivalencias.Enabled = true;
            }
            catch (Exception ex)
            {
                TimerVerOfertasEquivalencias.Enabled = false;
                MessageBox.Show($"Error dentro de la accion 'Novedades Click': {ex.Message}", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnActualizar_Click(object sender, EventArgs e)
        {
            try
            {
                Module.ActualizacionXFTP();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en Click 'Actualizar': " + ex.ToString(), "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnGenerarTxt_Click(object sender, EventArgs e)
        {
            try
            {
                Module.Iniciar_Proceso_TXT();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en Click 'Generar TXT': " + ex.ToString(), "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtDescripcion_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Application.DoEvents();
                if (!string.IsNullOrEmpty(txtDescripcion.Text))
                {
                    Module.campo = "descripcion";
                    Module.valor = txtDescripcion.Text;
                    if (!ckbOfertas.Checked)
                    {
                        Module.consultas_Relacionales();
                    }
                    else
                    {
                        Module.consultas_Ofertas();
                    }
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Descripción': {ex.Message}");
            }
        }

        private void BtnEquivalencia_Click(object sender, EventArgs e)
        {
            CargarEquivalencias();
        }

        private void CargarEquivalencias()
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                if (!string.IsNullOrEmpty(Module.NumEqui))
                {
                    Module.ConectarDBArti();
                    Application.DoEvents();

                    string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "'Dato Pendiente' as Linea," +
                    //"l.lidesc as Linea," +
                    "a.armode as Modelo, a.arinte as Posición, a.ARMARC as Marca," +
                    "a.arpre2 as Preciopublico, " +//"IIF(a.artt IS NOT NULL AND a.artt <> 0, a.arpre5, a.arpre2) as Preciopublico, "+
                    "IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from listas as a ";

                    //da = new OleDbDataAdapter(
                    //    "SELECT a.arprim as Primera, a.arterc as Codigo, p.prrazo as Marca, a.ardesc as Descripcion, " +
                    //    "a.arobse as Observacioness, a.arpre2 as Preciopublico , a.arprov, a.arobse, a.arubic1 as Obs from listas as a " +
                    //    "inner join proveedo as p on p.prcodi=a.arprov " +
                    //    "Where a.arubic1='" + Module.NumEqui + "' " +
                    //    "order by arcodi", Module.cnn);
                    da = new OleDbDataAdapter(
                        SelectQuery +
                        "Where a.arubic1='" + Module.NumEqui + "' " +
                        "order by arcodi", Module.cnn);

                    DataSet ds = new DataSet();
                    Application.DoEvents();
                    da.Fill(ds);
                    Application.DoEvents();
                    DataGridView2.DataSource = ds.Tables[0];
                    DataGridView2.Refresh();

                    DataGridView2.Columns["aroo"].Visible = false;
                    DataGridView2.Columns["artt"].Visible = false;
                    DataGridView2.Columns["arubic1"].Visible = false;
                    DataGridView2.Columns["arobse"].Visible = false;
                    DataGridView2.Columns["Linea"].Visible = false;
                    DataGridView2.Columns["Modelo"].Visible = false;
                    DataGridView2.Columns["Descripcion"].Visible = false;

                    var miEstiloStock = new DataGridViewCellStyle(DataGridView2.Columns["Stock"].DefaultCellStyle)
                    {
                        Font = new Font("Arial", 2, FontStyle.Regular)
                    };

                    DataGridView2.Columns["Stock"].DefaultCellStyle = miEstiloStock;


                    if (DataGridView1.RowCount > 0)
                        lblTotalRegistros.Text = "articulo: " + DataGridView2.RowCount.ToString();
                    if (Module.cnn.State == ConnectionState.Open)
                    {
                        Module.cnn.Close();
                    }
                }
                //btnEquivalencia.Enabled = false;
                TimerVerOfertasEquivalencias.Enabled = true;
            }
            catch (Exception ex)
            {
                TimerVerOfertasEquivalencias.Enabled = false;
                //btnEquivalencia.Enabled = false;
                if (Module.cnn.State == ConnectionState.Open)
                {
                    Module.cnn.Close();
                }
                Module.GrabarError(" - " + DateTime.Now.ToLongDateString() + " => " +
                            DateTime.Now.ToString("hh:mm:ss") +
                            " 'Metodo CargarEquivalencias': " + ex.Message);
            }
            finally
            {
                da.Dispose();
            }
        }

        private void BtnConfUsu_Click(object sender, EventArgs e)
        {
            try
            {
                var frmDatos = new frmDatosUsuario();
                frmDatos.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir configuración: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Btncerrar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void BtnVerCarrito_Click(object sender, EventArgs e)
        {
            try
            {
                _frmCarroCompras.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error dentro de 'Ver Carrito': " + ex.Message);
            }
        }

        private void BtnLimpiar_Click(object sender, EventArgs e)
        {
            try
            {
                btnLimpiar.Enabled = false;
                Module.LimpiarCampos(0);
                btnLimpiar.Enabled = true;
            }
            catch (Exception ex)
            {
                btnLimpiar.Enabled = true;
                Module.GrabarError(" - " + DateTime.Now.ToLongDateString() + " => " +
                            DateTime.Now.ToString("hh:mm:ss") +
                            " 'Click Limpiar': " + ex.Message);
            }
        }

        private void InitializeMenuEvents()
        {
            // Archivo
            ActualizarBasesToolStripMenuItem.Click += new EventHandler(ActualizarBasesToolStripMenuItem_Click);
            ActualizarDatosLocalToolStripMenuItem.Click += new EventHandler(ActualizarDatosLocalToolStripMenuItem_Click);

            // Configuración
            ClientesToolStripMenuItem.Click += new EventHandler(ClientesToolStripMenuItem_Click);
            DescPorMarcaToolStripMenuItem.Click += new EventHandler(DescPorMarcaToolStripMenuItem_Click);

            // Ver
            PedidosEnviadosToolStripMenuItem.Click += new EventHandler(PedidosEnviadosToolStripMenuItem_Click);

            // Web y Acerca de
            WwwsistemasvscomarToolStripMenuItem.Click += new EventHandler(WwwsistemasvscomarToolStripMenuItem_Click);
            AcercaDeToolStripMenuItem.Click += new EventHandler(AcercaDeToolStripMenuItem_Click);
        }

        private void ActualizarBasesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Module.ActualizacionXFTP();
                Module.consultas_Relacionales();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Actualizar Bases': {ex.Message}");
                MessageBox.Show("Error al actualizar bases: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ActualizarDatosLocalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Module.UpdateValorDBLocal("");  // Pasando string vacío como parámetro
                Module.consultas_Relacionales();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Actualizar Datos Local': {ex.Message}");
                MessageBox.Show("Error al actualizar datos locales: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ClientesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var frmDatos = new frmDatosUsuario();
                frmDatos.ShowDialog();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Abrir Clientes': {ex.Message}");
                MessageBox.Show("Error al abrir datos de clientes: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DescPorMarcaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var frmDescMarcas = new frmDescMarcas();
                frmDescMarcas.ShowDialog();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Abrir Desc. por Marca': {ex.Message}");
                MessageBox.Show("Error al abrir descuentos por marca: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PedidosEnviadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (Module.GetCarroComprasForm() == null || Module.GetCarroComprasForm().IsDisposed)
                {
                    var frmCarroCompras = new frmCarroCompras();
                    Module.SetCarroComprasForm(frmCarroCompras);
                    frmCarroCompras.Show();
                }
                else
                {
                    Module.GetCarroComprasForm().BringToFront();
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Ver Pedidos Enviados': {ex.Message}");
                MessageBox.Show("Error al abrir pedidos enviados: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void WwwsistemasvscomarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "http://www.sistemasvs.com.ar",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Abrir Web': {ex.Message}");
                MessageBox.Show("Error al abrir sitio web: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AcercaDeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Distribuidora Palomar\nVersión 1.0\n\nDesarrollado por Sistemas VS\nwww.sistemasvs.com.ar", "Acerca de", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void frmPrincipal_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Al momento de cerrar, para liberar la memoria
            Module.vectornomfoto = null;
        }

        private void frmPrincipal_Load(object sender, EventArgs e)
        {
            try
            {
                // Carga general de combos, grillas, imágenes (banners)
                Module.LeerConfig();
                bool ret = Module.DescargaArchivo("BannerPalomar.gif",
                    "C:\\DistribuidoraPalomarNet\\DistribuidoraPalomarNet\\Imagenes\\BannerPalomar.gif",
                    Module.ini_usu_ftp, Module.ini_pass_ftp, Module.ini_ruta_ftp);
                CargarBanner();
                Module.Inicializacion();
                Module.BuscarActualizacionDisponible();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Inicio Formulario Principal': {ex.Message}");
            }
            finally
            {
                //// Configura el Panel con 70% de opacidad (Alpha = 178)
                //panel2.BackColor = Color.FromArgb(178, panel2.BackColor); // 178 ≈ 70% opacidad
                //PictureBox3.Parent = panel2;
                //PictureBox3.BackColor = Color.Transparent; // Asegura que el PictureBox no tape el Panel
            }
        }

        private void btncerrar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            try
            {
                btnLimpiar.Enabled = false;
                Module.LimpiarCampos(0);
                btnLimpiar.Enabled = true;
            }
            catch (Exception ex)
            {
                btnLimpiar.Enabled = true;
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Click Limpiar': {ex.Message}");
            }
        }

        private void DataGridView1_Click(object sender, EventArgs e)
        {
            try
            {
                if (DataGridView1.CurrentRow != null)
                {
                    Module.AccionGrilla(DataGridView1);
                    CargarEquivalencias();

                    int numFila = DataGridView1.CurrentCell.RowIndex;
                    string StringSelect = "SELECT TOP 1 IIF(R.RBDESC IS NULL, '', R.RBDESC) AS RBDESC, IIF(A.ARORIG IS NULL, '', A.ARORIG) AS ARORIG" +
                        ", IIF(A.AROBSE IS NULL, '', A.AROBSE) AS AROBSE " +
                        ", IIF(A.ARUBIC1 IS NULL, '', A.ARUBIC1) AS ARUBIC1 " +
                    " from listas AS A " +
                    " INNER JOIN RUBROS AS R ON A.ARRUB2 = R.RBCODI  " +
                    " WHERE A.ARTERC='" + this.DataGridView1["CODIGO", numFila].Value.ToString() +
                    "' AND A.ARMARC='" + this.DataGridView1["Marca", numFila].Value.ToString() + "'";
                    Module.CargarDatosExtras(StringSelect);
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Click Grilla': {ex.Message}");
            }
        }

        private void ptbImagen_Click(object sender, EventArgs e)
        {
            try
            {
                /*var imagenForm = new frmImagen();
                imagenForm.ptbImagen.Image = this.ptbImagen.Image;
                imagenForm.Show();*/

                if (DataGridView1.CurrentRow != null)
                {

                    int numFila = DataGridView1.CurrentCellAddress.Y;
                    //string imagePath = Path.Combine("C:", "DistribuidoraPalomarNet", "DistribuidoraPalomarNet", "imagenes", "FOTOS",
                    //    $"{DataGridView1["Codigo", numFila].Value}.jpg");

                    var imgForm = new frmImagenZoom();
                    imgForm.txtCode.Text = DataGridView1["Codigo", numFila].Value.ToString();
                    imgForm.Show();
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError(" - " + DateTime.Now.ToLongDateString() + " => " +
                           DateTime.Now.ToString("hh:mm:ss") +
                           " 'Click Imagen': " + ex.Message.ToString());
            }
        }

        private void Pintar_Ofertas()
        {
            try
            {
                foreach (DataGridViewRow row in DataGridView1.Rows)
                {
                    if (row.Cells["oferta"].Value != null && row.Cells["oferta"].Value.ToString() == "1")
                    {
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Pintar Ofertas': {ex.Message}");
            }
        }

        private void Despintar_Ofertas()
        {
            try
            {
                foreach (DataGridViewRow row in DataGridView1.Rows)
                {
                    row.DefaultCellStyle.BackColor = Color.White;
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Despintar Ofertas': {ex.Message}");
            }
        }

        // Métodos auxiliares
        private void CargarBanner()
        {
            try
            {
                imgBanner.Image = Image.FromFile("C:\\DistribuidoraPalomarNet\\DistribuidoraPalomarNet\\Imagenes\\BannerPalomar.gif");
                imgBanner.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Carga Banner': {ex.Message}");
            }
        }

        private void CheckBox1_Click(object sender, EventArgs e)
        {
            bool visible = CheckBox1.Checked;

            // Labels
            lblPL.Visible = visible;
            lblD1.Visible = visible;
            lblD2.Visible = visible;
            lblPC.Visible = visible;
            lblB.Visible = visible;
            lblPP.Visible = visible;

            // TextBoxes
            txtPreciolista.Visible = visible;
            txtDescuento1.Visible = visible;
            txtDescuento2.Visible = visible;
            txtPreciopublico.Visible = visible;
            txtBeneficio.Visible = visible;
            txtSugerido.Visible = visible;
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            //Label1.Text = (int.Parse(Label1.Text) + 1).ToString();
            toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
            if (toolStripProgressBar1.Value == 10)
            {
                toolStripProgressBar1.Value = 0;
            }
            else
            {
                toolStripProgressBar1.Value++;
            }
            //toolStripProgressBar1.Refresh();
        }

        private void txtCodigo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                try
                {
                    if (!string.IsNullOrEmpty(txtCodigo.Text))
                    {
                        string queryRel = ckbOfertas.Checked
                            ? "SELECT a.arterc as Codigo, A.ARMARC as Marca, a.ardesc as Descripcion, " +
                              "a.arpre5 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                              ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse from listas as a " +
                              "where a.ARTT<>0 AND a.arprim+a.arterc like '%" + txtCodigo.Text +
                              "%' order by a.arcodi"
                            : "SELECT a.arterc as Codigo, A.ARMARC as Marca, a.ardesc as Descripcion, " +
                              "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                              ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse from listas as a " +
                              "where a.arprim+a.arterc like '%" + txtCodigo.Text +
                              "%' order by a.arcodi";

                        using (var da = new OleDbDataAdapter())
                        using (var ds = new DataSet())
                        {
                            Module.ConectarDBArti();

                            da.SelectCommand = new OleDbCommand(queryRel, Module.cnn);
                            da.Fill(ds);

                            if (ds.Tables.Count > 0)
                            {
                                DataGridView1.DataSource = ds.Tables[0];
                                DataGridView1.Refresh();
                            }

                            if (DataGridView1.RowCount > 0)
                            {
                                Module.SeleccionarGrilla();
                            }

                            if (Module.cnn.State == ConnectionState.Open)
                            {
                                Module.cnn.Close();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (Module.cnn.State == ConnectionState.Open)
                    {
                        Module.cnn.Close();
                    }
                    Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Código': {ex.Message}");
                }
                finally
                {
                    //If TimerVerOfertas.Enabled = False Then TimerVerOfertas.Enabled = True
                    if (TimerVerOfertas.Enabled == false) TimerVerOfertas.Enabled = true;
                }
            }
        }

        private void ckbOfertas_Click(object sender, EventArgs e)
        {
            try
            {
                // Limpio todo los campos
                // filtro las ofertas de articulo
                if (ckbOfertas.Checked)
                {
                    Module.consultas_Ofertas();
                    //Pintar_Ofertas();
                    if (TimerVerOfertas.Enabled == false) TimerVerOfertas.Enabled = true;
                }
                else
                {
                    if (!Migra)
                    {
                        Module.consultas_Relacionales();
                    }
                    //Despintar_Ofertas();
                }
            }
            catch (Exception ex)
            {
                TimerVerOfertas.Enabled = false;
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Ver Ofertas': {ex.Message}");
            }
        }

        private void cmbRubros_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(cmbRubros.Text) && cmbRubros.Text != "System.Data.DataRowView")
                {
                    Module.campo = string.Empty;
                    Module.valor = string.Empty;
                    Module.LimpiarCampos(1);

                    if (cmbRubros.Text == " TODOS" || cmbRubros.Text == "TODOS")
                    {
                        cmbRubros.Text = string.Empty;
                        Module.funcion_inicial(0);
                    }
                    else
                    {
                        Module.consultas_Relacionales();
                    }
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Rubros': {ex.Message}");
            }
        }

        private void txtEquivalencia_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                try
                {
                    if (!string.IsNullOrEmpty(txtEquivalencia.Text))
                    {
                        Module.campo = "equivalencia";
                        Module.valor = txtEquivalencia.Text;
                        if (!ckbOfertas.Checked)
                        {
                            Module.consultas_Relacionales();
                        }
                        else
                        {
                            Module.consultas_Ofertas();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Equivalencia': {ex.Message}");
                }
            }
        }

        private void Update_Precios_Ofertas()
        {
            try
            {
                Module.Update_PrecioCosto_Ofertas();
                Module.Update_PrecioSug_Ofertas();
                Module.consultas_Ofertas();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Update Precios Ofertas': {ex.Message}");
            }
        }

        private void Update_Precios()
        {
            try
            {
                Module.Update_PrecioCosto();
                Module.Update_PrecioSug();
                Module.consultas_Relacionales();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Update Precios': {ex.Message}");
            }
        }

        private void btnOfertas_Click(object sender, EventArgs e)
        {
            try
            {
                if (ckbOfertas.Checked)
                {
                    ckbOfertas.Checked = false;
                }
                else
                {
                    ckbOfertas.Checked = true;
                }

                ckbOfertas_Click(sender, e);
                TimerVerOfertasEquivalencias.Enabled = true;
            }
            catch (Exception ex)
            {
                TimerVerOfertasEquivalencias.Enabled = false;

                throw;
            }
        }

        private void cmbLineas_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Verificar si hay un elemento seleccionado y que SelectedValue no sea null
                if (cmbLineas.SelectedValue != null && int.TryParse(cmbLineas.SelectedValue.ToString(), out int selectedValue))
                {
                    if (selectedValue != 0)
                    {
                        Filtrar_x_linea(selectedValue);
                    }
                }
                //cmbLineas.Visible = false;
            }
            catch (Exception ex)
            {
                //cmbLineas.Visible = false;
                MessageBox.Show("Error dentro de Filtro por Montadora: " + ex.Message, "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //try
            //{
            //    if (!string.IsNullOrEmpty(cmbLineas.Text) && cmbLineas.Text != "System.Data.DataRowView")
            //    {
            //        Module.campo = string.Empty;
            //        Module.valor = string.Empty;
            //        Module.LimpiarCampos(1);

            //        if (cmbLineas.Text == " TODOS" || cmbLineas.Text == "TODOS")
            //        {
            //            cmbLineas.Text = string.Empty;
            //            Module.funcion_inicial(0);
            //        }
            //        else
            //        {
            //            Module.consultas_Relacionales();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Lineas': {ex.Message}");
            //}
        }
        private void Filtrar_x_linea(int numL)
        {
            OleDbDataAdapter da = null;
            try
            {
                Application.DoEvents();
                string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "'Dato Pendiente' as Linea," +
                    //"l.lidesc as Linea," +
                    "a.armode as Modelo, a.arinte as Posición, a.ARMARC as Marca," +
                    "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from listas as a ";

                da = new OleDbDataAdapter(SelectQuery +
                                         "inner join proveedo as p on p.prcodi=a.arprov " +
                                         "where a.arcuar='" + numL.ToString() + "' " +
                                         "order by a.arcodi", Module.cnn);

                DataSet ds = new DataSet();
                Application.DoEvents();
                da.Fill(ds);
                Application.DoEvents();
                DataGridView1.DataSource = ds.Tables[0];
                DataGridView1.Refresh();

                Module.SeleccionarGrilla();

                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();

                throw new Exception(", error dentro de la Función 'Filtrar_x_linea': " + ex.Message);
            }
            finally
            {
                if (da != null)
                    da.Dispose();

                if (this.DataGridView1.RowCount > 0)
                {
                    lblTotalRegistros.Text = "Articulo: " + DataGridView1.RowCount.ToString();
                    if (!TimerVerOfertas.Enabled)
                        TimerVerOfertas.Enabled = true;

                    SimulateClickOnDataGridViewCell(DataGridView1, 0, 0);
                    SimulateClickOnDataGridViewCell(DataGridView1, 1, 0);
                }
            }
        }


        public void SimulateClickOnDataGridViewCell(DataGridView dgv, int rowIndex, int columnIndex)
        {
            try
            {
                // Verifica que los índices sean válidos
                if (rowIndex < 0 || rowIndex >= dgv.RowCount || columnIndex < 0 || columnIndex >= dgv.ColumnCount)
                {
                    throw new ArgumentOutOfRangeException("Índices de fila o columna fuera del rango.");
                }

                // Crea el evento de celda y llama al método OnCellClick
                var cellEventArgs = new DataGridViewCellEventArgs(columnIndex, rowIndex);
                dgv.CurrentCell = dgv[columnIndex, rowIndex];

                if (dgv.Name == "DataGridView1")
                {
                    llamaSimulaClick = true;
                    DataGridView1_CellClick(dgv, cellEventArgs);
                }
                else
                {
                    DataGridView2_CellClick(dgv, cellEventArgs);
                }
            }
            catch (Exception)
            {
                // Manejo silencioso de excepciones como en el código original
            }
            finally
            {
                llamaSimulaClick = false;
            }
        }
        private void TimerVerOfertas_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!terminoTadControlSelecting)
                {
                    ColorearGrillaxOfertas(DataGridView1);
                    //SeleccionarGrilla(); // ver que paso, sino lo comento por las dudas.
                    DataGridView1.Enabled = true;
                    TimerVerOfertas.Enabled = false;
                }
            }
            catch (Exception)
            {
                DataGridView1.Enabled = true;
                TimerVerOfertas.Enabled = false;
            }
            finally
            {
                TimerVerOfertas.Enabled = false;
                Application.DoEvents();
                // Commented out original code block as in VB.NET version
                /*
                if (terminoTadControlSelecting)
                {
                    if (DataGridView1.DefaultCellStyle.SelectionBackColor != Color.Transparent)
                    {
                        DataGridView1.DefaultCellStyle.SelectionBackColor = Color.Transparent;
                        DataGridView1.DefaultCellStyle.SelectionForeColor = Color.Transparent;
                        DataGridView1.Refresh();
                        Application.DoEvents();
                        int NumFila = DataGridView1.CurrentCell.RowIndex;
                        if (NumFila == 0)
                        {
                            DataGridView1.Focus();
                            SimulateClickOnDataGridViewCell(DataGridView1, 0, 0);
                            SimulateClickOnDataGridViewCell(DataGridView1, NumFila, 0);
                        }
                        else
                        {
                            SimulateClickOnDataGridViewCell(DataGridView1, NumFila - 1, 0);
                            SimulateClickOnDataGridViewCell(DataGridView1, NumFila, 0);
                        }
                        Application.DoEvents();
                    }
                }
                */
            }
        }
        private void ColorearGrillaxOfertas(DataGridView dgv)
        {
            Color customColorGreen = Color.FromArgb(118, 250, 111);
            try
            {
                if (!colorearGrillaenEjecucion)
                {
                    colorearGrillaenEjecucion = true;
                    if (dgv.Name == "DataGridView1")
                    {
                        for (int w = 0; w < dgv.Rows.Count; w++)
                        {
                            if (dgv["artt", w].Value != null)
                            {
                                if (Convert.ToInt32(dgv["artt", w].Value) > 0)
                                {
                                    dgv.Rows[w].DefaultCellStyle.BackColor = customColorGreen;
                                    dgv.Rows[w].DefaultCellStyle.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dgv.Rows[w].DefaultCellStyle.BackColor = Color.White;
                                    dgv.Rows[w].DefaultCellStyle.ForeColor = Color.Black;
                                }
                            }
                            else
                            {
                                dgv.Rows[w].DefaultCellStyle.BackColor = Color.White;
                                dgv.Rows[w].DefaultCellStyle.ForeColor = Color.Black;
                            }

                            // Semáforo de Stock
                            if (dgv["stock", w].Value != null)
                            {
                                int stockValue = Convert.ToInt32(dgv["stock", w].Value);
                                if (stockValue >= Module.ini_cond_verde)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Green;
                                    dgv["stock", w].Style.ForeColor = Color.Green;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Green;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Green;
                                }
                                else if (stockValue < Module.ini_cond_verde && stockValue > Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Yellow;
                                    dgv["stock", w].Style.ForeColor = Color.Yellow;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Yellow;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Yellow;
                                }
                                else if (stockValue <= Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Red;
                                    dgv["stock", w].Style.ForeColor = Color.Red;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Red;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Red;
                                }
                            }
                            else
                            {
                                dgv["stock", w].Style.BackColor = Color.Red;
                                dgv["stock", w].Style.ForeColor = Color.White;
                            }
                        }
                    }
                    else
                    {
                        for (int w = 0; w < dgv.Rows.Count; w++)
                        {
                            if (dgv["artt", w].Value != null)
                            {
                                if (Convert.ToInt32(dgv["artt", w].Value) > 0)
                                {
                                    dgv.Rows[w].DefaultCellStyle.BackColor = customColorGreen;
                                    dgv.Rows[w].DefaultCellStyle.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dgv.Rows[w].DefaultCellStyle.BackColor = Color.White;
                                    dgv.Rows[w].DefaultCellStyle.ForeColor = Color.Black;
                                }
                            }
                            else
                            {
                                dgv.Rows[w].DefaultCellStyle.BackColor = Color.White;
                                dgv.Rows[w].DefaultCellStyle.ForeColor = Color.Black;
                            }
                            // Semáforo de Stock
                            if (dgv["stock", w].Value != null)
                            {
                                int stockValue = Convert.ToInt32(dgv["stock", w].Value);
                                if (stockValue >= Module.ini_cond_verde)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Green;
                                    dgv["stock", w].Style.ForeColor = Color.Green;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Green;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Green;
                                }
                                else if (stockValue < Module.ini_cond_verde && stockValue > Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Yellow;
                                    dgv["stock", w].Style.ForeColor = Color.Yellow;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Yellow;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Yellow;
                                }
                                else if (stockValue <= Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Red;
                                    dgv["stock", w].Style.ForeColor = Color.Red;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionBackColor = Color.Red;
                                    dgv.Columns["Stock"].DefaultCellStyle.SelectionForeColor = Color.Red;
                                }
                            }
                            else
                            {
                                dgv["stock", w].Style.BackColor = Color.Red;
                                dgv["stock", w].Style.ForeColor = Color.White;
                            }
                        }
                    }

                    if (dgv.Rows.Count == 1)
                    {
                        // Suponemos que el DataGridView está enlazado a un DataTable
                        var tabla = (DataTable)dgv.DataSource;

                        // Crear una nueva fila
                        var nuevaFila = tabla.NewRow();

                        // Agregar la nueva fila a la tabla
                        tabla.Rows.Add(nuevaFila);

                        // Refrescar el DataGridView
                        dgv.Refresh();

                        // Obtener la fila y establecer su altura
                        var fila = dgv.Rows[dgv.Rows.Count - 1];
                        fila.Height = 1; // En C# usamos 1 en lugar de 0.1 ya que Height es int
                    }
                }
            }
            catch (Exception ex)
            {
                dgv.Enabled = true;
                colorearGrillaenEjecucion = false;
            }
            finally
            {
                colorearGrillaenEjecucion = false;
            }
        }

        private void txtModelo_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Application.DoEvents();
                if (!string.IsNullOrEmpty(txtModelo.Text))
                {
                    Module.campo = "modelo";
                    Module.valor = txtModelo.Text;
                    if (!ckbOfertas.Checked)
                    {
                        Module.consultas_Relacionales();
                    }
                    else
                    {
                        Module.consultas_Ofertas();
                    }
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Modelo': {ex.Message}");
            }
        }

        private void TimerVerOfertasEquivalencias_Tick(object sender, EventArgs e)
        {
            try
            {
                ColorearGrillaxOfertas(DataGridView2);
                // 'SeleccionarGrilla()' ver que paso, sino lo comento por las dudas.
                DataGridView2.Enabled = true;
                TimerVerOfertasEquivalencias.Enabled = false;
            }
            catch (Exception ex)
            {
                DataGridView2.Enabled = true;
                TimerVerOfertasEquivalencias.Enabled = false;
            }
            finally
            {
                // Simula un clic en la celda (fila 0, columna 0) al cargar el formulario
                DataGridView2.Refresh();
                //SimulateClickOnDataGridViewCell(DataGridView2, 1, 0);
                //SimulateClickOnDataGridViewCell(DataGridView2, 1, 0);
            }
        }

        private void DataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (this.DataGridView1.RowCount > 0)
                {
                    Module.AccionGrilla(DataGridView1);
                }
                else
                {
                    // Borrar txt precios
                    //lblCodigoSelec.Text = "";
                    lblTotalRegistros.Text = "";
                    txtPreciolista.Text = "";
                    txtDescuento1.Text = "";
                    txtDescuento2.Text = "";
                    txtDescuento3.Text = "";
                    txtPreciopublico.Text = "";
                    txtBeneficio.Text = "";
                    txtSugerido.Text = "";
                    ptbImagen.Image = System.Drawing.Image.FromFile(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\imagenes\SINFOTO.png");
                }

                if (!string.IsNullOrEmpty(this.txtPreciopublico.Text))
                {
                    int NumFila = this.DataGridView1.CurrentCell.RowIndex;
                    string Cantidad = Microsoft.VisualBasic.Interaction.InputBox("Ingrese la cantidad Solicitada",
                        "Agregar Pieza " + this.DataGridView1["CODIGO", NumFila].Value.ToString() +
                        " Marca " + this.DataGridView1["Marca", NumFila].Value.ToString(), "1");

                    string StringSelect = "Select ARFLIA, ARPROV from listas " +
                                        " WHERE ARTERC='" + this.DataGridView1["CODIGO", NumFila].Value.ToString() +
                                        "' AND ARMARC='" + this.DataGridView1["Marca", NumFila].Value.ToString() + "'";

                    //"' AND ARPRE2='" + this.DataGridView1[4, NumFila].Value.ToString() + "'"
                    string[] ArraySimple = Module.TraerCamposArticulo(StringSelect);

                    if (ArraySimple.Length == 2)
                    {
                        ModuloPreparac.WHERE_PREPARAC = " Where a.ARTERC='" + this.DataGridView1["CODIGO", NumFila].Value.ToString() +
                                        "' AND a.ARPROV=" + ArraySimple[1].ToString() +
                                        " AND a.ARFLIA=" + ArraySimple[0].ToString();

                        if (decimal.TryParse(Cantidad, out _))
                        {
                            ModuloPreparac.CANT_PREPARAC = Convert.ToDecimal(Cantidad);
                            //############################ INICIO DEFINITIVO ###############################
                            if (!ModuloPreparac.ExisteARTICULOPedido(this.DataGridView1["CODIGO", NumFila].Value.ToString(),
                                                    ArraySimple[0].ToString(),
                                                    ArraySimple[1].ToString())) //no existe en el carro de compras
                            {
                                //############################ INICIO INSERT PREPARAC ###########################
                                ModuloPreparac.BuscarArticuloParaPreparacAsync("INSERT");
                                //############################ FIN INSERT PREPARAC ##############################
                            }
                            else
                            {
                                //############################ INICIO UPDATE PREPARAC ###########################
                                //if (MessageBox.Show("¿Desea cambiar la cantidad, por la actual?",
                                //    "USTED YA PIDIÓ ESTE ARTICULO...",
                                //    MessageBoxButtons.YesNo,
                                //    MessageBoxIcon.Question) == DialogResult.Yes)
                                if (frmMensajePersonalizado.Show("¿Desea Sumar " + Cantidad + "?", "USTED YA PIDIÓ (" + ModuloPreparac.CantExistente.ToString() + " unidades de) ESTE ARTICULO...", Color.White, Color.Red) == 1)
                                {
                                    ModuloPreparac.CANT_PREPARAC = Convert.ToDecimal(Cantidad) + Convert.ToDecimal(ModuloPreparac.CantExistente);
                                    ModuloPreparac.BuscarArticuloParaPreparacAsync("UPDATE");
                                }
                                //############################ FIN UPDATE PREPARAC ##############################
                            }
                        }
                        else
                        {
                            MessageBox.Show("Debe ingresar valores númericos", "Dato no Valida",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se pudo ingresar el Articulo al carrito", "Datos Faltantes",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Debe cargar sus descuentos, Menu:Configuración-->Usuario",
                        "Acción denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError(" - " + DateTime.Now.ToLongDateString() + " => " +
                            DateTime.Now.ToString("hh:mm:ss") +
                            " 'Doble Click Grilla': " + ex.Message);
            }
        }

        private void btnWhatsapp_Click(object sender, EventArgs e)
        {
            // Número de teléfono en formato internacional (sin espacios ni caracteres especiales)
            string phoneNumber = "+5491165679428"; // Asegúrate de incluir el código de país si es necesario
                                                   // Mensaje que deseas enviar
            string message = "Hola, este es un mensaje desde PalomarNet."; // Personaliza tu mensaje aquí

            // Codificar el mensaje para que sea seguro para la URL
            string encodedMessage = Uri.EscapeDataString(message);

            // Crear el enlace de WhatsApp
            string whatsappUrl = $"https://wa.me/{phoneNumber}?text={encodedMessage}";

            // Abrir el enlace en el navegador predeterminado
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = whatsappUrl,
                UseShellExecute = true
            });
        }

        private void DataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (TimerVerOfertas.Enabled == false)
                {
                    TimerVerOfertas.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                TimerVerOfertas.Enabled = false;
            }
        }

        private void DataGridView2_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (TimerVerOfertasEquivalencias.Enabled == false)
                {
                    TimerVerOfertasEquivalencias.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                TimerVerOfertasEquivalencias.Enabled = false;
            }
        }

        private void DataGridView2_Click(object sender, EventArgs e)
        {
            try
            {
                if (DataGridView2.CurrentRow != null)
                {
                    Module.AccionGrilla(DataGridView2);
                    //CargarEquivalencias();

                    int numFila = DataGridView2.CurrentCell.RowIndex;
                    string StringSelect = "SELECT TOP 1 IIF(R.RBDESC IS NULL, '', R.RBDESC) AS RBDESC, IIF(A.ARORIG IS NULL, '', A.ARORIG) AS ARORIG" +
                        ", IIF(A.AROBSE IS NULL, '', A.AROBSE) AS AROBSE " +
                        ", IIF(A.ARUBIC1 IS NULL, '', A.ARUBIC1) AS ARUBIC1 " +
                    " from listas AS A " +
                    " INNER JOIN RUBROS AS R ON A.ARRUB2 = R.RBCODI  " +
                    " WHERE A.ARTERC='" + this.DataGridView2["CODIGO", numFila].Value.ToString() +
                    "' AND A.ARMARC='" + this.DataGridView2["Marca", numFila].Value.ToString() + "'";
                    Module.CargarDatosExtras(StringSelect);
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Click Grilla': {ex.Message}");
            }
        }

        private void DataGridView2_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (this.DataGridView2.RowCount > 0)
                {
                    Module.AccionGrilla(DataGridView2);
                }
                else
                {
                    // Borrar txt precios
                    //lblCodigoSelec.Text = "";
                    lblTotalRegistros.Text = "";
                    txtPreciolista.Text = "";
                    txtDescuento1.Text = "";
                    txtDescuento2.Text = "";
                    txtDescuento3.Text = "";
                    txtPreciopublico.Text = "";
                    txtBeneficio.Text = "";
                    txtSugerido.Text = "";
                    ptbImagen.Image = System.Drawing.Image.FromFile(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\imagenes\SINFOTO.png");
                }

                if (!string.IsNullOrEmpty(this.txtPreciopublico.Text))
                {
                    int NumFila = this.DataGridView2.CurrentCell.RowIndex;
                    string Cantidad = Microsoft.VisualBasic.Interaction.InputBox("Ingrese la cantidad Solicitada",
                        "Agregar Pieza " + this.DataGridView2["CODIGO", NumFila].Value.ToString() +
                        " Marca " + this.DataGridView2["Marca", NumFila].Value.ToString(), "1");

                    string StringSelect = "Select ARFLIA, ARPROV from listas " +
                                        " WHERE ARTERC='" + this.DataGridView2["CODIGO", NumFila].Value.ToString() +
                                        "' AND ARMARC='" + this.DataGridView2["Marca", NumFila].Value.ToString() + "'";

                    //"' AND ARPRE2='" + this.DataGridView2[4, NumFila].Value.ToString() + "'"
                    string[] ArraySimple = Module.TraerCamposArticulo(StringSelect);

                    if (ArraySimple.Length == 2)
                    {
                        ModuloPreparac.WHERE_PREPARAC = " Where a.ARTERC='" + this.DataGridView2["CODIGO", NumFila].Value.ToString() +
                                        "' AND a.ARPROV=" + ArraySimple[1].ToString() +
                                        " AND a.ARFLIA=" + ArraySimple[0].ToString();

                        if (decimal.TryParse(Cantidad, out _))
                        {
                            ModuloPreparac.CANT_PREPARAC = Convert.ToDecimal(Cantidad);
                            //############################ INICIO DEFINITIVO ###############################
                            if (!ModuloPreparac.ExisteARTICULOPedido(this.DataGridView2["CODIGO", NumFila].Value.ToString(),
                                                    ArraySimple[0].ToString(),
                                                    ArraySimple[1].ToString())) //no existe en el carro de compras
                            {
                                //############################ INICIO INSERT PREPARAC ###########################
                                ModuloPreparac.BuscarArticuloParaPreparacAsync("INSERT");
                                //############################ FIN INSERT PREPARAC ##############################
                            }
                            else
                            {
                                //############################ INICIO UPDATE PREPARAC ###########################
                                //if (MessageBox.Show("¿Desea cambiar la cantidad, por la actual?",
                                //    "USTED YA PIDIÓ ESTE ARTICULO...",
                                //    MessageBoxButtons.YesNo,
                                //    MessageBoxIcon.Question) == DialogResult.Yes)
                                if (frmMensajePersonalizado.Show("¿Desea Sumar " + Cantidad + "?", "USTED YA PIDIÓ (" + ModuloPreparac.CantExistente.ToString() + " unidades de) ESTE ARTICULO...", Color.White, Color.Red) == 1)
                                {
                                    ModuloPreparac.CANT_PREPARAC = Convert.ToDecimal(Cantidad) + Convert.ToDecimal(ModuloPreparac.CantExistente);
                                    ModuloPreparac.BuscarArticuloParaPreparacAsync("UPDATE");
                                }
                                //############################ FIN UPDATE PREPARAC ##############################
                            }
                        }
                        else
                        {
                            MessageBox.Show("Debe ingresar valores númericos", "Dato no Valida",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se pudo ingresar el Articulo al carrito", "Datos Faltantes",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Debe cargar sus descuentos, Menu:Configuración-->Usuario",
                        "Acción denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError(" - " + DateTime.Now.ToLongDateString() + " => " +
                            DateTime.Now.ToString("hh:mm:ss") +
                            " 'Doble Click Grilla': " + ex.Message);
            }
        }


        //private void PictureBox3_Paint(object sender, PaintEventArgs e)
        //{
        //    if (PictureBox3.Image != null)
        //    {
        //        // Crear una matriz de color para ajustar la opacidad (0.5 = 50% opacidad)
        //        float opacity = 0.5f; // Ajusta este valor (0.0 a 1.0)

        //        ColorMatrix matrix = new ColorMatrix();
        //        matrix.Matrix33 = opacity; // Componente alfa (transparencia)

        //        ImageAttributes attributes = new ImageAttributes();
        //        attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

        //        // Dibujar la imagen con la opacidad aplicada
        //        e.Graphics.DrawImage(
        //            PictureBox3.Image,
        //            new System.Drawing.Rectangle(0, 0, PictureBox3.Width, PictureBox3.Height),
        //            0, 0, PictureBox3.Image.Width, PictureBox3.Image.Height,
        //            GraphicsUnit.Pixel,
        //            attributes);
        //    }
        //}
    }
}

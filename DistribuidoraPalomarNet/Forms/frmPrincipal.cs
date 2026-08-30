using DistribuidoraPalomarNet.Forms;
using ImageViewerApp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.VisualBasic.FileIO;
using System;
using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Image = System.Drawing.Image;


namespace DistribuidoraPalomarNet
{
    public partial class frmPrincipal : Form
    {
        private static Image _iconoOferta;
        private bool _mensajeClienteMostrado = false;
        /// <summary>Bandera para evitar disparos en cascada de filtros cuando se recargan combos dependientes.</summary>
        public bool _aplicandoFiltros = false;

        public void ApplyEstadoClienteToUI()
        {
            bool autorizado = Module.ClienteAutorizado;

            if (autorizado)
            {
                _mensajeClienteMostrado = false;
            }

            if (DataGridView1 != null) DataGridView1.Enabled = autorizado;
            if (DataGridView2 != null) DataGridView2.Enabled = autorizado;

            if (txtCodigo != null) txtCodigo.Enabled = autorizado;
            if (txtDescripcion != null) txtDescripcion.Enabled = autorizado;
            if (txtModelo != null) txtModelo.Enabled = autorizado;
            if (cmbArmadora != null) cmbArmadora.Enabled = autorizado;
            if (cmbLineas != null) cmbLineas.Enabled = autorizado;
            if (btnLimpiar != null) btnLimpiar.Enabled = autorizado;
            if (ckbOfertas != null) ckbOfertas.Enabled = autorizado;

            if (btnGenerarTxt != null) btnGenerarTxt.Enabled = autorizado;

            if (btnVerCarrito != null) btnVerCarrito.Enabled = autorizado;
            if (btnActualizar != null) btnActualizar.Enabled = true;

            if (!autorizado && !_mensajeClienteMostrado)
            {
                string msg = string.IsNullOrWhiteSpace(Module.ClienteMensajeValidacion)
                    ? "Datos inválidos"
                    : Module.ClienteMensajeValidacion;
                MessageBox.Show(msg, "Acceso restringido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _mensajeClienteMostrado = true;
            }
        }

        private void DataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!Module.ClienteAutorizado) return;
            Application.DoEvents();
            //DataGridView2.DefaultCellStyle.SelectionBackColor = Color.Transparent;
            //DataGridView2.DefaultCellStyle.SelectionForeColor = Color.Transparent;
            ColorearCelda(DataGridView2);
        }


        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!Module.ClienteAutorizado) return;
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

        private void Combo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back)
            {
                if (DateTime.Now - comboLastKeyTime > TimeSpan.FromSeconds(1))
                    comboSearchBuffer = string.Empty;
                else if (comboSearchBuffer.Length > 0)
                    comboSearchBuffer = comboSearchBuffer.Substring(0, comboSearchBuffer.Length - 1);
                comboLastKeyTime = DateTime.Now;
                var cb = sender as ComboBox;
                if (cb != null)
                {
                    int idx = FindIndex(cb, comboSearchBuffer);
                    if (idx >= 0)
                        cb.SelectedIndex = idx;
                    cb.DroppedDown = true;
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                comboSearchBuffer = string.Empty;
                var cb = sender as ComboBox;
                if (cb != null) cb.DroppedDown = false;
                e.Handled = true;
            }
        }

        private void Combo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar))
            {
                if (DateTime.Now - comboLastKeyTime > TimeSpan.FromSeconds(1))
                    comboSearchBuffer = string.Empty;
                comboLastKeyTime = DateTime.Now;
                comboSearchBuffer += e.KeyChar;
                var cb = sender as ComboBox;
                if (cb != null)
                {
                    int idx = FindIndex(cb, comboSearchBuffer);
                    if (idx >= 0)
                        cb.SelectedIndex = idx;
                    cb.DroppedDown = true;
                }
                e.Handled = true;
            }
        }

        private int FindIndex(ComboBox cb, string buffer)
        {
            if (string.IsNullOrEmpty(buffer)) return -1;
            string buf = buffer.ToLowerInvariant();
            string displayMember = cb.DisplayMember;
            for (int i = 0; i < cb.Items.Count; i++)
            {
                object item = cb.Items[i];
                string text = string.Empty;
                if (item != null)
                {
                    if (!string.IsNullOrEmpty(displayMember))
                    {
                        if (item is System.Data.DataRowView drv)
                        {
                            try { text = Convert.ToString(drv[displayMember]); }
                            catch { text = item.ToString(); }
                        }
                        else
                        {
                            var prop = item.GetType().GetProperty(displayMember);
                            text = prop != null ? Convert.ToString(prop.GetValue(item, null)) : item.ToString();
                        }
                    }
                    else
                    {
                        text = item.ToString();
                    }
                }
                if (!string.IsNullOrEmpty(text))
                {
                    string cmp = text.ToLowerInvariant();
                    if (cmp.StartsWith(buf) || cmp.Contains(buf))
                        return i;
                }
            }
            return -1;
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
        private string comboSearchBuffer = string.Empty;
        private DateTime comboLastKeyTime = DateTime.MinValue;

        private static frmPrincipal _instance;
        private bool _porLimpieza = false;
        private bool _webView2Inicializado = false;
        private bool _webView22Inicializado = false;
        private string _textoMarquesinaPendiente = string.Empty;

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
            DataGridView1.CellPainting += DataGridView_CellPainting_OfertaIcon;
            DataGridView2.CellPainting += DataGridView_CellPainting_OfertaIcon;
            InitializeMenuEvents();

            // Asignar manejadores de eventos para botones
            btnGenerarTxt.Click += BtnGenerarTxt_Click;
            btnEquivalencia.Click += BtnEquivalencia_Click;
            //btnConfUsu.Click += BtnConfUsu_Click;
            btncerrar.Click += Btncerrar_Click;
            btnVerCarrito.Click += BtnVerCarrito_Click;
            btnLimpiar.Click += BtnLimpiar_Click;
            btnnovedades.Click += Btnnovedades_Click;
            ckbOfertas.Click += ckbOfertas_Click;
            cmbLineas.KeyPress += Combo_KeyPress;
            cmbArmadora.KeyPress += Combo_KeyPress;
            cmbLineas.KeyDown += Combo_KeyDown;
            cmbArmadora.KeyDown += Combo_KeyDown;


            InitializeWebView2Async();
        }

        private async void InitializeWebView2Async()
        {
            try
            {
                webView21.DefaultBackgroundColor = System.Drawing.Color.Black;
                await webView21.EnsureCoreWebView2Async(null);
                webView21.CoreWebView2.Settings.IsStatusBarEnabled = false;
                webView21.CoreWebView2.Settings.AreDevToolsEnabled = false;
                webView21.CoreWebView2.Settings.IsZoomControlEnabled = false;
                webView21.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
                _webView2Inicializado = true;
                CargarMarquesina();

                if (!string.IsNullOrWhiteSpace(_textoMarquesinaPendiente))
                {
                    var txt = _textoMarquesinaPendiente;
                    _textoMarquesinaPendiente = string.Empty;
                    ActualizarTexto(txt);
                }
            }
            catch
            {
                _webView2Inicializado = false;
            }

            try
            {
                await webView22.EnsureCoreWebView2Async(null);
                webView22.CoreWebView2.Settings.IsStatusBarEnabled = false;
                webView22.CoreWebView2.Settings.AreDevToolsEnabled = false;
                webView22.CoreWebView2.Settings.IsZoomControlEnabled = false;
                webView22.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
                _webView22Inicializado = true;
            }
            catch
            {
                _webView22Inicializado = false;
            }
        }

        private void CargarMarquesina()
        {
            string textoLargo = "📢 AVISO 1: Oferta especial esta semana... || " +
                                "📢 AVISO 2: Pronto tendra Novedades Especiales...";

            string htmlContent = $@"
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset='UTF-8'>
            <meta name='color-scheme' content='dark'>
            <style>
                html, body {{
                    margin: 0;
                    padding: 0;
                    overflow: hidden;
                    background-color: #000000 !important;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    font-size: 14pt;
                    color: #ffffff !important;
                    height: 100%;
                    display: flex;
                    align-items: center;
                }}
                .marquee {{
                    white-space: nowrap;
                    display: inline-block;
                    animation: scroll 35s linear infinite;
                    padding-left: 100%;
                    color: #ffffff;
                }}
                @keyframes scroll {{
                    0% {{ transform: translateX(0); }}
                    100% {{ transform: translateX(-100%); }}
                }}
            </style>
        </head>
        <body>
            <div class='marquee'>{textoLargo}</div>
        </body>
        </html>";

            webView21.NavigateToString(htmlContent);
        }

        // Método público para actualizar el texto sin recargar todo
        public void ActualizarTexto(string nuevoTexto)
        {
            if (!_webView2Inicializado)
            {
                _textoMarquesinaPendiente = nuevoTexto ?? string.Empty;
                return;
            }

            // Verificar si estamos en el hilo UI
            if (InvokeRequired)
            {
                Invoke(new Action<string>(ActualizarTexto), nuevoTexto);
                return;
            }

            try
            {
                // Verificación adicional de seguridad
                if (webView21?.CoreWebView2 == null)
                {
                    _textoMarquesinaPendiente = nuevoTexto ?? string.Empty;
                    _webView2Inicializado = false; // Marcar como no inicializado si CoreWebView2 es null
                    return;
                }

                string script = $@"
                var marqueeDiv = document.querySelector('.marquee');
                if (marqueeDiv) {{
                    marqueeDiv.innerHTML = '{nuevoTexto.Replace("'", @"\'")}';
                }}
            ";
                webView21.CoreWebView2.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                // Log del error para debugging (opcional)
                System.Diagnostics.Debug.WriteLine($"Error en ActualizarTexto: {ex.Message}");
                _textoMarquesinaPendiente = nuevoTexto ?? string.Empty;
                _webView2Inicializado = false; // Marcar como no inicializado en caso de error
            }
        }

        private static Image GetIconoOferta()
        {
            if (_iconoOferta != null) return _iconoOferta;

            try
            {
                string path = Path.Combine(AppContext.BaseDirectory, "Resources", "Icons", "iconoOferta.png");
                if (File.Exists(path))
                {
                    using (var img = Image.FromFile(path))
                    {
                        _iconoOferta = new Bitmap(img);
                    }
                }
            }
            catch
            {
                _iconoOferta = null;
            }

            return _iconoOferta;
        }

        private void DataGridView_CellPainting_OfertaIcon(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var dgv = sender as DataGridView;
            if (dgv == null) return;

            if (!string.Equals(dgv.Columns[e.ColumnIndex].Name, "Preciopublico", StringComparison.OrdinalIgnoreCase))
                return;

            DataGridViewRow row = dgv.Rows[e.RowIndex];
            if (row == null) return;

            object arttValue = null;
            try { arttValue = row.Cells["artt"].Value; } catch { return; }

            int artt;
            if (arttValue == null || !int.TryParse(Convert.ToString(arttValue), out artt) || artt <= 0)
                return;

            var icon = GetIconoOferta();
            if (icon == null) return;

            e.PaintBackground(e.ClipBounds, true);

            var cellBounds = e.CellBounds;
            int padding = 2;
            int iconSize = 16;

            int iconX = cellBounds.Left + padding;
            int iconY = cellBounds.Top + (cellBounds.Height - iconSize) / 2;
            var iconRect = new System.Drawing.Rectangle(iconX, iconY, iconSize, iconSize);
            e.Graphics.DrawImage(icon, iconRect);

            string text = Convert.ToString(e.FormattedValue) ?? string.Empty;
            var textRect = new System.Drawing.Rectangle(
                cellBounds.Left + iconSize + (padding * 2),
                cellBounds.Top,
                cellBounds.Width - iconSize - (padding * 3),
                cellBounds.Height);

            TextRenderer.DrawText(
                e.Graphics,
                text,
                e.CellStyle.Font,
                textRect,
                e.CellStyle.ForeColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            e.Paint(cellBounds, DataGridViewPaintParts.Border);
            e.Handled = true;
        }
        private async void Btnnovedades_Click(object sender, EventArgs e)
        {
            try
            {
                await Module.Funcion_Novedades();
                TimerVerOfertas.Enabled = true;
            }
            catch (Exception ex)
            {
                TimerVerOfertas.Enabled = false;
                MessageBox.Show($"Error dentro de la accion 'Novedades Click': {ex.Message}", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnActualizar_Click(object sender, EventArgs e)
        {
            try
            {
                Module.FechaActualizacion();
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
                //MessageBox.Show("Error en Click 'Generar TXT': " + ex.ToString(), "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtDescripcion_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    Application.DoEvents();
            //    if (!string.IsNullOrEmpty(txtDescripcion.Text))
            //    {
            //        Module.campo = "descripcion";
            //        Module.valor = txtDescripcion.Text;
            //        if (!ckbOfertas.Checked)
            //        {
            //            await Module.consultas_Relacionales();
            //        }
            //        else
            //        {
            //            await Module.consultas_Ofertas();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Descripción': {ex.Message}");
            //}
            try
            {
                timer2.Stop();
                timer2.Start();
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

                    string SelectQuery = "SELECT DISTINCT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    " a.arinte as Posición," +
                    "IIF(a.artt IS NOT NULL AND a.artt <> 0, a.arpre5, a.arpre2) as Preciopublico, " +
                    "IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from (listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) ";

                    da = new OleDbDataAdapter(
                        SelectQuery +
                        "Where a.AROBSE='" + Module.NumEqui + "' ", Module.cnn);

                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    DataGridView2.DataSource = ds.Tables[0];
                    DataGridView2.Refresh();
                    DataGridView2.ClearSelection();

                    DataGridView2.Columns["aroo"].Visible = false;
                    DataGridView2.Columns["artt"].Visible = false;
                    DataGridView2.Columns["arubic1"].Visible = false;
                    DataGridView2.Columns["arobse"].Visible = false;
                    //DataGridView2.Columns["Linea"].Visible = false;
                    //DataGridView2.Columns["Modelo"].Visible = false;
                    DataGridView2.Columns["Descripcion"].Visible = false;

                    // NUEVO: Configurar columna CODIGO como link clickeable
                    EnsureCodigoColumnLinkStyling();

                    var miEstiloStock = new DataGridViewCellStyle(DataGridView2.Columns["Stock"].DefaultCellStyle)
                    {
                        Font = new Font("Arial", 2, FontStyle.Regular)
                    };

                    DataGridView2.Columns["Stock"].DefaultCellStyle = miEstiloStock;


                    if (DataGridView1.RowCount > 0)
                        lblTotalRegistros.Text = "ARTÍCULO: " + DataGridView2.RowCount.ToString();
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
                Module.LimpiarCampos(0);
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Click Limpiar': {ex.Message}");
            }
        }

        private void InitializeMenuEvents()
        {
            // Archivo
            ActualizarBasesToolStripMenuItem.Click += new EventHandler(ActualizarBasesToolStripMenuItem_Click);
            ActualizarDatosLocalToolStripMenuItem.Click += new EventHandler(ActualizarDatosLocalToolStripMenuItem_Click);

            // Configuración
            // ClientesToolStripMenuItem.Click += new EventHandler(ClientesToolStripMenuItem_Click); // ya está conectado en el Designer
            DescPorMarcaToolStripMenuItem.Click += new EventHandler(DescPorMarcaToolStripMenuItem_Click);

            // Ver
            PedidosEnviadosToolStripMenuItem.Click += new EventHandler(PedidosEnviadosToolStripMenuItem_Click);

            // Web y Acerca de
            //WwwsistemasvscomarToolStripMenuItem.Click += new EventHandler(WwwsistemasvscomarToolStripMenuItem_Click);
            //AcercaDeToolStripMenuItem.Click += new EventHandler(AcercaDeToolStripMenuItem_Click);
        }

        private async void ActualizarBasesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Module.ActualizacionXFTP();
                await Module.consultas_Relacionales();
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

        private async void ActualizarDatosLocalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Module.UpdateValorDBLocal("");  // Pasando string vacío como parámetro
                await Module.consultas_Relacionales();
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
                // CORREGIDO: Ruta correcta con "Net"
                string rutaPedidos = @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Pedidos\";

                // NUEVO: Verificar si la carpeta existe, si no, crearla
                if (!Directory.Exists(rutaPedidos))
                {
                    Directory.CreateDirectory(rutaPedidos);
                }

                // Abrir la carpeta en el explorador de Windows
                Process.Start("explorer.exe", rutaPedidos);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir la carpeta de pedidos:\n\n" +
                    $"Ruta: C:\\DistribuidoraPalomarNet\\DistribuidoraPalomarNet\\Pedidos\\\n\n" +
                    $"Error: {ex.Message}\n\n" +
                    $"Verifique que tenga permisos para acceder a esta ubicación.",
                    "Error - Pedidos Generados", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            MessageBox.Show("Distribuidora Palomar\nVersión 1.0\n\nDesarrollado por Sistemas VS\nwww.sistemavs.com", "Acerca de", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void frmPrincipal_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Al momento de cerrar, para liberar la memoria
            Module.vectornomfoto = null;
        }

        private async void frmPrincipal_Load(object sender, EventArgs e)
        {
            try
            {
                // Carga general de combos, grillas, imágenes (banners)
                Module.LeerConfig();
                if (!string.IsNullOrWhiteSpace(Module._textoBanner))
                    ActualizarTexto(Module._textoBanner);

                // ✅ OPTIMIZACIÓN: Inicializar datos críticos primero (sin bloquear)
                Module.SincronizarDescuentosCliente();
                Module.Cargar_Usuario();
                await Module.Inicializacion();
                ApplyEstadoClienteToUI();

                // ✅ OPTIMIZACIÓN: Descargar banners en segundo plano (no bloquea la UI)
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await Module.DescargaArchivoAsync("BannerPalomar.gif",
                            Module.ini_ruta_gral + "\\DistribuidoraPalomarNet\\Imagenes\\BannerPalomar.gif",
                            Module.ini_usu_ftp, Module.ini_pass_ftp, Module.ini_ruta_ftp);

                        await Module.DescargaArchivoAsync("bannerweb.gif",
                            Module.ini_ruta_gral + "\\DistribuidoraPalomarNet\\Imagenes\\bannerweb.gif",
                            Module.ini_usu_ftp, Module.ini_pass_ftp, Module.ini_ruta_ftp);

                        // Cargar banners en el hilo UI
                        this.Invoke((Action)(() => CargarBanner()));
                    }
                    catch (Exception ex)
                    {
                        Module.GrabarError($" - {DateTime.Now:dd/MM/yyyy HH:mm:ss} 'Descarga Banners Async': {ex.Message}");
                    }
                });

                // ✅ OPTIMIZACIÓN: Verificar actualizaciones en segundo plano (no bloquea la UI)
                _ = Task.Run(async () =>
                {
                    try
                    {
                        Module.FechaActualizacion();
                        bool hayActualizacion = await Module.BuscarActualizacionDisponibleAsync();

                        if (hayActualizacion)
                        {
                            this.Invoke((Action)(() =>
                            {
                                if (MessageBox.Show("Hay una nueva actualización de la lista", "", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                                {
                                    Module.ActualizacionXFTP();
                                }
                            }));
                        }
                    }
                    catch (Exception ex)
                    {
                        Module.GrabarError($" - {DateTime.Now:dd/MM/yyyy HH:mm:ss} 'Verificación Actualización Async': {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Inicio Formulario Principal': {ex.Message}");
            }
        }

        private void CargarTextoBanner(string textoBanner)
        {
            string textoLargo = textoBanner;

            string htmlContent = $@"
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset='UTF-8'>
                    <meta name='color-scheme' content='dark'>
                    <style>
                        html, body {{
                            margin: 0;
                            padding: 0;
                            overflow: hidden;
                            background-color: #000000 !important;
                            font-family: 'Segoe UI', Arial, sans-serif;
                            font-size: 14pt;
                            color: #ffffff !important;
                            height: 100%;
                            display: flex;
                            align-items: center;
                        }}
                        .marquee {{
                            white-space: nowrap;
                            display: inline-block;
                            animation: scroll 25s linear infinite;
                            padding-left: 100%;
                            color: #ffffff;
                        }}
                        @keyframes scroll {{
                            0% {{ transform: translateX(0); }}
                            100% {{ transform: translateX(-100%); }}
                        }}
                    </style>
                </head>
                <body>
                    <div class='marquee'>{textoLargo}</div>
                </body>
                </html>";

            webView21.NavigateToString(htmlContent);
        }

        private void btncerrar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private async void DataGridView1_Click(object sender, EventArgs e)
        {
            try
            {
                if (DataGridView1.CurrentRow != null)
                {
                    await Module.AccionGrilla(DataGridView1);
                    CargarEquivalencias();

                    int numFila = DataGridView1.CurrentCell.RowIndex;
                    string StringSelect = "SELECT TOP 1 IIF(R.RBDESC IS NULL, '', R.RBDESC) AS RBDESC, IIF(A.ARORIG IS NULL, '', A.ARORIG) AS ARORIG" +
                        ", IIF(A.AROBSE IS NULL, '', A.AROBSE) AS AROBSE " +
                        ", IIF(A.ARUBIC1 IS NULL, '', A.ARUBIC1) AS ARUBIC1 " +
                    " FROM (listas AS A INNER JOIN proveedo AS p ON p.prcodi = A.arprov) " +
                    " INNER JOIN RUBROS AS R ON A.ARRUB2 = R.RBCODI " +
                    " WHERE A.ARTERC='" + this.DataGridView1["CODIGO", numFila].Value.ToString() +
                    "' AND p.prmarc='" + this.DataGridView1["Marca", numFila].Value.ToString() + "'";
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
                // Cargar banner grande si el archivo ya existe localmente
                string rutaBannerGrande = "C:\\DistribuidoraPalomarNet\\DistribuidoraPalomarNet\\Imagenes\\BannerPalomar.gif";
                if (File.Exists(rutaBannerGrande))
                {
                    imgBanner.Image = Image.FromFile(rutaBannerGrande);
                    imgBanner.SizeMode = PictureBoxSizeMode.StretchImage;
                }

                // Cargar banner botón si el archivo ya existe localmente
                string rutaBannerBoton = "C:\\DistribuidoraPalomarNet\\DistribuidoraPalomarNet\\Imagenes\\bannerweb.gif";
                if (File.Exists(rutaBannerBoton))
                {
                    PictureBox3.Image = Image.FromFile(rutaBannerBoton);
                    PictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Cargar Banner': {ex.Message}");
            }
        }

        // Evento de clic para el banner grande
        private void ImgBanner_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(Module.ini_link_banner_grande))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = Module.ini_link_banner_grande,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Click Banner Grande': {ex.Message}");
            }
        }

        // Evento de clic para PictureBox3 (banner botón)
        private void PictureBox3_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(Module.ini_link_banner_boton))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = Module.ini_link_banner_boton,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Click Banner Botón': {ex.Message}");
            }
        }

        private void CheckBox1_Click(object sender, EventArgs e)
        {
            bool visible = CheckBox1.Checked;

            // Labels
            lblPL.Visible = visible;
            lblD1.Visible = visible;
            lblD2.Visible = visible;
            lblD3Fam.Visible = visible;
            lblPC.Visible = visible;
            lblB.Visible = visible;
            lblIvaFlia.Visible = visible;
            //lblPP.Visible = visible;

            // TextBoxes
            txtPreciolista.Visible = visible;
            txtDescuento1.Visible = visible;
            txtDescuento2.Visible = visible;
            txtDescuentoFamilia.Visible = visible;
            txtPreciopublico.Visible = visible;
            txtBeneficio.Visible = visible;
            txtIvaFlia.Visible = visible;
            //txtSugerido.Visible = visible;
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

        private void txtCodigo_TextChanged(object sender, EventArgs e)
        {
            try
            {
                timer2.Stop();
                timer2.Start();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Código': {ex.Message}");
            }
        }

        private async void FiltrarPorCodigo()
        {
            try
            {
                if (!string.IsNullOrEmpty(txtCodigo.Text))
                {
                    string queryRel = ckbOfertas.Checked
                        ? "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                            "p.PRMARC as Marca," +
                            "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                            //"a.armarc as Marca," 
                            //"l.lidesc as Linea," +
                            "a.armode as Linea," +
                            " a.arinte as Posición," +
                          //" a.ARMARC as Marca2," + 
                          "a.arpre5 as Preciopublico, " +
                          "IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                          ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse from (listas as a " +
                            "INNER JOIN proveedo AS p ON p.prcodi = a.arprov) " +
                            "LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI " +
                          //" left join lineas as l on a.arcuar=l.licodi= " +
                          "where a.ARTT<>0 AND a.arprim+a.arterc like '%" + txtCodigo.Text +
                          "%' order by a.arcodi"
                        : "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                            "p.PRMARC as Marca," +
                            "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                            //"a.armarc as Marca," 
                            //"l.lidesc as Linea," +
                            "a.armode as Linea," +
                            " a.arinte as Posición," +
                            //" a.ARMARC as Marca2," + 
                            //"a.arpre2 as Preciopublico, " +//"IIF(a.artt IS NOT NULL AND a.artt <> 0, a.arpre5, a.arpre2) as Preciopublico, "+
                            "IIF(a.artt IS NOT NULL AND a.artt <> 0, a.arpre5, a.arpre2) as Preciopublico, " +
                          " IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                          ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse from (listas as a " +
                            "INNER JOIN proveedo AS p ON p.prcodi = a.arprov) " +
                            "LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI " +
                          //" left join lineas as l on a.arcuar=l.licodi= " +
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
                            Module.Definir_Grilla();
                        }
                        if (DataGridView1.RowCount > 0)
                            await Module.SeleccionarGrilla();
                        if (Module.cnn.State == ConnectionState.Open)
                            Module.cnn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                if (Module.cnn.State == ConnectionState.Open)
                    Module.cnn.Close();
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Código': {ex.Message}");
            }
            finally
            {
                if (TimerVerOfertas.Enabled == false) TimerVerOfertas.Enabled = true;
            }
        }

        public async void RefrescarSegunFiltros()
        {
            bool hayFiltroLinea = cmbLineas.SelectedValue is int lineaVal && lineaVal != 0;
            bool hayFiltroCodigo = !string.IsNullOrWhiteSpace(txtCodigo.Text);
            bool hayFiltroDescModelo = !string.IsNullOrWhiteSpace(txtDescripcion.Text) || !string.IsNullOrWhiteSpace(txtModelo.Text);
            bool hayFiltroCombo =
                (!string.IsNullOrWhiteSpace(cmbArmadora.Text) && cmbArmadora.Text.Trim() != "TODOS") ||
                (!string.IsNullOrWhiteSpace(cmbRubros.Text) && cmbRubros.Text.Trim() != "TODOS") ||
                (!string.IsNullOrWhiteSpace(cmbMarcas.Text) && cmbMarcas.Text.Trim() != "TODOS");

            if (hayFiltroLinea)
            {
                Filtrar_x_linea(cmbLineas.Text);
            }
            else if (hayFiltroCodigo)
            {
                FiltrarPorCodigo();
            }
            else if (ckbOfertas.Checked)
            {
                await Module.consultas_Ofertas();
            }
            else if (hayFiltroDescModelo || hayFiltroCombo)
            {
                await Module.consultas_Relacionales();
            }
            else
            {
                Module.funcion_inicial(1);
            }
        }

        private async void ckbOfertas_Click(object sender, EventArgs e)
        {
            try
            {
                // Limpio todo los campos
                // filtro las ofertas de articulo
                if (ckbOfertas.Checked)
                {
                    await Module.consultas_Ofertas();
                    //Pintar_Ofertas();
                    if (TimerVerOfertas.Enabled == false) TimerVerOfertas.Enabled = true;
                }
                else
                {
                    if (!Migra)
                    {
                        await Module.consultas_Relacionales();
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

        private void cmbArmadora_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_aplicandoFiltros) return;
            try
            {
                if (!string.IsNullOrEmpty(cmbArmadora.Text) && cmbArmadora.Text != "System.Data.DataRowView")
                {
                    if (cmbArmadora.Text == " TODOS" || cmbArmadora.Text == "TODOS")
                    {
                        cmbArmadora.Text = string.Empty;
                    }
                    _aplicandoFiltros = true;
                    Module.CargarComboLineas(cmbArmadora.Text);
                    _aplicandoFiltros = false;
                    // Filtrar sobre datos ya cargados en memoria (sin consultar la BD)
                    Module.AplicarFiltrosLocales();
                }
            }
            catch (Exception ex)
            {
                _aplicandoFiltros = false;
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Armadora': {ex.Message}");
            }
        }

        private void cmbMarcas_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_aplicandoFiltros) return;
            try
            {
                // Filtrar sobre datos ya cargados en memoria (sin consultar la BD)
                Module.AplicarFiltrosLocales();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Marca': {ex.Message}");
            }
        }

        private async void cmbRubros_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_aplicandoFiltros) return;
            try
            {
                string rubrosSel = cmbRubros.Text?.Trim() ?? string.Empty;

                // Actualizar las marcas basadas en el rubro seleccionado
                _aplicandoFiltros = true;
                Module.CargarComboMarcasPorRubro(rubrosSel);
                _aplicandoFiltros = false;

                // Filtrar via consulta SQL a la BD (el rubro se filtra con JOIN RUBROS)
                if (!ckbOfertas.Checked)
                    await Module.consultas_Relacionales();
                else
                    await Module.consultas_Ofertas();
            }
            catch (Exception ex)
            {
                _aplicandoFiltros = false;
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Rubro': {ex.Message}");
            }
        }

        private async void txtEquivalencia_KeyPress(object sender, KeyPressEventArgs e)
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
                            await Module.consultas_Relacionales();
                        }
                        else
                        {
                            await Module.consultas_Ofertas();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Equivalencia': {ex.Message}");
                }
            }
        }

        /// <summary>Muestra la barra de advertencia de descuentos no configurados en webView22.</summary>
        public void MostrarNotificacionDescuento()
        {
            if (!_webView22Inicializado) return;
            try
            {
                string html = @"<!DOCTYPE html>
<html><head><style>
*{margin:0;padding:0;box-sizing:border-box}
body{
  font-family:'Segoe UI',system-ui,sans-serif;
  background:linear-gradient(90deg,#78350f 0%,#b45309 45%,#d97706 100%);
  height:100vh;display:flex;align-items:center;
  padding:0 14px;gap:9px;overflow:hidden;
}
.dot{width:7px;height:7px;border-radius:50%;background:#fcd34d;flex-shrink:0;
  animation:p 1.5s ease-in-out infinite}
@keyframes p{0%,100%{opacity:1}50%{opacity:.25}}
.ic{flex-shrink:0}
.msg{font-size:11.5px;color:#fff;font-weight:600;flex:1;letter-spacing:.1px}
.badge{
  background:rgba(0,0,0,.28);border-radius:4px;
  padding:2px 8px;font-size:10px;color:#fde68a;font-weight:700;flex-shrink:0;
  white-space:nowrap;
}
</style></head><body>
<div class='dot'></div>
<svg class='ic' width='14' height='14' viewBox='0 0 24 24' fill='none'
  stroke='#fcd34d' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'>
  <path d='M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z'/>
  <line x1='12' y1='9' x2='12' y2='13'/><line x1='12' y1='17' x2='12.01' y2='17'/>
</svg>
<div class='msg'>Sin descuentos configurados &mdash; Pida sus descuentos a la distribuidora</div>
<div class='badge'>DESDESC1 = 0</div>
</body></html>";
                webView22.NavigateToString(html);
                webView22.Visible = true;
            }
            catch { }
        }

        /// <summary>Oculta la barra de advertencia de descuentos en webView22.</summary>
        public void OcultarNotificacionDescuento()
        {
            if (!_webView22Inicializado) return;
            try
            {
                if (webView22.Visible)
                {
                    webView22.NavigateToString("<html><body style='margin:0;background:transparent'></body></html>");
                    webView22.Visible = false;
                }
            }
            catch { }
        }

        private async void Update_Precios_Ofertas()
        {
            try
            {
                if (!Module.DescuentosConfigurados())
                {
                    MostrarNotificacionDescuento();
                    MessageBox.Show(
                        "Descuentos no configurados (DESDESC1 = 0).\nPida sus descuentos a la distribuidora.",
                        "Sin Descuentos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                Module.Update_PrecioCosto_Ofertas();
                Module.Update_PrecioSug_Ofertas();
                await Module.consultas_Ofertas();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Update Precios Ofertas': {ex.Message}");
            }
        }

        private async void Update_Precios()
        {
            try
            {
                if (!Module.DescuentosConfigurados())
                {
                    MostrarNotificacionDescuento();
                    MessageBox.Show(
                        "Descuentos no configurados (DESDESC1 = 0).\nPida sus descuentos a la distribuidora.",
                        "Sin Descuentos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                Module.Update_PrecioCosto();
                Module.Update_PrecioSug();
                await Module.consultas_Relacionales();
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
            if (_aplicandoFiltros) return;
            try
            {
                // Filtrar sobre datos ya cargados en memoria (sin consultar la BD)
                Module.AplicarFiltrosLocales();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Líneas': {ex.Message}");
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
            //            await Module.consultas_Relacionales();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Lineas': {ex.Message}");
            //}
        }

        private async void Filtrar_x_linea(string _linea)
        {
            OleDbDataAdapter da = null;
            try
            {
                Application.DoEvents();
                string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.prrazo as Marca," +
                    "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                    //"l.lidesc as Linea," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    //" a.ARMARC as Marca2," + 
                    //("a.arpre2 as Preciopublico, " +//"IIF(a.artt IS NOT NULL AND a.artt <> 0, a.arpre5, a.arpre2) as Preciopublico, "+
                    "IIF(a.artt IS NOT NULL AND a.artt <> 0, a.arpre5, a.arpre2) as Preciopublico, " +
                    "IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from (listas as a ";

                da = new OleDbDataAdapter(SelectQuery +
                                         "inner join proveedo as p on p.prcodi=a.arprov) " +
                                         "LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI " +
                                         "where a.armode='" + _linea + "' " +
                                         "order by a.arcodi", Module.cnn);

                DataSet ds = new DataSet();
                Application.DoEvents();
                da.Fill(ds);
                Application.DoEvents();
                DataGridView1.DataSource = ds.Tables[0];
                DataGridView1.Refresh();

                await Module.SeleccionarGrilla();

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
                    lblTotalRegistros.Text = "ARTÍCULO: " + DataGridView1.RowCount.ToString();
                    if (!TimerVerOfertas.Enabled)
                        TimerVerOfertas.Enabled = true;

                    SimulateClickOnDataGridViewCell(DataGridView1, 0, 0);
                    SimulateClickOnDataGridViewCell(DataGridView1, 1, 0);
                }
            }
        }

        //private void Filtrar_x_linea(int numL)
        //{
        //    OleDbDataAdapter da = null;
        //    try
        //    {
        //        Application.DoEvents();
        //        string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
        //            "a.armarc as Marca," +
        //            //"l.lidesc as Linea," +
        //            "a.armode as Linea," +
        //            " a.arinte as Posición," +
        //            //" a.ARMARC as Marca2," + 
        //            "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
        //            ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
        //            " from listas as a ";

        //        da = new OleDbDataAdapter(SelectQuery +
        //                                 "inner join proveedo as p on p.prcodi=a.arprov " +
        //                                 "where a.arcuar='" + numL.ToString() + "' " +
        //                                 "order by a.arcodi", Module.cnn);

        //        DataSet ds = new DataSet();
        //        Application.DoEvents();
        //        da.Fill(ds);
        //        Application.DoEvents();
        //        DataGridView1.DataSource = ds.Tables[0];
        //        DataGridView1.Refresh();

        //        await Module.SeleccionarGrilla();

        //        if (Module.cnn.State == ConnectionState.Open)
        //            Module.cnn.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        if (Module.cnn.State == ConnectionState.Open)
        //            Module.cnn.Close();

        //        throw new Exception(", error dentro de la Función 'Filtrar_x_linea': " + ex.Message);
        //    }
        //    finally
        //    {
        //        if (da != null)
        //            da.Dispose();

        //        if (this.DataGridView1.RowCount > 0)
        //        {
        //            lblTotalRegistros.Text = "ARTÍCULO: " + DataGridView1.RowCount.ToString();
        //            if (!TimerVerOfertas.Enabled)
        //                TimerVerOfertas.Enabled = true;

        //            SimulateClickOnDataGridViewCell(DataGridView1, 0, 0);
        //            SimulateClickOnDataGridViewCell(DataGridView1, 1, 0);
        //        }
        //    }
        //}


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
        public void ColorearGrillaxOfertas(DataGridView dgv)
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
                            if (dgv["stock", w].Value != null && dgv["stock", w].Value != DBNull.Value)
                            {
                                int stockValue = Convert.ToInt32(dgv["stock", w].Value);
                                if (stockValue >= Module.ini_cond_verde)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Green;
                                    dgv["stock", w].Style.ForeColor = Color.Green;
                                    dgv["stock", w].Style.SelectionBackColor = Color.Green;
                                    dgv["stock", w].Style.SelectionForeColor = Color.Green;
                                }
                                else if (stockValue < Module.ini_cond_verde && stockValue > Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Yellow;
                                    dgv["stock", w].Style.ForeColor = Color.Yellow;
                                    dgv["stock", w].Style.SelectionBackColor = Color.Yellow;
                                    dgv["stock", w].Style.SelectionForeColor = Color.Yellow;
                                }
                                else if (stockValue <= Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Red;
                                    dgv["stock", w].Style.ForeColor = Color.Red;
                                    dgv["stock", w].Style.SelectionBackColor = Color.Red;
                                    dgv["stock", w].Style.SelectionForeColor = Color.Red;
                                }
                            }
                            else
                            {
                                dgv["stock", w].Style.BackColor = Color.Red;
                                dgv["stock", w].Style.ForeColor = Color.Red;
                                dgv["stock", w].Style.SelectionBackColor = Color.Red;
                                dgv["stock", w].Style.SelectionForeColor = Color.Red;
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

                            // NUEVO: Preservar estilo de links en columna Codigo para DataGridView2
                            if (dgv.Name == "DataGridView2" && dgv.Columns.Contains("Codigo"))
                            {
                                dgv["Codigo", w].Style.ForeColor = Color.Blue;
                                var font = dgv.Font;
                                dgv["Codigo", w].Style.Font = new Font(font, FontStyle.Underline);
                                dgv["Codigo", w].ToolTipText = "VER";
                            }
                            // Semáforo de Stock
                            if (dgv["stock", w].Value != null && dgv["stock", w].Value != DBNull.Value)
                            {
                                int stockValue = Convert.ToInt32(dgv["stock", w].Value);
                                if (stockValue >= Module.ini_cond_verde)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Green;
                                    dgv["stock", w].Style.ForeColor = Color.Green;
                                    dgv["stock", w].Style.SelectionBackColor = Color.Green;
                                    dgv["stock", w].Style.SelectionForeColor = Color.Green;
                                }
                                else if (stockValue < Module.ini_cond_verde && stockValue > Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Yellow;
                                    dgv["stock", w].Style.ForeColor = Color.Yellow;
                                    dgv["stock", w].Style.SelectionBackColor = Color.Yellow;
                                    dgv["stock", w].Style.SelectionForeColor = Color.Yellow;
                                }
                                else if (stockValue <= Module.ini_cond_rojo)
                                {
                                    dgv["stock", w].Style.BackColor = Color.Red;
                                    dgv["stock", w].Style.ForeColor = Color.Red;
                                    dgv["stock", w].Style.SelectionBackColor = Color.Red;
                                    dgv["stock", w].Style.SelectionForeColor = Color.Red;
                                }
                            }
                            else
                            {
                                dgv["stock", w].Style.BackColor = Color.Red;
                                dgv["stock", w].Style.ForeColor = Color.Red;
                                dgv["stock", w].Style.SelectionBackColor = Color.Red;
                                dgv["stock", w].Style.SelectionForeColor = Color.Red;
                            }
                        }
                    }

                    if (dgv.Rows.Count == 1)
                    {
                        // Obtener DataTable desde DataSource (puede ser DataTable o DataView)
                        DataTable tabla = null;
                        if (dgv.DataSource is DataTable dt)
                            tabla = dt;
                        else if (dgv.DataSource is DataView dv)
                            tabla = dv.Table;

                        if (tabla != null)
                        {
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
            //try
            //{
            //    Application.DoEvents();
            //    if (!string.IsNullOrEmpty(txtModelo.Text))
            //    {
            //        Module.campo = "a.AROBSE";
            //        Module.valor = txtModelo.Text;
            //        if (!ckbOfertas.Checked)
            //        {
            //            await Module.consultas_Relacionales();
            //        }
            //        else
            //        {
            //            await Module.consultas_Ofertas();
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Filtro Modelo': {ex.Message}");
            //}
            try
            {
                timer2.Stop();
                timer2.Start();
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

                // NUEVO: Re-aplicar estilo de links después del refresh
                EnsureCodigoColumnLinkStyling();

                //SimulateClickOnDataGridViewCell(DataGridView2, 1, 0);
                //SimulateClickOnDataGridViewCell(DataGridView2, 1, 0);
            }
        }

        private async void DataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (!Module.ClienteAutorizado)
                {
                    ApplyEstadoClienteToUI();
                    return;
                }

                // El evento Click ya se disparó con el primer clic del doble clic
                // y ya ejecutó AccionGrilla + Cargar_Imagen.
                // Acá solo limpiamos si la grilla está vacía; si tiene datos,
                // el InputBox aparece inmediatamente sin volver a cargar imagen.
                if (this.DataGridView1.RowCount == 0)
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
                    ptbImagen.Image = System.Drawing.Image.FromFile(Module.ObtenerRutaImagenPorDefectoPublica());
                }

                if (!string.IsNullOrEmpty(this.txtPreciopublico.Text))
                {
                    int NumFila = this.DataGridView1.CurrentCell.RowIndex;
                    string Cantidad = Microsoft.VisualBasic.Interaction.InputBox("Ingrese la cantidad Solicitada",
                        "Agregar Pieza " + this.DataGridView1["CODIGO", NumFila].Value.ToString() +
                        " Marca " + this.DataGridView1["Marca", NumFila].Value.ToString(), "1");

                    string StringSelect = "Select ARFLIA, ARPROV " +
                                        " from listas as a " +
                                        "INNER JOIN proveedo AS p ON p.prcodi = a.arprov " +
                                        " WHERE ARTERC='" + this.DataGridView1["CODIGO", NumFila].Value.ToString() +
                                        "' AND p.prmarc='" + this.DataGridView1["Marca", NumFila].Value.ToString() + "'";

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
                    int numFila = DataGridView2.CurrentCell.RowIndex;
                    int numColumna = DataGridView2.CurrentCell.ColumnIndex;

                    // NUEVO: Si se hizo clic en la columna CODIGO, ejecutar búsqueda
                    if (DataGridView2.Columns[numColumna].Name == "Codigo")
                    {
                        string codigoBuscar = DataGridView2["Codigo", numFila].Value.ToString();

                        // Limpiar campos de búsqueda y poner el código
                        txtDescripcion.Text = "";
                        txtModelo.Text = "";
                        txtEquivalencia.Text = "";
                        txtCodigo.Text = codigoBuscar;

                        // Resetear combos
                        cmbRubros.SelectedIndex = -1;
                        cmbMarcas.SelectedIndex = -1;
                        cmbArmadora.SelectedIndex = -1;
                        cmbLineas.SelectedIndex = -1;

                        // CORREGIDO: Ejecutar filtro por código (como si presionara Enter)
                        FiltrarPorCodigo();

                        return; // No ejecutar el resto del código
                    }

                    // MANTENER: Funcionalidad original para otras columnas
                    string StringSelect = "SELECT TOP 1 IIF(R.RBDESC IS NULL, '', R.RBDESC) AS RBDESC, IIF(A.ARORIG IS NULL, '', A.ARORIG) AS ARORIG" +
                        ", IIF(A.AROBSE IS NULL, '', A.AROBSE) AS AROBSE " +
                        ", IIF(A.ARUBIC1 IS NULL, '', A.ARUBIC1) AS ARUBIC1 " +
                    " FROM (listas AS A INNER JOIN proveedo AS p ON p.prcodi = A.arprov) " +
                    " INNER JOIN RUBROS AS R ON A.ARRUB2 = R.RBCODI " +
                    " WHERE A.ARTERC='" + this.DataGridView2["CODIGO", numFila].Value.ToString() +
                    "' AND p.prmarc='" + this.DataGridView2["Marca", numFila].Value.ToString() + "'";
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
                if (!Module.ClienteAutorizado)
                {
                    ApplyEstadoClienteToUI();
                    return;
                }

                if (this.DataGridView2.RowCount > 0 && this.DataGridView2.CurrentRow != null)
                {
                    int NumFila = this.DataGridView2.CurrentCell.RowIndex;
                    string Cantidad = Microsoft.VisualBasic.Interaction.InputBox("Ingrese la cantidad Solicitada",
                        "Agregar Pieza " + this.DataGridView2["CODIGO", NumFila].Value.ToString() +
                        " Marca " + this.DataGridView2["Marca", NumFila].Value.ToString(), "1");

                    //string StringSelect = "Select ARFLIA, ARPROV from listas " +
                    //                    " WHERE ARTERC='" + this.DataGridView2["CODIGO", NumFila].Value.ToString() +
                    //                    "' AND ARMARC='" + this.DataGridView2["Marca", NumFila].Value.ToString() + "'";


                    string StringSelect = "Select ARFLIA, ARPROV " +
                                        " from listas as a " +
                                        "INNER JOIN proveedo AS p ON p.prcodi = a.arprov " +
                                        " WHERE ARTERC='" + this.DataGridView2["CODIGO", NumFila].Value.ToString() +
                                        "' AND p.prmarc='" + this.DataGridView2["Marca", NumFila].Value.ToString() + "'";

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

        private void confPorFamiliaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Module.DescuentosConfigurados())
                {
                    MostrarNotificacionDescuento();
                    MessageBox.Show(
                        "No es posible acceder a la configuración por familia.\n" +
                        "El cliente no tiene descuentos configurados (DESDESC1 = 0).\n\n" +
                        "Pida sus descuentos a la distribuidora.",
                        "Sin Descuentos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var frmConfFamilia = new frmConfFamilia();
                frmConfFamilia.ShowDialog();
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Abrir Conf. por Familia': {ex.Message}");
                MessageBox.Show("Error al abrir configuración por familia: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void timer2_Tick(object sender, EventArgs e)
        {
            try
            {
                if (_porLimpieza == false)
                {
                    if (!string.IsNullOrWhiteSpace(txtCodigo.Text) || !string.IsNullOrWhiteSpace(txtDescripcion.Text) || !string.IsNullOrWhiteSpace(txtModelo.Text))
                    {
                        // Detiene el temporizador
                        timer2.Stop();
                        // Aquí puedes poner el código que deseas ejecutar cuando el usuario deja de escribir
                        // MessageBox.Show("Terminaste de escribir en el TextBox");

                        if (ckbOfertas.Checked == false) // Filtros comunes
                        {
                            // CargarComboCodigos(".", cmbFamilias.Text);
                            await Module.consultas_Relacionales();
                        }
                        else
                        {
                            await Module.consultas_Ofertas();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Manejo de excepciones (puedes dejar vacío como en VB o agregar logging)
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'timer2_consultas': {ex.Message}");
            }
            finally
            {
                // Detiene el temporizador
                timer2.Stop();
            }
        }

        private void actualizarAplicaciónToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ActualizarAplicacionToolStripMenuItem_Click(sender, e);
        }

        private void ActualizarAplicacionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool ret = false;
            try
            {
                //TimerVerificaUpdate.Stop();
                //ret = Module.DescargaArchivo("UpdateDistribuidoraPalomarNet.exe",
                //                      @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\UpdateDistribuidoraPalomarNet.exe",
                //                      "Descargas",
                //                      "Noelia_1056",
                //                      Module.ini_ruta_ftp.Replace("CatalagoPinky", "UpdateDesktop"));//ini_ruta_ftp ruta donde se aloja el update.

                ret = Module.DescargaArchivo(Module.ini_apli_ftp,
                                    @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\" + Module.ini_apli_ftp,
                                    Module.ini_usu_ftp,
                                    Module.ini_pass_ftp,
                                    Module.ini_ruta_ftp);


                if (ret && File.Exists(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\" + Module.ini_apli_ftp))
                {
                    Process.Start(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\TerminarUpdateDistribuidoraPalomarNet.bat"); //Funciona, pero no es lo ideal
                }
            }
            catch (Exception ex)
            {
                // Puedes agregar manejo de errores aquí si lo deseas
                // Console.WriteLine(ex.Message);
            }
            finally
            {
                //TimerVerificaUpdate.Start();
            }
        }

        private void lblusuario_Click(object sender, EventArgs e)
        {

        }

        private void CheckBox1_Click_1(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

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

        private void DataGridView2_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // Solo cambiar el cursor si estamos sobre la columna "Codigo"
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0 &&
                    DataGridView2.Columns[e.ColumnIndex].Name == "Codigo")
                {
                    DataGridView2.Cursor = Cursors.Hand;
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Mouse Enter Grilla': {ex.Message}");
            }
        }

        private void DataGridView2_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // Restaurar el cursor normal al salir de cualquier celda
                DataGridView2.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Mouse Leave Grilla': {ex.Message}");
            }
        }

        private void EnsureCodigoColumnLinkStyling()
        {
            try
            {
                // Reapply link styling to Codigo column
                if (DataGridView2.Columns.Contains("Codigo"))
                {
                    var estiloLink = new DataGridViewCellStyle
                    {
                        ForeColor = Color.Blue,
                        Font = new Font(DataGridView2.Font, FontStyle.Underline),
                        SelectionForeColor = Color.Blue,
                        SelectionBackColor = Color.LightBlue
                    };
                    DataGridView2.Columns["Codigo"].DefaultCellStyle = estiloLink;

                    // CORREGIDO: No poner tooltip en el header, se maneja individualmente por celda
                    DataGridView2.Columns["Codigo"].ToolTipText = "";

                    // Aplicar tooltip "VER" a cada celda individual de la columna Codigo
                    foreach (DataGridViewRow row in DataGridView2.Rows)
                    {
                        if (row.Cells["Codigo"] != null)
                        {
                            row.Cells["Codigo"].ToolTipText = "VER";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Module.GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Ensure Link Styling': {ex.Message}");
            }
        }

    }
}

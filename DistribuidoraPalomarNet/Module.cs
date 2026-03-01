using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Policy;
using System.Threading.Tasks;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Image = System.Drawing.Image;
using ITextFont = iTextSharp.text.Font;
using ITextImage = iTextSharp.text.Image;

namespace DistribuidoraPalomarNet
{
    public static class Module
    {
        #region Variables Globales
        // Referencias a formularios
        private static frmPrincipal _frmPrincipal;
        private static frmCarroCompras _frmCarroCompras;
        private static frmDatosUsuario _frmDatosUsuario;
        private static frmCarroCompras _carroComprasForm;
        private static frmDatosUsuario _datosUsuarioForm;

        public static void SetMainForm(frmPrincipal form)
        {
            _frmPrincipal = form ?? throw new ArgumentNullException(nameof(form));
        }

        public static void SetCarroComprasForm(frmCarroCompras form)
        {
            _frmCarroCompras = form ?? throw new ArgumentNullException(nameof(form));
        }

        public static void SetDatosUsuarioForm(frmDatosUsuario form)
        {
            if (form == null) throw new ArgumentNullException(nameof(form));
            _frmDatosUsuario = form;
            ModuloUsuario.SetDatosUsuarioForm(form); // Actualizar también la referencia en ModuloUsuario
        }

        public static frmCarroCompras GetCarroComprasForm()
        {
            if (_carroComprasForm == null || _carroComprasForm.IsDisposed)
            {
                _carroComprasForm = new frmCarroCompras();
            }
            return _carroComprasForm;
        }

        public static void Funcion_Novedades()
        {
            using var da = new OleDbDataAdapter();
            try
            {
                // Realizo un select de la tabla articulo y filtro por el campo arfecasig
                // Realizo la conexion con la base de datos
                ConectarDBArti();
                
                da.SelectCommand = new OleDbCommand(
                    "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    //"l.lidesc as Linea," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    //" a.ARMARC as Marca2," + 
                    "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from listas as a " +
                    "INNER JOIN proveedo AS p ON p.prcodi = a.arprov " +
                    "WHERE a.ARFECASIG <> 0 and a.arprim<>'000' order by a.arterc", cnn);

                var ds = new DataSet();
                da.Fill(ds);

                // Una vez que obtengo los datos, los cargo en la grilla para hacerlos visible al usuario
                if (ds.Tables[0].Rows.Count > 0)
                {
                    _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                    _frmPrincipal.DataGridView1.Refresh();
                    SeleccionarGrilla();
                }

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();

                Definir_Grilla();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                    
                throw new Exception($"Error en la Funcion 'Funcion_Novedades' {ex.Message}");
            }
        }

        public static frmDatosUsuario GetDatosUsuarioForm()
        {
            if (_datosUsuarioForm == null || _datosUsuarioForm.IsDisposed)
            {
                _datosUsuarioForm = new frmDatosUsuario();
            }
            return _datosUsuarioForm;
        }

        // Variables de configuración
        public static string ini_usuario { get; set; }
        public static string ini_fecha { get; set; }

        // Conexiones a base de datos
        public static OleDbConnection cnn = new OleDbConnection();
        public static OleDbConnection cnnU = new OleDbConnection();
        public static OleDbConnection cnnO = new OleDbConnection();
        public static OleDbCommand cm;
        public static OleDbDataReader dr;
        public static DataSet ds = new DataSet();
        public static DataTable dt;
        public static string campo = "";
        public static string valor = "";
        
        // Variables globales de configuración
        public static string ini_ruta_arti { get; set; }
        public static string ini_ruta_usu { get; set; }

        // Variables globales de precios
        public static decimal precio_final_desc3;
        public static decimal precio_beneficio;
        public static decimal precio_final_beneficio;
        public static decimal precio_con_desc3;
        public static decimal precio_sin_desc3;
        public static decimal PRECIO_DESC3;
        public static decimal preciofinalxx;
        public static decimal precio_lista;
        public static decimal precio_desc1_desc2;
        public static decimal precio_desc1;
        public static decimal precio_desc2;
        public static decimal precio_final;
        public static decimal precio_costo_sin_iva;
        public static int AccionDatosUsu;
        public static string NumEqui = string.Empty;
        public static string[] vectornomfoto;
        public static string prcodi = "";

        // Variables de configuración FTP y Email
        public static string ini_emisor_email;
        public static string ini_pass_email;
        public static string ini_host_email;
        public static string ini_port_email;
        public static bool ini_enable_ssl;
        public static string ini_receptor_email;
        public static string ini_ruta_ftp = "";
        public static string ini_usu_ftp;
        public static string ini_pass_ftp;
        public static string ini_arch_ftp;
        public static string ini_ruta_base;
        public static string ini_ruta_base_descarga;
        public static string ini_ruta_banner;
        public static string ini_ruta_gral;
        public static string ini_apli_ftp;

        public static string _textoBanner;


        //'para el semaforo de Stock.
        public static int ini_cond_rojo;
        public static int ini_cond_amarillo;
        public static int ini_cond_verde;

        // Variable fecha base
        public static string fecha = "";

        public static bool Migra = false;

        // Configuración Usuario
        public static string US_CODI = "0";
        public static string US_CLRAZO = "0";
        public static string US_CUIT = "0";
        public static string US_EMAIL = "0";

        public static int US_CLCIVA = 0;
        public static int US_CLVEND = 0;
        public static int US_CLCPAG = 0;

        // Variables globales de artículo
        private static int _ARTI_ARFLIA;

        public static int accionDatosUsu;

        public static bool ExisteCodi_Cliente = false;
        public static bool ExisteCuit_Cliente = false;

        public static bool ClienteAutorizado = false;
        public static bool ClienteRestringido = false;
        public static decimal ClienteDescuento1 = 0;
        public static decimal ClienteDescuento2 = 0;
        public static string ClienteMensajeValidacion = string.Empty;

        public static bool _descargarFichero;

        public static string[] TraerCamposArticulo(string stringSelect)
        {
            using var cm = new OleDbCommand();
            OleDbDataReader dr;
            string[] datos = { "0" };

            try
            {
                ConectarDBArti();
                cm.Connection = cnn;
                cm.CommandType = CommandType.Text;
                cm.CommandText = stringSelect;
                dr = cm.ExecuteReader();
                _frmPrincipal.lblSelecion.Text = "";

                while (dr.Read())
                {
                    datos = new[] { dr["ARFLIA"].ToString(), dr["ARPROV"].ToString() };
                }

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();

                return datos;
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();

                throw new Exception($"Error en la Funcion 'TraerCamposArticulo' {ex.Message}");
            }
        }

        public static void SetWherePreparac(string value)
        {
            ModuloPreparac.WHERE_PREPARAC = value;
        }

        public static string GetWherePreparac()
        {
            return ModuloPreparac.WHERE_PREPARAC;
        }

        public static void SetCantPreparac(string value)
        {
            if (decimal.TryParse(value.Replace(".", ","), out decimal result))
            {
                ModuloPreparac.CANT_PREPARAC = result;
            }
            else
            {
                throw new ArgumentException("Invalid decimal value");
            }
        }

        public static decimal GetCantPreparac()
        {
            return ModuloPreparac.CANT_PREPARAC;
        }

        #region Variables Preparac
        public static string ini_codcli = "";
        public static string ini_barr = "";
        public static string ini_razon = "";
        #endregion

        #endregion

        #region Métodos TXT
        public static void Iniciar_Proceso_TXT()
        {
            try
            {
                //frmPrincipal.btnTXT.Enabled = false;
                _frmPrincipal.toolStripProgressBar1.Visible = true;
                //_frmPrincipal.toolStripProgressBar1..Refresh();
                //PasarArti_a_TXT();
                PasarArti_a_xls();
                _frmPrincipal.toolStripProgressBar1.Visible = false;
                //frmPrincipal.btnTXT.Enabled = true;
            }
            catch (Exception ex)
            {
                _frmPrincipal.toolStripProgressBar1.Visible = false;
                //frmPrincipal.btnTXT.Enabled = true;
                throw new Exception(", dentro de 'Iniciar_Proceso_TXT': " + ex.Message.ToString());
            }
        }

        public static void PasarArti_a_TXT()
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                da = new OleDbDataAdapter(
                    "SELECT iif(isnull(a.arterc),0,a.arterc) as arterc, iif(isnull(a.arcuar),0,a.arcuar) as arcuar, " +
                    "mid(a.armarc,1,15) as armarc, mid(iif(isnull(a.ardesc),0,a.ardesc),1,80) as ardesc, " +
                    "mid(iif(isnull(a.arobse),0,a.arobse),1,80) as arobse, iif(isnull(a.arpre2),0,a.arpre2) as aroo " +
                    "from listas as a where a.arprim<>'000' order by a.arcodi", cnn);

                DataSet ds = new DataSet();
                da.Fill(ds, "articulo");

                using (StreamWriter writer = new StreamWriter(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\lista_txt.txt", false))
                {
                    foreach (DataRow fila in ds.Tables["articulo"].Rows)
                    {
                        string codi = fila["arterc"].ToString().PadRight(10);
                        string cuar = fila["arcuar"].ToString().PadRight(4);
                        string marc = fila["armarc"].ToString().PadRight(15);
                        string desc = fila["ardesc"].ToString().PadRight(80);
                        string obse = fila["arobse"].ToString().PadRight(80);
                        string precio = fila["aroo"].ToString().PadRight(10);

                        string linea = codi + " " + cuar + " " + precio + " " + marc + " " + desc + " " + obse;
                        writer.WriteLine(linea.Replace('\n', ' ').Replace('\r', ' '));
                    }
                }

                ds.Dispose();

                if (File.Exists(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\lista_txt.txt"))
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\lista_txt.txt",
                        UseShellExecute = true
                    };
                    Process.Start(startInfo);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error dentro de 'PasarArti_a_TXT': " + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                da.Dispose();
            }
        }

        #endregion
        #region Métodos XLSX
        public static void PasarArti_a_xls()
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                da = new OleDbDataAdapter(
                    "SELECT iif(isnull(a.arterc),0,a.arterc) as arterc, iif(isnull(a.arcuar),0,a.arcuar) as arcuar, " +
                    "mid(a.armarc,1,15) as armarc, mid(iif(isnull(a.ardesc),0,a.ardesc),1,80) as ardesc, " +
                    "mid(iif(isnull(a.arobse),0,a.arobse),1,80) as arobse, iif(isnull(a.arpre2),0,a.arpre2) as aroo " +
                    "from listas as a where a.arprim<>'000' order by a.arcodi", cnn);

                DataSet ds = new DataSet();
                da.Fill(ds, "articulo");

                string excelPath = @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\lista_excel.xlsx";

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Lista de Artículos");

                    // Agregar encabezados
                    worksheet.Cell(1, 1).Value = "Código";
                    worksheet.Cell(1, 2).Value = "Cuarto";
                    worksheet.Cell(1, 3).Value = "Marca";
                    worksheet.Cell(1, 4).Value = "Descripción";
                    worksheet.Cell(1, 5).Value = "Observaciones";
                    worksheet.Cell(1, 6).Value = "Precio de Lista";

                    // Dar formato al encabezado
                    var headerRow = worksheet.Row(1);
                    headerRow.Style.Font.Bold = true;
                    headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                    // Agregar datos
                    int row = 2;
                    foreach (DataRow fila in ds.Tables["articulo"].Rows)
                    {
                        worksheet.Cell(row, 1).Value = fila["arterc"].ToString();
                        worksheet.Cell(row, 2).Value = fila["arcuar"].ToString();
                        worksheet.Cell(row, 3).Value = fila["armarc"].ToString();
                        worksheet.Cell(row, 4).Value = fila["ardesc"].ToString();
                        worksheet.Cell(row, 5).Value = fila["arobse"].ToString();
                        worksheet.Cell(row, 6).Value = Convert.ToDouble(fila["aroo"]);
                        row++;
                    }

                    // Ajustar el ancho de las columnas automáticamente
                    worksheet.Columns().AdjustToContents();

                    // Dar formato a la columna de precio como moneda
                    var rangoPrecios = worksheet.Column(6).CellsUsed();
                    rangoPrecios.Style.NumberFormat.Format = "$#,##0.00";

                    // Guardar el archivo
                    workbook.SaveAs(excelPath);
                }

                ds.Dispose();

                if (File.Exists(excelPath))
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = excelPath,
                        UseShellExecute = true
                    };
                    Process.Start(startInfo);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error dentro de 'PasarArti_a_xls': " + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                da.Dispose();
            }
        }
        #endregion
        #region Métodos Principales
        public static void ActualizarListado()
        {
            try
            {
                ConectarDBArti();
                string sql = "SELECT * from listass ORDER BY codigo";
                OleDbDataAdapter da = new OleDbDataAdapter(sql, cnn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                _frmPrincipal.DataGridView1.DataSource = dt;
                _frmPrincipal.lblTotalRegistros.Text = $"Total Registros: {dt.Rows.Count}";
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al actualizar listado: {ex.Message}");
            }
            finally
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
        }

        public static void CargarListado()
        {
            try
            {
                ConectarDBArti();
                string sql = "SELECT * from listass";

                if (!string.IsNullOrEmpty(_frmPrincipal.txtCodigo.Text))
                    sql += $" WHERE codigo LIKE '%{_frmPrincipal.txtCodigo.Text}%'";
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text))
                    sql += $" WHERE descripcion LIKE '%{_frmPrincipal.txtDescripcion.Text}%'";
                else if (_frmPrincipal.cmbRubros.SelectedIndex > -1)
                    sql += $" WHERE rubro = '{_frmPrincipal.cmbRubros.Text}'";

                sql += " ORDER BY codigo";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, cnn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                _frmPrincipal.DataGridView1.DataSource = dt;
                _frmPrincipal.lblTotalRegistros.Text = $"Total Registros: {dt.Rows.Count}";
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al cargar listado: {ex.Message}");
            }
            finally
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
        }

        public static void GenerarTxt()
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Archivos de texto|*.txt",
                    Title = "Guardar archivo TXT"
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    using (StreamWriter writer = new StreamWriter(saveDialog.FileName))
                    {
                        foreach (DataGridViewRow row in _frmPrincipal.DataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                string linea = $"{row.Cells["codigo"].Value}|{row.Cells["descripcion"].Value}|{row.Cells["precio"].Value}";
                                writer.WriteLine(linea);
                            }
                        }

                        if (cnnU.State == ConnectionState.Open)
                            cnnU.Close();
                    }
                    MessageBox.Show("Archivo TXT generado correctamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al generar TXT: {ex.Message}");
            }
        }
        #endregion

        #region Conexiones
        public static void Conectar_Origen_Actualizar(string dbOrigen)
        {
            try
            {
                if (cnnO.State == ConnectionState.Open)
                {
                    cnnO.Close();
                }
                cnnO.ConnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={dbOrigen};";
                cnnO.Open();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error dentro del método 'Conectar_Origen_Actualizar': {ex.Message}");
            }
        }

        public static void ConectarDBArti()
        {
            try
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                cnn.ConnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={ini_ruta_base};";
                cnn.Open();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en el sub 'ConectarDBArti': {ex.Message}");
            }
        }

        public static void ConectarDBUsuario()
        {
            try
            {
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
                cnnU.ConnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={ini_ruta_usu};";//"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={ini_ruta_usu};";
                cnnU.Open();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en el sub 'ConectarDBUsuario': {ex.Message}");
            }
        }
        #endregion

        public static void Inicializacion()
        {
            try
            {
                //DetectarAplicacion();
                Application.DoEvents();
                LeerConfig();
                Cargar_Usuario();
                //ActualizarFAMIVA();
                funcion_inicial(0);
                aplicar_desc_Cliente(); // Cargo Descuentos 1 y 2
                _frmPrincipal.lblFechaAct.Text = $"Actualizado al {FechaActualizacion()}";
                Application.DoEvents();
                SeleccionarGrilla();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error dentro de 'Inicializacion': {ex.Message}");
            }
        }

        public static void funcion_inicial(int opc)
        {
            using var da = new OleDbDataAdapter();
            try
            {
                ConectarDBArti();
                
                if (opc == 0)
                {
                    // Solo inicialización visual: combos + estructura de columnas vacía
                    CargarComboRubros();
                    CargarComboLineas("0");

                    // Crear una tabla vacía con las columnas que se usan en la grilla
                    var tablaVacia = new DataTable();
                    tablaVacia.Columns.Add("Codigo", typeof(string));
                    tablaVacia.Columns.Add("Descripcion", typeof(string));
                    tablaVacia.Columns.Add("Marca", typeof(string));
                    tablaVacia.Columns.Add("Linea", typeof(string));
                    tablaVacia.Columns.Add("Posición", typeof(string));
                    tablaVacia.Columns.Add("Preciopublico", typeof(decimal));
                    tablaVacia.Columns.Add("aroo", typeof(decimal));
                    tablaVacia.Columns.Add("Stock", typeof(int));
                    tablaVacia.Columns.Add("artt", typeof(decimal));
                    tablaVacia.Columns.Add("arubic1", typeof(string));
                    tablaVacia.Columns.Add("arobse", typeof(string));

                    _frmPrincipal.DataGridView1.DataSource = tablaVacia;
                    _frmPrincipal.DataGridView1.Refresh();

                    Application.DoEvents();
                    Definir_Grilla();
                }
                else
                {
                    Application.DoEvents();

                    da.SelectCommand = new OleDbCommand(
                        "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                        "p.PRMARC as Marca," +
                        //"a.armarc as Marca," +
                        //"l.lidesc as Linea," +
                        "a.armode as Linea," + 
                        " a.arinte as Posición," +
                        //" a.ARMARC as Marca2," + 
                        "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                        ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                        " from listas as a " +
                        "INNER JOIN proveedo AS p ON p.prcodi = a.arprov " +
                        //" left join lineas as l on a.arcuar=l.licodi= " +
                        "where a.arprim<>'000' order by a.arterc", cnn);

                    var ds = new DataSet();
                    da.Fill(ds);

                    _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                    _frmPrincipal.DataGridView1.Refresh();

                    SeleccionarGrilla();
                }
                
                if (cnn.State == ConnectionState.Open)
                {
                    cnn.Close();
                }
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                {
                    cnn.Close();
                }
                throw new Exception($"Error en la Funcion 'funcion_inicial' {ex.Message}");
            }
            finally
            {
                if (_frmPrincipal.DataGridView1.RowCount>0)
                {
                    if (_frmPrincipal.TimerVerOfertas.Enabled==false)
                    {
                        _frmPrincipal.TimerVerOfertas.Enabled = true;
                    }
                    _frmPrincipal.SimulateClickOnDataGridViewCell(_frmPrincipal.DataGridView1, 0, 0);
                    _frmPrincipal.SimulateClickOnDataGridViewCell(_frmPrincipal.DataGridView1, 1, 0);
                }
            }

        }

        //public static void Definir_Grilla()
        //{
        //    var miEstilo = new DataGridViewCellStyle(_frmPrincipal.DataGridView1.Columns["Preciopublico"].DefaultCellStyle)
        //    {
        //        Font = new Font("Arial", 14, FontStyle.Bold)
        //    };

        //    _frmPrincipal.DataGridView1.Columns["Preciopublico"].HeaderCell.Style = miEstilo;
        //    _frmPrincipal.DataGridView1.Columns["Preciopublico"].HeaderCell.Style.BackColor = Color.Black;
        //    _frmPrincipal.DataGridView1.Columns["Preciopublico"].HeaderCell.Style.ForeColor = Color.White;
        //    _frmPrincipal.DataGridView1.Columns["Preciopublico"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        //    // Set column widths (converting VB.NET percentage to approximate pixel values)
        //    _frmPrincipal.DataGridView1.Columns["Codigo"].Width = 70;  // codigo - 7%
        //    _frmPrincipal.DataGridView1.Columns["Marca"].Width = 150;  // Marca - 7%
        //    _frmPrincipal.DataGridView1.Columns["Descripcion"].Width = 300; // descripcion - 30%
        //    //_frmPrincipal.DataGridView1.Columns["Observaciones"].Visible = false; // Observaciones
        //    _frmPrincipal.DataGridView1.Columns["Linea"].Width = 100; // 30%
        //    _frmPrincipal.DataGridView1.Columns["Modelo"].Width = 100; // 30%
        //    _frmPrincipal.DataGridView1.Columns["Posición"].Width = 100; // 30%
        //    //_frmPrincipal.DataGridView1.Columns["Observaciones"].Width = 300; // 30%
        //    //_frmPrincipal.DataGridView1.Columns["Observaciones"].HeaderText = "Observaciones";
        //    _frmPrincipal.DataGridView1.Columns["Preciopublico"].Width = 150; // aroo=precio - 15%
        //    _frmPrincipal.DataGridView1.Columns["Preciopublico"].HeaderText = "Precio Público";
        //    _frmPrincipal.DataGridView1.Columns["aroo"].Width = 10;  // 1%
        //    _frmPrincipal.DataGridView1.Columns["aroo"].Visible = false; // Cod familia
        //    _frmPrincipal.DataGridView1.Columns["artt"].Visible = false;
        //    _frmPrincipal.DataGridView1.Columns["arubic1"].Visible = false;
        //    _frmPrincipal.DataGridView1.Columns["arobse"].Visible = false;
        //}
        public static void Definir_Grilla()
        {
            var miEstilo = new DataGridViewCellStyle(_frmPrincipal.DataGridView1.Columns["Preciopublico"].DefaultCellStyle)
            {
                Font = new Font("Arial", 12, FontStyle.Bold)
            };

            var miEstiloStock = new DataGridViewCellStyle(_frmPrincipal.DataGridView1.Columns["Stock"].DefaultCellStyle)
            {
                Font = new Font("Arial", 2, FontStyle.Regular)
            };

            _frmPrincipal.DataGridView1.Columns["Stock"].DefaultCellStyle = miEstiloStock;

            _frmPrincipal.DataGridView1.Columns["Preciopublico"].HeaderCell.Style = miEstilo;
            _frmPrincipal.DataGridView1.Columns["Preciopublico"].HeaderCell.Style.BackColor = Color.Black;
            _frmPrincipal.DataGridView1.Columns["Preciopublico"].HeaderCell.Style.ForeColor = Color.White;
            _frmPrincipal.DataGridView1.Columns["Preciopublico"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            // Establecer el modo de ajuste de tamaño de las columnas
            _frmPrincipal.DataGridView1.Columns["Codigo"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _frmPrincipal.DataGridView1.Columns["Marca"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _frmPrincipal.DataGridView1.Columns["Descripcion"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _frmPrincipal.DataGridView1.Columns["Linea"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //_frmPrincipal.DataGridView1.Columns["Modelo"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _frmPrincipal.DataGridView1.Columns["Posición"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _frmPrincipal.DataGridView1.Columns["Preciopublico"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _frmPrincipal.DataGridView1.Columns["aroo"].Visible = false; // Cod familia
            _frmPrincipal.DataGridView1.Columns["artt"].Visible = false;
            _frmPrincipal.DataGridView1.Columns["arubic1"].Visible = false;
            _frmPrincipal.DataGridView1.Columns["arobse"].Visible = false;
            _frmPrincipal.DataGridView1.Columns["Posición"].Visible = false;

            // Establecer pesos de llenado para cada columna
            _frmPrincipal.DataGridView1.Columns["Codigo"].FillWeight = 7;
            _frmPrincipal.DataGridView1.Columns["Marca"].FillWeight = 7;
            _frmPrincipal.DataGridView1.Columns["Descripcion"].FillWeight = 30;
            _frmPrincipal.DataGridView1.Columns["Linea"].FillWeight = 30;
            //_frmPrincipal.DataGridView1.Columns["Modelo"].FillWeight = 30;
            _frmPrincipal.DataGridView1.Columns["Posición"].FillWeight = 30;
            _frmPrincipal.DataGridView1.Columns["Preciopublico"].FillWeight = 15;
        }

        public static void SeleccionarGrilla()
        {
            try
            {
                if (_frmPrincipal.DataGridView1.RowCount != 0)
                {
                    _frmPrincipal.lblTotalRegistros.Text = $"ARTÍCULO: {_frmPrincipal.DataGridView1.RowCount}";
                    _frmPrincipal.DataGridView1.CurrentCell = _frmPrincipal.DataGridView1.Rows[0].Cells[0];
                    AccionGrilla(_frmPrincipal.DataGridView1);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error dentro 'SeleccionarGrilla': {ex.Message}");
            }
        }

        public static void AccionGrilla(DataGridView dataGrid)
        {
            try
            {
                Cargar_Imagen();
                // Por ahora comento las equivalencias
                if(dataGrid.Name== "DataGridView1") ValidarEquivalencias();
                
                int numFila = dataGrid.CurrentCell.RowIndex;
                if (dataGrid.Name == "DataGridView1") TraerDatosExtras(numFila);
                ColorearStock();
                
                // frmPrincipal.lblCodigoSelec.Text = "Código: " + frmPrincipal.DataGridView1["Codigo", NumFila].Value.ToString();
                
                // Obtener precio lista
                string precioLista = dataGrid["aroo", numFila].Value.ToString();
                _frmPrincipal.txtPreciolista.Text = precioLista;
                
                // Por ahora comento desc3 y bene
                // aplicar_desc_Benef_prov(DataGridView1[2, NumFila].Value); // Cargo Descuente 3 y Beneficio Cliente y el iva
                
                Calculadora();
                //aplicar_descuento3();
                aplicar_beneficio();

                try
                {
                    _frmPrincipal.txtDescuentoFamilia.Text = string.Format("{0:F2}", aplicar_desc_Familia());
                }
                catch
                {
                    _frmPrincipal.txtDescuentoFamilia.Text = "0";
                }
                //txtIvaFlia
                try
                {
                    _frmPrincipal.txtIvaFlia.Text = string.Format("{0:F2}", aplicar_iva_Familia());
                }
                catch
                {
                    _frmPrincipal.txtIvaFlia.Text = "0";
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error dentro de la funcion 'AccionGrilla': {ex.Message}");
            }
        }

        public static void TraerDatosExtras(int numFila)
        {
            using var cm = new OleDbCommand();
            try
            {
                string pieza = _frmPrincipal.DataGridView1["Codigo", numFila].Value.ToString();
                string desc = string.Empty;
                if (!string.IsNullOrEmpty(_frmPrincipal.DataGridView1["Descripcion", numFila].Value?.ToString()))
                {
                    desc = _frmPrincipal.DataGridView1["Descripcion", numFila].Value.ToString();
                }

                string mode = string.Empty;
                _frmPrincipal.DataGridView1.Columns["arobse"].Visible = true;
                if (!string.IsNullOrEmpty(_frmPrincipal.DataGridView1["arobse", numFila].Value?.ToString()))
                {
                    mode = _frmPrincipal.DataGridView1["arobse", numFila].Value.ToString();
                }
                _frmPrincipal.DataGridView1.Columns["arobse"].Visible = false;

                string sqlString = $"select arubic1, arobse, arff, arflia from listas where arterc='{pieza}' and ardesc='{desc}' and arobse='{mode}'";

                ConectarDBArti();
                cm.Connection = cnn;
                cm.CommandType = CommandType.Text;
                cm.CommandText = sqlString;

                using var dr = cm.ExecuteReader();
                _frmPrincipal.lblSelecion.Text = string.Empty;

                while (dr.Read())
                {
                    _frmPrincipal.lblSelecion.Text = $"{dr["arubic1"]} | {dr["arobse"]}";
                    _frmPrincipal.txtStock.Text = dr["arff"].ToString();
                    _ARTI_ARFLIA = Convert.ToInt32(dr["arflia"]);
                }

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                throw new Exception($"Error dentro de 'TraerDatosExtras': {ex}");
            }
            finally
            {
                _frmPrincipal.DataGridView1.Columns["arobse"].Visible = false;
            }
        }

        public static void Cargar_Imagen()
        {
            try
            {
                int numFila = _frmPrincipal.DataGridView1.CurrentCellAddress.Y;
                string imagePath = Path.Combine("C:", "DistribuidoraPalomarNet", "DistribuidoraPalomarNet", "imagenes", "FOTOS",
                    $"{_frmPrincipal.DataGridView1["Codigo", numFila].Value}.jpg");

                string defaultImagePath = Path.Combine("C:", "DistribuidoraPalomarNet", "DistribuidoraPalomarNet", "imagenes", "SINFOTO.png");

                _frmPrincipal.ptbImagen.Image = File.Exists(imagePath)
                    ? Image.FromFile(imagePath)
                    : Image.FromFile(defaultImagePath);
            }
            catch (Exception)
            {
                string defaultImagePath = Path.Combine("C:", "DistribuidoraPalomarNet", "DistribuidoraPalomarNet", "imagenes", "SINFOTO.png");
                _frmPrincipal.ptbImagen.Image = Image.FromFile(defaultImagePath);
            }
        }

        public static void ColorearStock()
        {
            try
            {
                if (!string.IsNullOrEmpty(_frmPrincipal.txtStock.Text))
                {
                    int cantStock = Convert.ToInt32(_frmPrincipal.txtStock.Text);
                    
                    // Default color - white
                    _frmPrincipal.txtStock.BackColor = Color.White;
                    _frmPrincipal.txtStock.ForeColor = Color.White;

                    // Stock equal to 2
                    if (cantStock == 2)
                    {
                        _frmPrincipal.txtStock.BackColor = Color.Yellow;
                        _frmPrincipal.txtStock.ForeColor = Color.Yellow; // Black in original
                    }

                    // Stock greater than or equal to 3
                    if (cantStock >= 3)
                    {
                        _frmPrincipal.txtStock.BackColor = Color.Green;
                        _frmPrincipal.txtStock.ForeColor = Color.Green; // White in original
                    }

                    // Stock less than or equal to 1
                    if (cantStock <= 1)
                    {
                        _frmPrincipal.txtStock.BackColor = Color.Red;
                        _frmPrincipal.txtStock.ForeColor = Color.Red; // White in original
                    }
                }
            }
            catch (Exception)
            {
                // Original empty catch block maintained
            }
        }

        public static void aplicar_beneficio()
        {
            try
            {
                // Base para sugerido: si hay desc3 aplicado, usar precio_final_desc3; si no, precio_final
                decimal baseSugerido = precio_final_desc3 != 0 && _frmPrincipal.txtDescuento3.Text != "0"
                    ? precio_final_desc3
                    : precio_final;

                // Buscar FADESC por _ARTI_ARFLIA en base de artículos
                string fad = string.Empty;
                using (var cmArti = new OleDbCommand("SELECT TOP 1 FADESC FROM familias WHERE FACODI = ?", cnn))
                {
                    try
                    {
                        ConectarDBArti();
                        cmArti.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
                        var objF = cmArti.ExecuteScalar();
                        if (objF != null && objF != DBNull.Value)
                            fad = objF.ToString();
                    }
                    finally
                    {
                        if (cnn.State == ConnectionState.Open)
                            cnn.Close();
                    }
                }

                // Buscar FABEN1 por FADESC en base de usuario
                decimal? faben1 = null;
                if (!string.IsNullOrEmpty(fad))
                {
                    using (var cmUsu = new OleDbCommand("SELECT TOP 1 FABEN1 FROM FAMILIAS WHERE FADESC = ?", cnnU))
                    {
                        try
                        {
                            ConectarDBUsuario();
                            cmUsu.Parameters.AddWithValue("@p1", fad);
                            var objB = cmUsu.ExecuteScalar();
                            if (objB != null && objB != DBNull.Value && decimal.TryParse(objB.ToString(), out var benTmp))
                                faben1 = benTmp;
                        }
                        finally
                        {
                            if (cnnU.State == ConnectionState.Open)
                                cnnU.Close();
                        }
                    }
                }

                if (faben1.HasValue && faben1.Value != 0)
                {
                    _frmPrincipal.txtBeneficio.Text = string.Format("{0:F2}", faben1.Value);
                    _frmPrincipal.txtSugerido.Text = string.Format("{0:F2}", baseSugerido + (baseSugerido * faben1.Value / 100));
                }
                else
                {
                    _frmPrincipal.txtBeneficio.Text = "0";
                    _frmPrincipal.txtSugerido.Text = string.Format("{0:F2}", baseSugerido);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en el sub 'aplicar_beneficio': {ex.Message}");
            }
        }

        public static void aplicar_descuento3()
        {
            try
            {
                if (!string.IsNullOrEmpty(_frmPrincipal.txtDescuento1.Text) && 
                    _frmPrincipal.txtDescuento1.Text != "0" &&
                    !string.IsNullOrEmpty(_frmPrincipal.txtDescuento3.Text))
                {
                    if (_frmPrincipal.txtDescuento3.Text != "0")
                    {
                        preciofinalxx = precio_final;
                        // precio al publico con desc y iva
                        PRECIO_DESC3 = decimal.Parse(string.Format("{0:F2}", 
                            preciofinalxx * Convert.ToDecimal(_frmPrincipal.txtDescuento3.Text) / 100));
                        precio_final_desc3 = decimal.Parse(string.Format("{0:F2}", (preciofinalxx - PRECIO_DESC3)));
                        precio_con_desc3 = precio_final_desc3;
                        _frmPrincipal.txtPreciopublico.Text = string.Format("{0:F2}", precio_final_desc3);
                    }
                    else
                    {
                        precio_sin_desc3 = precio_final;
                        _frmPrincipal.txtPreciopublico.Text = string.Format("{0:F2}", precio_final);
                    }
                }
                else
                {
                    _frmPrincipal.txtPreciopublico.Text = string.Format("{0:F2}", precio_final);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en el sub 'aplicar_descuento3': {ex.Message}");
            }
        }

        public static void Calculadora()
        {
            try
            {
                if (_frmPrincipal.txtPreciolista.Text != string.Empty)
                {
                    precio_lista = Convert.ToDecimal(_frmPrincipal.txtPreciolista.Text);
                    if (_frmPrincipal.txtDescuento1.Text != string.Empty && _frmPrincipal.txtDescuento2.Text != string.Empty)
                    {
                        decimal Descuento1 = Convert.ToDecimal(_frmPrincipal.txtDescuento1.Text);
                        decimal Descuento2 = Convert.ToDecimal(_frmPrincipal.txtDescuento2.Text);
                        if (decimal.TryParse(Descuento1.ToString(), out _) && Descuento1 != 0)
                        {
                            // descuentos 1 y 2
                            precio_desc1 = precio_lista - (precio_lista * Descuento1 / 100);
                            precio_desc2 = precio_desc1 - (precio_desc1 * Descuento2 / 100);
                            decimal descFamilia = aplicar_desc_Familia();
                            if (descFamilia != 0)
                            {
                                precio_desc2 = precio_desc2 - (precio_desc2 * descFamilia / 100);
                            }
                            // neto sin IVA
                            precio_costo_sin_iva = precio_desc2;

                            // aplicar IVA por familia si existe
                            decimal precioConIva = precio_desc2;
                            string fad = string.Empty;
                            using (var cmArti = new OleDbCommand("SELECT TOP 1 FADESC FROM familias WHERE FACODI = ?", cnn))
                            {
                                try
                                {
                                    ConectarDBArti();
                                    cmArti.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
                                    var objF = cmArti.ExecuteScalar();
                                    if (objF != null && objF != DBNull.Value)
                                        fad = objF.ToString();
                                }
                                finally
                                {
                                    if (cnn.State == ConnectionState.Open)
                                        cnn.Close();
                                }
                            }

                            if (!string.IsNullOrEmpty(fad))
                            {
                                using (var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA FROM FAMILIAS WHERE FADESC = ?", cnnU))
                                {
                                    try
                                    {
                                        ConectarDBUsuario();
                                        cmUsu.Parameters.AddWithValue("@p1", fad);
                                        var objI = cmUsu.ExecuteScalar();
                                        if (objI != null && objI != DBNull.Value && decimal.TryParse(objI.ToString(), out var ivaTmp) && ivaTmp != 0)
                                        {
                                            precioConIva = precio_desc2 + (precio_desc2 * ivaTmp / 100);
                                        }
                                    }
                                    finally
                                    {
                                        if (cnnU.State == ConnectionState.Open)
                                            cnnU.Close();
                                    }
                                }
                            }

                            precio_final = Convert.ToDecimal(precioConIva);
                            _frmPrincipal.txtPreciopublico.Text = string.Format("{0:F2}", precio_final);
                        }
                        if (!decimal.TryParse(Descuento1.ToString(), out _) || Descuento1 == 0)
                        {
                            _frmPrincipal.txtPreciolista.Text = string.Format("{0:F2}", precio_lista);
                            precio_costo_sin_iva = Convert.ToDecimal(_frmPrincipal.txtPreciolista.Text);

                            decimal descFamilia = aplicar_desc_Familia();
                            if (descFamilia != 0)
                            {
                                precio_costo_sin_iva = precio_costo_sin_iva - (precio_costo_sin_iva * descFamilia / 100);
                            }

                            // aplicar IVA por familia también cuando no hay descuentos
                            decimal precioConIva = precio_costo_sin_iva;
                            string fad = string.Empty;
                            using (var cmArti = new OleDbCommand("SELECT TOP 1 FADESC FROM familias WHERE FACODI = ?", cnn))
                            {
                                try
                                {
                                    ConectarDBArti();
                                    cmArti.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
                                    var objF = cmArti.ExecuteScalar();
                                    if (objF != null && objF != DBNull.Value)
                                        fad = objF.ToString();
                                }
                                finally
                                {
                                    if (cnn.State == ConnectionState.Open)
                                        cnn.Close();
                                }
                            }

                            if (!string.IsNullOrEmpty(fad))
                            {
                                using (var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA FROM FAMILIAS WHERE FADESC = ?", cnnU))
                                {
                                    try
                                    {
                                        ConectarDBUsuario();
                                        cmUsu.Parameters.AddWithValue("@p1", fad);
                                        var objI = cmUsu.ExecuteScalar();
                                        if (objI != null && objI != DBNull.Value && decimal.TryParse(objI.ToString(), out var ivaTmp) && ivaTmp != 0)
                                        {
                                            precioConIva = precio_lista + (precio_lista * ivaTmp / 100);
                                        }
                                    }
                                    finally
                                    {
                                        if (cnnU.State == ConnectionState.Open)
                                            cnnU.Close();
                                    }
                                }
                            }

                            precio_final = precioConIva;
                            _frmPrincipal.txtPreciopublico.Text = string.Format("{0:F2}", precio_final);
                        }
                    }
                    else
                    {
                        // notifico que debe cargar los desc y beneficios para que tenga un precio mas aproximado
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en la funcion 'Calculadora': {ex.Message}");
            }
        }

        public static decimal aplicar_desc_Familia()
        {
            try
            {
                if (_ARTI_ARFLIA == 0)
                    return 0;

                using var cmd = new OleDbCommand("SELECT TOP 1 FADES1 FROM FAMILIAS WHERE FACODI = ?", cnnU);
                ConectarDBUsuario();
                cmd.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
                var obj = cmd.ExecuteScalar();
                if (obj != null && obj != DBNull.Value && decimal.TryParse(obj.ToString(), out var d))
                    return d;
                return 0;
            }
            catch
            {
                return 0;
            }
            finally
            {
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
            }
        }
        public static decimal aplicar_iva_Familia()
        {
            try
            {
                if (_ARTI_ARFLIA == 0)
                    return 0;

                using var cmd = new OleDbCommand("SELECT TOP 1 FAIVA FROM FAMILIAS WHERE FACODI = ?", cnnU);
                ConectarDBUsuario();
                cmd.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
                var obj = cmd.ExecuteScalar();
                if (obj != null && obj != DBNull.Value && decimal.TryParse(obj.ToString(), out var d))
                    return d;
                return 0;
            }
            catch
            {
                return 0;
            }
            finally
            {
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
            }
        }
        public static void funcion_inicial_carrito()
        {
            // carga inicial de la grilla de pedidos
            using var da = new OleDbDataAdapter();
            try
            {
                // "|nºorden|cod.art|cuarta|Descripción|cant.|marca|fecha|p. lista|p.x cant."
                ConectarDBUsuario();
                da.SelectCommand = new OleDbCommand(
                    "SELECT pr.nº, pr.Codigo, pr.Marca, pr.Descripcion, " +
                    "pr.Ubicacion as CódigoOrigen, pr.PrecioUnidad as PrecioListaSinIVA, pr.Cantidad, pr.Total AS SubTotalSinIVA " +
                    "from PedidosRealizados as pr", cnnU);

                var ds = new DataSet();
                da.Fill(ds);
                
                _frmCarroCompras.gridPedidos.DataSource = ds.Tables[0];
                _frmCarroCompras.gridPedidos.Refresh();
                _frmCarroCompras.gridPedidos.Columns[0].Visible = false; // nuevo

                if (_frmCarroCompras.gridPedidos.RowCount >= 1) // si existen registros en la grilla de pedidos
                {
                    // sumo los precios Totales, para dar un estimado
                    double suma = 0;
                    for (int s = 0; s < _frmCarroCompras.gridPedidos.Rows.Count; s++)
                    {
                        Application.DoEvents();
                        suma += Convert.ToDouble(_frmCarroCompras.gridPedidos.Rows[s].Cells["SubTotalSinIVA"].Value.ToString());
                    }

                    _frmCarroCompras.lblTotal.Text = suma != 0 ? $"Total sin IVA: {suma}" : string.Empty;
                }
                else
                {
                    _frmCarroCompras.lblTotal.Text = string.Empty;
                }

                // gridPedidos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
            }
            catch (Exception ex)
            {
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
                throw new Exception($"Error en la Función 'funcion_inicial_carrito' {ex.Message}");
            }
        }

        //public static void consultas_Relacionales()
        //{
        //    using var da = new OleDbDataAdapter();
        //    string QueryFinal = string.Empty;
        //    string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
        //            "a.armarc as Marca," +
        //            //"l.lidesc as Linea," +
        //            "a.armode as Linea, a.arinte as Posición, a.ARMARC as Marca," +
        //            "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
        //            ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
        //            " from listas as a ";

        //    string WhereQuery = string.Empty;
        //    try
        //    {
        //        using var ds = new DataSet();
        //        if (_frmPrincipal.cmbRubros.Text == "TODOS" || _frmPrincipal.cmbRubros.Text == " TODOS")
        //            _frmPrincipal.cmbRubros.Text = string.Empty;

        //        if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // por default
        //            WhereQuery = " where a.arprim<>'000' ";
        //        }
        //        else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // solo por descripcion
        //            campo = "a.ARDESC";
        //            valor = _frmPrincipal.txtDescripcion.Text;
        //            string[] arraydesc = _frmPrincipal.txtDescripcion.Text.Split(' ');
        //            if (arraydesc.Length > 1)
        //            {
        //                for (int r = 0; r < arraydesc.Length; r++)
        //                {
        //                    if (r == 0 && !string.IsNullOrEmpty(arraydesc[r]))
        //                    {
        //                        valor = $" where ({campo} like  '%{arraydesc[r]}%'";
        //                    }
        //                    else if (!string.IsNullOrEmpty(arraydesc[r]))
        //                    {
        //                        valor = $"{valor} AND {campo} like  '%{arraydesc[r]}%'";
        //                    }
        //                }
        //                WhereQuery = $"{valor}) and a.arprim<>'000' ";
        //            }
        //            else // filtro simple
        //            {
        //                valor = _frmPrincipal.txtDescripcion.Text;
        //                WhereQuery = $" where {campo} like  '%{valor}%' and a.arprim<>'000' ";
        //            }
        //        }
        //        // INICIO: todo por RUBROS
        //        else if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // solo rubro
        //            WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
        //        }
        //        else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // solo rubro y descripcion
        //            campo = "a.ARDESC";
        //            valor = _frmPrincipal.txtDescripcion.Text;
        //            string[] arraydesc = _frmPrincipal.txtDescripcion.Text.Split(' ');
        //            if (arraydesc.Length > 1)
        //            {
        //                for (int r = 0; r < arraydesc.Length; r++)
        //                {
        //                    if (r == 0 && !string.IsNullOrEmpty(arraydesc[r]))
        //                    {
        //                        valor = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where ({campo} like  '%{arraydesc[r]}%'";
        //                    }
        //                    else if (!string.IsNullOrEmpty(arraydesc[r]))
        //                    {
        //                        valor = $"{valor} AND {campo} like  '%{arraydesc[r]}%'";
        //                    }
        //                }
        //                WhereQuery = $"{valor}) and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
        //            }
        //            else // filtro simple
        //            {
        //                valor = _frmPrincipal.txtDescripcion.Text;
        //                WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where {campo} like  '%{valor}%' " +
        //                             $"and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
        //            }
        //        }
        //        else if (string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // por default
        //            WhereQuery = " where a.arprim<>'000' ";
        //        }
        //        else if (!string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // solo por descripcion
        //            campo = "a.AROBSE";
        //            valor = _frmPrincipal.txtModelo.Text;
        //            string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');
        //            if (arraymode.Length > 1)
        //            {
        //                for (int r = 0; r < arraymode.Length; r++)
        //                {
        //                    if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
        //                    {
        //                        valor = $" where ({campo} like  '%{arraymode[r]}%'";
        //                    }
        //                    else if (!string.IsNullOrEmpty(arraymode[r]))
        //                    {
        //                        valor = $"{valor} AND {campo} like  '%{arraymode[r]}%'";
        //                    }
        //                }
        //                WhereQuery = $"{valor}) and a.arprim<>'000' ";
        //            }
        //            else // filtro simple
        //            {
        //                valor = _frmPrincipal.txtModelo.Text;
        //                WhereQuery = $" where {campo} like  '%{valor}%' and a.arprim<>'000' ";
        //            }
        //        }
        //        // INICIO: todo por RUBROS
        //        else if (string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // solo rubro
        //            WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
        //        }
        //        else if (!string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
        //        {
        //            // solo rubro y descripcion
        //            campo = "a.AROBSE";
        //            valor = _frmPrincipal.txtModelo.Text;
        //            string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');
        //            if (arraymode.Length > 1)
        //            {
        //                for (int r = 0; r < arraymode.Length; r++)
        //                {
        //                    if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
        //                    {
        //                        valor = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where ({campo} like  '%{arraymode[r]}%'";
        //                    }
        //                    else if (!string.IsNullOrEmpty(arraymode[r]))
        //                    {
        //                        valor = $"{valor} AND {campo} like  '%{arraymode[r]}%'";
        //                    }
        //                }
        //                WhereQuery = $"{valor}) and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
        //            }
        //            else // filtro simple
        //            {
        //                valor = _frmPrincipal.txtModelo.Text;
        //                WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where {campo} like  '%{valor}%' " +
        //                             $"and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
        //            }
        //        }
        //        // Fin: todo por RUBROS

        //        QueryFinal = SelectQuery + WhereQuery + " order by a.arterc";
        //        if (!string.IsNullOrEmpty(QueryFinal))
        //        {
        //            da.SelectCommand = new OleDbCommand(QueryFinal, cnn);
        //            da.Fill(ds);
        //            if (ds.Tables[0].Rows.Count > 0) // mientras que contenga algun registro....
        //            {
        //                _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
        //            }
        //            if (_frmPrincipal.DataGridView1.RowCount > 0)
        //            {
        //                SeleccionarGrilla();
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception($", dentro de 'consultas_relacionadas' {ex.Message}");
        //    }
        //    finally
        //    {
        //        //If TimerVerOfertas.Enabled = False Then TimerVerOfertas.Enabled = True
        //        if (_frmPrincipal.TimerVerOfertas.Enabled == false) _frmPrincipal.TimerVerOfertas.Enabled = true;
        //    }
        //}

        public static void consultas_Relacionales()
        {
            // Declarar variables necesarias
            string campo = string.Empty;
            string valor = string.Empty;

            // Validaciones iniciales
            if (_frmPrincipal == null)
                throw new ArgumentNullException("El formulario principal no puede ser nulo");
            if (cnn == null)
                throw new InvalidOperationException("La conexión a la base de datos no está disponible");

            using var da = new OleDbDataAdapter();
            string QueryFinal = string.Empty;
            string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    //"l.lidesc as Linea," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    //" a.ARMARC as Marca2," + 
                    "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from (listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) ";

            string WhereQuery = string.Empty;
            try
            {
                using var ds = new DataSet();
                // Validar y limpiar inputs
                if (_frmPrincipal.txtDescripcion?.Text != null && _frmPrincipal.txtDescripcion.Text.Length > 100)
                    throw new ArgumentException("La descripción no puede exceder 100 caracteres");
                if (_frmPrincipal.txtModelo?.Text != null && _frmPrincipal.txtModelo.Text.Length > 50)
                    throw new ArgumentException("El modelo no puede exceder 50 caracteres");

                if (_frmPrincipal.cmbRubros.Text == "TODOS" || _frmPrincipal.cmbRubros.Text == " TODOS")
                    _frmPrincipal.cmbRubros.Text = string.Empty;

                if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // por default
                    WhereQuery = " where a.arprim<>'000' ";
                }
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo por descripcion
                    campo = "a.ARDESC";
                    valor = _frmPrincipal.txtDescripcion.Text;
                    string[] arraydesc = _frmPrincipal.txtDescripcion.Text.Split(' ');
                    if (arraydesc.Length > 1)
                    {
                        for (int r = 0; r < arraydesc.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraydesc[r]))
                            {
                                valor = $" where ({campo} like  '%{arraydesc[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraydesc[r]))
                            {
                                valor = $"{valor} AND {campo} like  '%{arraydesc[r]}%'";
                            }
                        }
                        WhereQuery = $"{valor}) and a.arprim<>'000' ";
                    }
                    else // filtro simple
                    {
                        valor = _frmPrincipal.txtDescripcion.Text;
                        WhereQuery = $" where {campo} like  '%{valor}%' and a.arprim<>'000' ";
                    }
                }
                // INICIO: todo por RUBROS
                else if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro
                    WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                }
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro y descripcion
                    campo = "a.ARDESC";
                    valor = _frmPrincipal.txtDescripcion.Text;
                    string[] arraydesc = _frmPrincipal.txtDescripcion.Text.Split(' ');
                    if (arraydesc.Length > 1)
                    {
                        for (int r = 0; r < arraydesc.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraydesc[r]))
                            {
                                valor = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where ({campo} like  '%{arraydesc[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraydesc[r]))
                            {
                                valor = $"{valor} AND {campo} like  '%{arraydesc[r]}%'";
                            }
                        }
                        WhereQuery = $"{valor}) and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                    }
                    else // filtro simple
                    {
                        valor = _frmPrincipal.txtDescripcion.Text;
                        WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where {campo} like  '%{valor}%' " +
                                     $"and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                    }
                }
                else if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo por modelo
                    campo = "a.AROBSE";
                    valor = _frmPrincipal.txtModelo.Text;
                    string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');
                    if (arraymode.Length > 1)
                    {
                        for (int r = 0; r < arraymode.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $" where ({campo} like  '%{arraymode[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $"{valor} AND {campo} like  '%{arraymode[r]}%'";
                            }
                        }
                        WhereQuery = $"{valor}) and a.arprim<>'000' ";
                    }
                    else // filtro simple
                    {
                        valor = _frmPrincipal.txtModelo.Text;
                        WhereQuery = $" where {campo} like  '%{valor}%' and a.arprim<>'000' ";
                    }
                }
                // INICIO: todo por RUBROS
                else if (string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro
                    WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                }
                else if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro y modelo
                    campo = "a.AROBSE";
                    valor = _frmPrincipal.txtModelo.Text;
                    string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');
                    if (arraymode.Length > 1)
                    {
                        for (int r = 0; r < arraymode.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where ({campo} like  '%{arraymode[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $"{valor} AND {campo} like  '%{arraymode[r]}%'";
                            }
                        }
                        WhereQuery = $"{valor}) and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                    }
                    else // filtro simple
                    {
                        valor = _frmPrincipal.txtModelo.Text;
                        WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where {campo} like  '%{valor}%' " +
                                     $"and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                    }
                }
                // NUEVO: Validación para descripción + modelo (sin rubro)
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // descripcion y modelo sin rubro
                    string[] arraydesc = _frmPrincipal.txtDescripcion.Text.Split(' ');
                    string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');

                    string condicionDesc = string.Empty;
                    string condicionModelo = string.Empty;

                    // Construir condición para descripción
                    if (arraydesc.Length > 1)
                    {
                        for (int r = 0; r < arraydesc.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraydesc[r]))
                            {
                                condicionDesc = $"(a.ARDESC like '%{arraydesc[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraydesc[r]))
                            {
                                condicionDesc = $"{condicionDesc} AND a.ARDESC like '%{arraydesc[r]}%'";
                            }
                        }
                        condicionDesc = $"{condicionDesc})";
                    }
                    else
                    {
                        condicionDesc = $"a.ARDESC like '%{_frmPrincipal.txtDescripcion.Text}%'";
                    }

                    // Construir condición para modelo
                    if (arraymode.Length > 1)
                    {
                        for (int r = 0; r < arraymode.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
                            {
                                condicionModelo = $"(a.AROBSE like '%{arraymode[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraymode[r]))
                            {
                                condicionModelo = $"{condicionModelo} AND a.AROBSE like '%{arraymode[r]}%'";
                            }
                        }
                        condicionModelo = $"{condicionModelo})";
                    }
                    else
                    {
                        condicionModelo = $"a.AROBSE like '%{_frmPrincipal.txtModelo.Text}%'";
                    }

                    WhereQuery = $" where {condicionDesc} AND {condicionModelo} and a.arprim<>'000' ";
                }
                // NUEVO: Validación para descripción + modelo + rubro (combinación completa)
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // descripcion, modelo y rubro
                    string[] arraydesc = _frmPrincipal.txtDescripcion.Text.Split(' ');
                    string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');

                    string condicionDesc = string.Empty;
                    string condicionModelo = string.Empty;

                    // Construir condición para descripción
                    if (arraydesc.Length > 1)
                    {
                        for (int r = 0; r < arraydesc.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraydesc[r]))
                            {
                                condicionDesc = $"(a.ARDESC like '%{arraydesc[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraydesc[r]))
                            {
                                condicionDesc = $"{condicionDesc} AND a.ARDESC like '%{arraydesc[r]}%'";
                            }
                        }
                        condicionDesc = $"{condicionDesc})";
                    }
                    else
                    {
                        condicionDesc = $"a.ARDESC like '%{_frmPrincipal.txtDescripcion.Text}%'";
                    }

                    // Construir condición para modelo
                    if (arraymode.Length > 1)
                    {
                        for (int r = 0; r < arraymode.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
                            {
                                condicionModelo = $"(a.AROBSE like '%{arraymode[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraymode[r]))
                            {
                                condicionModelo = $"{condicionModelo} AND a.AROBSE like '%{arraymode[r]}%'";
                            }
                        }
                        condicionModelo = $"{condicionModelo})";
                    }
                    else
                    {
                        condicionModelo = $"a.AROBSE like '%{_frmPrincipal.txtModelo.Text}%'";
                    }

                    WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI where {condicionDesc} AND {condicionModelo} and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                }
                // Fin: todo por RUBROS

                QueryFinal = SelectQuery + WhereQuery + " order by a.arterc";
                if (!string.IsNullOrEmpty(QueryFinal))
                {
                    // Validar longitud de consulta
                    if (QueryFinal.Length > 4000)
                        throw new InvalidOperationException("La consulta es demasiado larga");

                    da.SelectCommand = new OleDbCommand(QueryFinal, cnn);
                    da.Fill(ds);

                    // Validar que el DataSet y la tabla existen
                    if (ds?.Tables?.Count > 0 && ds.Tables[0]?.Rows?.Count > 0)
                    {
                        // Validar límite de resultados
                        if (ds.Tables[0].Rows.Count > 10000)
                            throw new InvalidOperationException("Demasiados resultados. Refine la búsqueda");

                        _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];

                        if (_frmPrincipal.DataGridView1.RowCount > 0)
                        {
                            SeleccionarGrilla();
                        }
                    }
                    else
                    {
                        // Dejar la grilla vacía preservando configuración y encabezados
                        if (_frmPrincipal.DataGridView1.DataSource is System.Data.DataTable dt)
                        {
                            dt.Rows.Clear();
                            _frmPrincipal.DataGridView1.Refresh();
                        }
                        else
                        {
                            _frmPrincipal.DataGridView1.Rows.Clear();
                        }
                    }
                }
            }
            catch (ArgumentException argEx)
            {
                throw new ArgumentException($"Error de validación en consultas_Relacionales: {argEx.Message}");
            }
            catch (InvalidOperationException invEx)
            {
                throw new InvalidOperationException($"Error de operación en consultas_Relacionales: {invEx.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en consultas_Relacionales: {ex.Message}");
            }
            finally
            {
                try
                {
                    // Validar que el timer existe antes de habilitarlo
                    if (_frmPrincipal?.TimerVerOfertas != null && !_frmPrincipal.TimerVerOfertas.Enabled)
                        _frmPrincipal.TimerVerOfertas.Enabled = true;
                }
                catch (Exception finallyEx)
                {
                    // Log del error pero no propagar desde finally
                    System.Diagnostics.Debug.WriteLine($"Error en finally: {finallyEx.Message}");
                }
            }
        }

        public static void consultas_Ofertas()
        {
            using var da = new OleDbDataAdapter();
            string QueryRel = string.Empty;
            try
            {
                using var ds = new DataSet();
                string QueryFinal = string.Empty;
                string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    //"l.lidesc as Linea," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    //" a.ARMARC as Marca2," + 
                    "a.arpre5 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from (listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) ";

                string WhereQuery = string.Empty;

                if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    WhereQuery = " WHERE a.ARTT <> 0 and a.arprim<>'000'";
                }
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    campo = "a.ARDESC";
                    valor = _frmPrincipal.txtDescripcion.Text;
                    WhereQuery = $" where a.ARTT<>0 AND {campo} like  '%{valor}%' and a.arprim<>'000'";
                }
                // INICIO: todo por RUBROS
                else if (string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro
                    WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI where a.ARTT<>0 AND r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' " +
                                 "and a.arprim<>'000'";
                }
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtDescripcion.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro y descripcion
                    WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI where a.ARTT<>0 AND r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' " +
                                 $"and a.ardesc like '%{_frmPrincipal.txtDescripcion.Text}%' and a.arprim<>'000'";
                }
                else if (string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // por default
                    WhereQuery = " where a.arprim<>'000' ";
                }
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo por descripcion
                    campo = "a.AROBSE";
                    valor = _frmPrincipal.txtModelo.Text;
                    string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');
                    if (arraymode.Length > 1)
                    {
                        for (int r = 0; r < arraymode.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $" where ({campo} like  '%{arraymode[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $"{valor} AND {campo} like  '%{arraymode[r]}%'";
                            }
                        }
                        WhereQuery = $"{valor}) and a.arprim<>'000' ";
                    }
                    else // filtro simple
                    {
                        valor = _frmPrincipal.txtModelo.Text;
                        WhereQuery = $" where {campo} like  '%{valor}%' and a.arprim<>'000' ";
                    }
                }
                // INICIO: todo por RUBROS
                else if (string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro
                    WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                }
                else if (!string.IsNullOrEmpty(_frmPrincipal.txtModelo.Text) && !string.IsNullOrEmpty(_frmPrincipal.cmbRubros.Text))
                {
                    // solo rubro y descripcion
                    campo = "a.AROBSE";
                    valor = _frmPrincipal.txtModelo.Text;
                    string[] arraymode = _frmPrincipal.txtModelo.Text.Split(' ');
                    if (arraymode.Length > 1)
                    {
                        for (int r = 0; r < arraymode.Length; r++)
                        {
                            if (r == 0 && !string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where ({campo} like  '%{arraymode[r]}%'";
                            }
                            else if (!string.IsNullOrEmpty(arraymode[r]))
                            {
                                valor = $"{valor} AND {campo} like  '%{arraymode[r]}%'";
                            }
                        }
                        WhereQuery = $"{valor}) and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                    }
                    else // filtro simple
                    {
                        valor = _frmPrincipal.txtModelo.Text;
                        WhereQuery = $" INNER JOIN RUBROS AS r ON a.ARRUB2 = r.RBCODI  where {campo} like  '%{valor}%' " +
                                     $"and r.rbdesc = '{_frmPrincipal.cmbRubros.Text}' and a.arprim<>'000' ";
                    }
                }
                // Fin: todo por RUBROS

                QueryFinal = SelectQuery + WhereQuery + " order by a.arterc";
                if (!string.IsNullOrEmpty(QueryFinal))
                {
                    // Validar longitud de consulta
                    if (QueryFinal.Length > 4000)
                        throw new InvalidOperationException("La consulta es demasiado larga");

                    da.SelectCommand = new OleDbCommand(QueryFinal, cnn);
                    da.Fill(ds);

                    // Validar que el DataSet y la tabla existen
                    if (ds?.Tables?.Count > 0 && ds.Tables[0]?.Rows?.Count > 0)
                    {
                        // Validar límite de resultados
                        if (ds.Tables[0].Rows.Count > 10000)
                            throw new InvalidOperationException("Demasiados resultados. Refine la búsqueda");

                        _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];

                        if (_frmPrincipal.DataGridView1.RowCount > 0)
                        {
                            SeleccionarGrilla();
                        }
                    }
                    else
                    {
                        // Dejar la grilla vacía preservando configuración y encabezados
                        if (_frmPrincipal.DataGridView1.DataSource is System.Data.DataTable dt)
                        {
                            dt.Rows.Clear();
                            _frmPrincipal.DataGridView1.Refresh();
                        }
                        else
                        {
                            _frmPrincipal.DataGridView1.Rows.Clear();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                throw new Exception($", dentro de 'consultas_Ofertas' {ex.Message}");
            }
            finally
            {
                //If TimerVerOfertas.Enabled = False Then TimerVerOfertas.Enabled = True
                if (_frmPrincipal.TimerVerOfertas.Enabled == false) _frmPrincipal.TimerVerOfertas.Enabled = true;
            }
        }


        public static string GenerarCondicionMultipalabra(string campo, string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return "";

            string[] palabras = texto.Trim().Split(' ')
                                      .Where(p => !string.IsNullOrWhiteSpace(p))
                                      .ToArray();

            if (palabras.Length == 0)
                return "";

            var condiciones = palabras.Select(p => $"{campo} like '%{p.Trim()}%'");
            return "(" + string.Join(" AND ", condiciones) + ")";
        }
        public static void LeerConfig()
        {
            try
            {
                if (File.Exists(@"C:\DistribuidoraPalomarNet\Config.ini"))
                {
                    // Configurar ruta de base de datos de usuarios
                    ini_ruta_usu = @"C:\DistribuidoraPalomarNet\cliente.dat";

                    string cadena;
                    int largoCadena;
                    using (var leer = new StreamReader(@"C:\DistribuidoraPalomarNet\Config.ini"))
                    {
                        // ini_emisor_email
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 6) == "Emisor")
                        {
                            largoCadena = cadena.Length;
                            ini_emisor_email = cadena.Substring(7, largoCadena - 7);
                        }

                        // ini_pass_email
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 4) == "Pass")
                        {
                            largoCadena = cadena.Length;
                            ini_pass_email = cadena.Substring(5, largoCadena - 5);
                        }

                        // ini_host_email
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 4) == "Host")
                        {
                            largoCadena = cadena.Length;
                            ini_host_email = cadena.Substring(5, largoCadena - 5);
                        }

                        // ini_port_email
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 4) == "Port")
                        {
                            largoCadena = cadena.Length;
                            ini_port_email = cadena.Substring(5, largoCadena - 5);
                        }

                        // ini_enable_ssl
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 9) == "EnableSsl")
                        {
                            largoCadena = cadena.Length;
                            string xxx = cadena.Substring(10, largoCadena - 10);
                            ini_enable_ssl = Convert.ToBoolean(cadena.Substring(10, largoCadena - 10));
                        }

                        // ini_receptor_email
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 8) == "Receptor")
                        {
                            largoCadena = cadena.Length;
                            ini_receptor_email = cadena.Substring(9, largoCadena - 9);
                        }

                        // *******************Inicio Config FTP*********************
                        // ini_ruta_ftp
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 7) == "RutaFTP")
                        {
                            largoCadena = cadena.Length;
                            ini_ruta_ftp = cadena.Substring(8, largoCadena - 8);
                        }

                        // usuFTP
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 6) == "UsuFTP")
                        {
                            largoCadena = cadena.Length;
                            ini_usu_ftp = cadena.Substring(7, largoCadena - 7);
                        }

                        // passFTP
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 7) == "PassFTP")
                        {
                            largoCadena = cadena.Length;
                            ini_pass_ftp = cadena.Substring(8, largoCadena - 8);
                        }

                        // archFTP
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 7) == "ArchFTP")
                        {
                            largoCadena = cadena.Length;
                            ini_arch_ftp = cadena.Substring(8, largoCadena - 8);
                        }

                        // ***********************Rutas Base Catalogo**************************
                        // ruta Base catalogo
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 8) == "RutaBase")
                        {
                            largoCadena = cadena.Length;
                            ini_ruta_base = cadena.Substring(9, largoCadena - 9);
                        }

                        // ruta base descargada
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 16) == "RutaBaseDescarga")
                        {
                            largoCadena = cadena.Length;
                            ini_ruta_base_descarga = cadena.Substring(17, largoCadena - 17);
                        }


                        // ***********************Para Semaforo de Stock**************************
                        // CondRojo
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 8) == "CondRojo")
                        {
                            largoCadena = cadena.Length;
                            ini_cond_rojo = Convert.ToInt32(cadena.Substring(9, largoCadena - 9));
                        }

                        // CondAmarillo
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 12) == "CondAmarillo")
                        {
                            largoCadena = cadena.Length;
                            ini_cond_amarillo = Convert.ToInt32(cadena.Substring(13, largoCadena - 13));
                        }

                        // CondVerde
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 9) == "CondVerde")
                        {
                            largoCadena = cadena.Length;
                            ini_cond_verde = Convert.ToInt32(cadena.Substring(10, largoCadena - 10));
                        }
                        // RutaBanner
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 10) == "RutaBanner")
                        {
                            largoCadena = cadena.Length;
                            ini_ruta_banner = Convert.ToString(cadena.Substring(11, largoCadena - 11));
                            LeerRutaBanner(ini_ruta_banner);
                        }
                        // RutaGral
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 8) == "RutaGral")
                        {
                            largoCadena = cadena.Length;
                            ini_ruta_gral = cadena.Substring(9, largoCadena - 9);
                        }
                        // apliFTP
                        cadena = leer.ReadLine();
                        if (cadena.Substring(0, 7) == "ApliFTP")
                        {
                            largoCadena = cadena.Length;
                            ini_apli_ftp = cadena.Substring(8, largoCadena - 8);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No se encuentra el archivo de Configuración.\r\n" +
                                  "La opción de Actualización y envio de Pedido,\r\n" +
                                  "no funcionaran correctamente", "Sin info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error dentro del proceso 'LeerConfig': " + ex.Message);
            }
        }

        public static async void LeerRutaBanner(string ini_ruta_banner)
        {
            try
            {
                string texto = await ObtenerTextoDesdeUrl(ini_ruta_banner);
                //Console.WriteLine("Contenido del archivo:");
                //Console.WriteLine(texto);
                _textoBanner = texto;

                try
                {
                    if (_frmPrincipal != null)
                        _frmPrincipal.ActualizarTexto(_textoBanner);
                }
                catch
                {
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error: {ex.Message}");
            }
        }
        public static async Task<string> ObtenerTextoDesdeUrl(string url)
        {
            using (HttpClient cliente = new HttpClient())
            {
                // Hacer una solicitud GET y obtener el contenido como string
                HttpResponseMessage respuesta = await cliente.GetAsync(url);
                respuesta.EnsureSuccessStatusCode(); // Lanza una excepción si hay un error HTTP
                return await respuesta.Content.ReadAsStringAsync();
            }
        }
        public static string FechaActualizacion()
        {
            var cmaccess = new OleDbCommand();
            OleDbDataReader draccess;
            try
            {
                ConectarDBArti();
                cmaccess.Connection = cnn;
                cmaccess.CommandType = CommandType.Text;
                cmaccess.CommandText = "select top 1 arpp from listas where arpp is not null";
                draccess = cmaccess.ExecuteReader();
                if (draccess.Read())
                {
                    fecha = Convert.ToDateTime(draccess["arpp"].ToString()).ToString("dd/MM/yyyy");
                }
                draccess.Close();
                if (cnn.State == ConnectionState.Open) cnn.Close();
                return fecha;
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                return fecha;
            }
            finally
            {
                cmaccess.Dispose();
            }
        }

        public static string FechaArchivoFTP(string Filename, string NewFileName, string UserNameFtp, string PassFTP, string URI)
        {
            FtpWebRequest Ftpconection = (FtpWebRequest)FtpWebRequest.Create(URI + "/" + Filename);
            // Creamos la nueva conexion del tipo web request donde URI es la direccion del servidor
            Ftpconection.Credentials = new NetworkCredential(UserNameFtp, PassFTP);
            string dte = "01/01/1900";
            try
            {
                //Asignamos las credenciales para el servidor
                Ftpconection.Method = WebRequestMethods.Ftp.GetDateTimestamp;
                using (FtpWebResponse response = (FtpWebResponse)Ftpconection.GetResponse())
                {
                    dte = response.LastModified.ToString("dd/MM/yyyy");
                }
                return dte.ToString();
            }
            catch (Exception)
            {
                return dte.ToString();
            }
        }

        public static void articulo_a_Pedido_Update(string cod, decimal pun, string Cantidad, decimal pt)
        {
            try
            {
                ConectarDBUsuario();
                using var cmd = new OleDbCommand();
                cmd.Connection = cnnU;
                cmd.CommandText = @"UPDATE PEDIDOS SET 
                    CANTIDAD = @cantidad,
                    PRECIO_UNITARIO = @precioUnitario,
                    PRECIO_TOTAL = @precioTotal 
                    WHERE CODIGO = @codigo";
                
                cmd.Parameters.AddWithValue("@cantidad", Cantidad);
                cmd.Parameters.AddWithValue("@precioUnitario", pun);
                cmd.Parameters.AddWithValue("@precioTotal", pt);
                cmd.Parameters.AddWithValue("@codigo", cod);
                
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en articulo_a_Pedido_Update: {ex.Message}");
            }
        }

        public static void ActualizacionXFTP()
        {
            bool ret;
            Form frmProgreso = null;
            Label lblProgreso = null;
            ProgressBar pbProgreso = null;

            try
            {
                _frmPrincipal.Timer1.Start();
                _frmPrincipal.toolStripProgressBar1.Visible = true;
                Application.DoEvents();

                frmProgreso = new Form();
                frmProgreso.StartPosition = FormStartPosition.CenterScreen;
                frmProgreso.FormBorderStyle = FormBorderStyle.FixedDialog;
                frmProgreso.ControlBox = false;
                frmProgreso.MinimizeBox = false;
                frmProgreso.MaximizeBox = false;
                frmProgreso.ShowInTaskbar = false;
                frmProgreso.TopMost = true;
                frmProgreso.Width = 520;
                frmProgreso.Height = 150;
                frmProgreso.Text = "Actualizacion Masiva";

                lblProgreso = new Label();
                lblProgreso.AutoSize = false;
                lblProgreso.Left = 16;
                lblProgreso.Top = 16;
                lblProgreso.Width = frmProgreso.ClientSize.Width - 32;
                lblProgreso.Height = 40;
                lblProgreso.Text = "Preparando actualizacion...";

                pbProgreso = new ProgressBar();
                pbProgreso.Left = 16;
                pbProgreso.Top = 70;
                pbProgreso.Width = frmProgreso.ClientSize.Width - 32;
                pbProgreso.Height = 22;
                pbProgreso.Style = ProgressBarStyle.Marquee;
                pbProgreso.MarqueeAnimationSpeed = 30;

                frmProgreso.Controls.Add(lblProgreso);
                frmProgreso.Controls.Add(pbProgreso);
                frmProgreso.Show();
                Application.DoEvents();

                ret = DescargaArchivo(ini_arch_ftp,
                                    @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\" + ini_arch_ftp,
                                    ini_usu_ftp,
                                    ini_pass_ftp,
                                    ini_ruta_ftp);

                if (ret)
                {
                    if (lblProgreso != null)
                    {
                        lblProgreso.Text = "Iniciando actualizacion...";
                        Application.DoEvents();
                    }

                    System.Diagnostics.Process.Start(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\" + ini_arch_ftp, "Update");

                    Application.DoEvents();

                    if (File.Exists(ini_ruta_base_descarga))
                    {
                        if (lblProgreso != null)
                        {
                            lblProgreso.Text = "Actualizando base de datos...";
                            Application.DoEvents();
                        }

                        Migra = UpdateBase(ini_ruta_base_descarga, ini_ruta_base);

                        if (Migra)
                        {
                            if (lblProgreso != null)
                            {
                                lblProgreso.Text = "Finalizando actualizacion...";
                                Application.DoEvents();
                            }

                            Application.DoEvents();
                            Update_FechaDB();
                            Cargar_Usuario();
                            _frmPrincipal.ApplyEstadoClienteToUI();

                            if (!ClienteAutorizado)
                            {
                                _frmPrincipal.ckbOfertas.Checked = false;
                                _frmPrincipal.lblFechaAct.Text = "Actualizado al  " + FechaActualizacion();
                                BuscarActualizacionDisponible();
                                _frmPrincipal.cmbRubros.DropDownWidth = 280;
                                _frmPrincipal.lblAlertaUpdate.Visible = false;
                                Migra = false;
                                return;
                            }

                            if (lblProgreso != null)
                            {
                                lblProgreso.Text = "Recalculando precios...";
                                Application.DoEvents();
                            }

                            Update_Precios_Gral(); //recalculo los precios
                            _frmPrincipal.ckbOfertas.Checked = false;
                            funcion_inicial(1);
                            //funcion_inicial_carrito();
                            aplicar_desc_Cliente(); //Cargo Descuentos 1 y 2
                            _frmPrincipal.lblFechaAct.Text = "Actualizado al  " + FechaActualizacion();
                            BuscarActualizacionDisponible();
                            _frmPrincipal.cmbRubros.DropDownWidth = 280;
                            if (lblProgreso != null)
                            {
                                lblProgreso.Text = "Actualizacion realizada con exito.";
                                Application.DoEvents();
                            }
                            _frmPrincipal.lblAlertaUpdate.Visible = false;
                            Migra = false;
                        }
                        else
                        {
                            MessageBox.Show("No Se pudo Terminar la Actualizacion entre Bases");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se encontro en Archivo a actualizar");
                    }
                }
                else
                {
                    MessageBox.Show("No se Pudo descargar el archivo de Actualizacion");
                }

                Application.DoEvents();
                _frmPrincipal.toolStripProgressBar1.Visible = false;
                _frmPrincipal.Timer1.Stop();
            }
            catch (Exception ex)
            {
                _frmPrincipal.toolStripProgressBar1.Visible = false;
                _frmPrincipal.Timer1.Stop();
                MessageBox.Show("Error dentro de Menu:Archivo-->Actualizar Base" + ex.Message,
                               "Alerta",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
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
        }
        public static bool UpdateBase(string Origen, string Destino)
        {
            try
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();

                if (File.Exists(Destino))
                {
                    File.Delete(Destino);
                }

                File.Copy(Origen, Destino);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public static void Update_FechaDB()
        {
            using (OleDbCommand cmaccessUpdate = new OleDbCommand())
            {
                try
                {
                    ConectarDBArti();
                    cmaccessUpdate.Connection = cnn;
                    cmaccessUpdate.CommandType = CommandType.Text;

                    string QueryUpdate = "update listas set arpp='" +
                                       FechaArchivoFTP(ini_arch_ftp,
                                                     @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\" + ini_arch_ftp,
                                                     ini_usu_ftp,
                                                     ini_pass_ftp,
                                                     ini_ruta_ftp) + "'";

                    cmaccessUpdate.CommandText = QueryUpdate;
                    cmaccessUpdate.ExecuteNonQuery();

                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                }
                catch (Exception ex)
                {
                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();

                    throw new Exception(", dentro de 'Update_FechaDB': " +
                                      (ex.InnerException != null ? ex.InnerException.ToString() : "") +
                                      " - " + ex.Message);
                }
            }
        }

        public static bool BuscarActualizacionDisponible()
        {
            bool ret=false;
            string FechaFTP;
            try
            {
                FechaFTP = FechaArchivoFTP(ini_arch_ftp, ini_ruta_gral + "\\DistribuidoraPalomarNet\\Descargas\\" + ini_arch_ftp, ini_usu_ftp, ini_pass_ftp, ini_ruta_ftp);
                if (Convert.ToDateTime(FechaFTP) != Convert.ToDateTime("01/01/1900"))
                {
                    //MessageBox.Show("Fecha Base en FTP: " + FechaFTP.ToString())
                    if (Convert.ToDateTime(fecha) < Convert.ToDateTime(FechaFTP))
                    {
                        _frmPrincipal.lblAlertaUpdate.Visible = true;
                        _frmPrincipal.lblAlertaUpdate.Text = "Hay una Actualizacion disponible, por favor actualice la base de datos";
                        ret=true;
                    }
                    else
                    {
                        _frmPrincipal.lblAlertaUpdate.Visible = false;
                        ret=false;
                    }
                }
                else
                {
                    _frmPrincipal.lblAlertaUpdate.Visible = false;
                    ret=false;
                }
            }
            catch (Exception ex)
            {
                _frmPrincipal.lblAlertaUpdate.Visible = false;
                throw new Exception("Error dentro del proceso 'BuscarActualizacionDisponible': " + ex.Message.ToString());
                ret= false;
            }
            return ret;

        }

        public static void ProcesoCorreo()
        {
            try
            {
                // Intentar generar el documento.
                Document doc = new Document(PageSize.A4, 10, 10, 10, 10);
                // Path que guarda el reporte en el escritorio de windows (Desktop).
                string filename = $@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\pedidos\Remito {Module.US_CODI} {DateTime.Now.ToString().Replace("/", "-").Replace(":", "")}.pdf";

                using (FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                {
                    PdfWriter.GetInstance(doc, file);
                    doc.Open();
                    ExportarDatosPDF(doc);
                    doc.Close();
                }

                if (_frmCarroCompras.cboxNoEnviar.CheckState != CheckState.Checked)
                {
                    // envio el archivo generado
                    string emisor, pass, mensaje, asunto, receptor;

                    emisor = ini_emisor_email;
                    pass = ini_pass_email;
                    mensaje = "";
                    //mensaje = "<h3>Obses:</h3>" +
                    //          "<p style='font-family:arial;font-size:11px;'>" +
                    //          "Enviaste el pedido de cotizacion correctamente. <br>" +
                    //          "Un ejecutivo de ventas te va a contactar a la brevedad.<br>" +
                    //          "</p>";
                    
                    //asunto = "Pedido Generado desktop. Cliente " + US_CODI;

                    asunto = !string.IsNullOrWhiteSpace(_frmCarroCompras.txtNotaPedido.Text)
                    ? "Pedido Catalogo. VER NOTA - Cliente " + US_CLRAZO
                    : "Pedido Catalogo. Cliente " + US_CLRAZO;

                    receptor = ini_receptor_email;

                    if (!string.IsNullOrEmpty(US_EMAIL))//_frmPrincipal.lblemail.Text))
                    {
                        if (ValidaEMail(US_EMAIL))//_frmPrincipal.lblemail.Text))
                        {
                            receptor = receptor + "," + US_EMAIL;// _frmPrincipal.lblemail.Text;
                        }
                    }

                    Correo.EnviarCorreo(emisor, pass, mensaje, asunto, receptor, filename);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Dentro de 'Modulo.ProcesoCorreo': " + ex.Message,
                                "Error inesperado",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                return;
            }
        }
        public static bool ValidaEMail(string EMail)
        {
            // Como primera regla un correo electronico debe contener la @
            if (!EMail.Contains("@"))
            {
                return false;
            }
            // Dividimos la cadena en secciones, obviamente estas deben ser 2
            // el usuario y el host, utilizamos como separador la @
            string[] SeccionesEMail = EMail.Split('@');

            // Ahora verificamos que evidentemente solo sean 2 secciones, ya que
            // en caso contrario eso significa que hay mas de una @ y eso es incorrecto
            if (SeccionesEMail.Length != 2)
            {
                return false;
            }
            // Ahora verificamos que la segunda seccion de la cadena de correo contenga
            // al menos un punto, ya que la seccion del dominio debe contener el punto
            // Podemos establecer un tamaño minimo para el dominio en este caso le puse 3
            if (!SeccionesEMail[1].Contains(".") || !(SeccionesEMail[1].Length >= 3))
            {
                return false;
            }
            return true;
        }
        //public static void ExportarDatosPDF(Document document)
        //{
        //    try
        //    {
        //        if (_frmCarroCompras == null) throw new InvalidOperationException("Formulario de carro de compras no inicializado");

        //        // Se crea un objeto PDFTable con el numero de columnas del DataGridView
        //        var datatable = new PdfPTable((_frmCarroCompras.gridPedidos.ColumnCount - 2));
        //        datatable.DefaultCell.Padding = 3;
        //        float[] headerwidths = { 15, 40, 15, 15, 15 };
        //        datatable.SetWidths(headerwidths);
        //        datatable.WidthPercentage = 100;
        //        datatable.DefaultCell.BorderWidth = 2;
        //        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

        //        // Se configura el título de las columnas
        //        foreach (DataGridViewColumn column in _frmCarroCompras.gridPedidos.Columns)
        //        {
        //            if (column.Index < _frmCarroCompras.gridPedidos.ColumnCount - 2)
        //            {
        //                var cell = new PdfPCell(new Phrase(column.HeaderText));
        //                cell.BackgroundColor = new BaseColor(240, 240, 240);
        //                datatable.AddCell(cell);
        //            }
        //        }

        //        // Se agregan las filas del DataGridView
        //        foreach (DataGridViewRow row in _frmCarroCompras.gridPedidos.Rows)
        //        {
        //            for (int i = 0; i < row.Cells.Count - 2; i++)
        //            {
        //                datatable.AddCell(row.Cells[i].Value?.ToString() ?? "");
        //            }
        //        }

        //        document.Add(datatable);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception($"Error en ExportarDatosPDF: {ex.Message}");
        //    }
        //}
        public static void ExportarDatosPDF(Document document)
        {
            try
            {
                // Se crea un objeto PDFTable con el numero de columnas del DataGridView.
                PdfPTable datatable = new PdfPTable((_frmCarroCompras.gridPedidos.ColumnCount - 2)); //reemplazo -1 por -3 para sacar los precios#lo agregue gmedina

                // Se asignan algunas propiedades para el diseño del PDF.
                datatable.DefaultCell.Padding = 1;
                float[] headerwidths = GetColumnasSize(_frmCarroCompras.gridPedidos);
                datatable.SetWidths(headerwidths);

                datatable.WidthPercentage = 100;
                datatable.DefaultCell.BorderWidth = 1;
                datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;

                // se crea el logo del pdf
                iTextSharp.text.Image imagendemo; // Declaracion de una imagen
                imagendemo = iTextSharp.text.Image.GetInstance(@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\imagenes\LogoImpresion.png"); // Dirreccion a la imagen que se hace referencia
                imagendemo.ScaleAbsoluteWidth(150); // Ancho de la imagen
                imagendemo.ScaleAbsoluteHeight(45); // Altura de la imagen
                                                    // Phrase texto2 = new Phrase("Catalogo Pedido", new Font(Font.Name = "Tahoma", 20, Font.Bold));

                // Se crea el encabezado en el PDF.
                Paragraph encabezado = new Paragraph("Reporte de Pedido. Cliente " + Module.US_CODI, FontFactory.GetFont(FontFactory.COURIER, 14, iTextSharp.text.Font.BOLD));
                // Se crea el texto abajo del encabezado.
                Phrase texto = new Phrase("Realizado la fecha : " + DateTime.Now, FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.BOLD));

                // Se capturan los nombres de las columnas del DataGridView.
                for (int i = 0; i < _frmCarroCompras.gridPedidos.ColumnCount; i++)
                {
                    if (i != 0 && i != 4) // saco columnas de id y precios#lo agregue gmedina
                    {
                        datatable.AddCell(new PdfPCell(new Paragraph(_frmCarroCompras.gridPedidos.Columns[i].HeaderText, FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.ITALIC))));
                    }
                }

                datatable.HeaderRows = 1;
                datatable.DefaultCell.BorderWidth = 1;

                // Se generan las columnas del DataGridView.
                for (int i = 0; i < _frmCarroCompras.gridPedidos.RowCount; i++)
                {
                    for (int j = 0; j < _frmCarroCompras.gridPedidos.ColumnCount; j++)
                    {
                        if (j != 0 && j != 4) // saco columnas de id y precios#lo agregue gmedina
                        {
                            datatable.AddCell(new PdfPCell(new Paragraph(_frmCarroCompras.gridPedidos[j, i].Value.ToString(), FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.ITALIC))));
                        }
                        _frmCarroCompras.PictureBox1.Refresh();
                    }
                    datatable.CompleteRow();
                }

                // Se agrega el PDFTable al documento.
                document.Add(imagendemo);
                // document.Add(texto2);
                document.Add(encabezado);
                document.Add(texto);
                document.Add(datatable);
            }
            catch (Exception ex)
            {
                throw new Exception(", dentro de la acción 'ExportarDatosPDF': " + ex.Message.ToString());
            }
        }
        public static float[] GetColumnasSize(DataGridView dg)
        {
            // funcion para obtener el tamaño de la columnas del datagridview
            float[] values = new float[dg.ColumnCount - 2]; // reemplazo -1 por -4#lo agregue gmedina

            // For loop commented out as in the original VB code
            //for (int i = 0; i <= dg.ColumnCount - 4; i++) // reemplazo -1 por -4#lo agregue gmedina
            //{
            //    values[i] = (float)dg.Columns[i + 1].Width;
            //}

            // Hardcoded values as in the original VB code
            values[0] = (float)105;
            values[1] = (float)95;
            values[2] = (float)420;
            values[3] = (float)215;
            values[4] = (float)125;
            values[5] = (float)175;

            return values;
        }
        public static void Update_Precios_Gral()
        {
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");
                
                _frmPrincipal.DataGridView1.Enabled = false;
                Cargar_Usuario();
                Update_Precios();
                Update_Precios_Ofertas();
                funcion_inicial(0);
                _frmPrincipal.DataGridView1.DefaultCellStyle.BackColor = Color.White;
                _frmPrincipal.DataGridView1.DefaultCellStyle.ForeColor = Color.Black;
                _frmPrincipal.DataGridView1.Refresh();
                _frmPrincipal.DataGridView1.Enabled = true;
            }
            catch (Exception)
            {
                _frmPrincipal.DataGridView1.Enabled = true;
            }
        }

        public static void Update_Precios()
        {
            try
            {
                Application.DoEvents();
                Update_PrecioCosto();
                Update_PrecioSug();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en Update_Precios: {ex.InnerException} - {ex.Message}");
            }
        }
        public static void Update_Precios_Ofertas()
        {
            string Desc_Calc_PList = "0";
            try
            {
                Application.DoEvents();
                Update_PrecioCosto_Ofertas();
                Update_PrecioSug_Ofertas();
            }
            catch (Exception ex)
            {
                throw new Exception($", dentro de 'Update_Precios_Ofertas': {(ex.InnerException?.ToString() ?? "")} - {ex.Message}");
            }
        }

        public static void Pintar_Ofertas()
        {
            if (_frmPrincipal.DataGridView1.RowCount > 0)
            {
                foreach (DataGridViewRow row in _frmPrincipal.DataGridView1.Rows)
                {
                    if (row.Cells["aroo"].Value != null && row.Cells["aroo"].Value.ToString() == "1")
                    {
                        row.DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
                _frmPrincipal.DataGridView1.Refresh();
            }
        }

        public static void Despintar_Ofertas()
        {
            if (_frmPrincipal.DataGridView1.RowCount > 0)
            {
                _frmPrincipal.DataGridView1.DefaultCellStyle.BackColor = Color.White;
                _frmPrincipal.DataGridView1.DefaultCellStyle.ForeColor = Color.Black;
                _frmPrincipal.DataGridView1.Refresh();
                _frmPrincipal.DataGridView1.Columns[5].Visible = false;
            }
        }

        public static void GrabarError(string mensaje)
        {
            try
            {
                File.AppendAllText("error.log", mensaje + Environment.NewLine);
            }
            catch
            {
                // Si falla el logging, no podemos hacer mucho más
            }
        }

        public static async Task UpdateValorDBLocalAsync(string query)
        {
            try
            {
                ConectarDBArti();
                using var cmd = new OleDbCommand(query, cnn);
                await cmd.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en UpdateValorDBLocal: {ex.Message}");
            }
            finally
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
        }

        public static void UpdateValorDBLocal(string query)
        {
            try
            {
                ConectarDBUsuario();
                using var cmd = new OleDbCommand(query, cnnU);
                cmd.ExecuteNonQuery();
                
                if (cnnU.State == ConnectionState.Open)
                    cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
                throw new Exception($"Error en UpdateValorDBLocal: {ex.Message}");
            }
        }

        public static void CargarComboRubros()
        {
            Application.DoEvents();
            ConectarDBArti();
            string sql = "SELECT DISTINCT R.rbcodi, R.rbdesc FROM rubros AS R INNER JOIN [listas] AS LT ON R.RBCODI = LT.ARRUB2 WHERE R.RBCODI > 100 ORDER BY R.rbdesc";
            var da = new OleDbDataAdapter(sql, cnn);
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");
                
                _frmPrincipal.cmbRubros.Items.Clear();
                var ds = new DataSet();
                Application.DoEvents();
                da.Fill(ds, "rubros");
                
                foreach (DataRow fila in ds.Tables["rubros"].Rows)
                {
                    Application.DoEvents();
                    _frmPrincipal.cmbRubros.Items.Add(fila["rbdesc"]);
                }
                _frmPrincipal.cmbRubros.Items.Insert(0, "TODOS");
                ds.Dispose();
                if (cnn.State == ConnectionState.Open) cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception(", dentro de 'CargarComboRubros': {ex}");
            }
            finally
            {
                da.Dispose();
            }
        }

        public static void Cargar_Usuario(bool promptIfMissing = true)
        {
            OleDbCommand cmUsu = new OleDbCommand();
            OleDbDataReader drUsu = null;
            bool datos;

            try
            {
                ConectarDBUsuario();
                cmUsu.Connection = cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = "SELECT US_CODI, US_CLRAZO, US_CUIT, US_MAIL, US_CLCIVA, US_CLVEND, US_CLCPAG from USUARIO";
                drUsu = cmUsu.ExecuteReader();
                if (drUsu.Read())
                {
                    datos = true;
                    US_CODI = drUsu["US_CODI"].ToString();
                    US_CUIT = drUsu["US_CUIT"].ToString();
                    US_EMAIL = drUsu["US_MAIL"].ToString();
                    US_CLCIVA = Convert.ToInt32(drUsu["US_CLCIVA"]);
                    US_CLVEND = Convert.ToInt32(drUsu["US_CLVEND"]);
                    US_CLRAZO = drUsu["US_CLRAZO"].ToString();
                    US_CLCPAG = Convert.ToInt32(drUsu["US_CLCPAG"]);

                    ValidarCuitCliente(US_CODI, US_CUIT);
                    drUsu.Close();
                    if (cnnU.State == ConnectionState.Open) cnnU.Close();
                }
                else
                {
                    drUsu.Close();
                    if (cnnU.State == ConnectionState.Open) cnnU.Close();
                    US_CODI = "";
                    US_CUIT = "";
                    US_EMAIL = "";
                    US_CLCIVA = 0;
                    US_CLVEND = 0;
                    US_CLRAZO = "";
                    US_CLCPAG = 0;
                    // Solo abrir el formulario si corresponde y no hay una instancia abierta
                    if (promptIfMissing && !Application.OpenForms.OfType<frmDatosUsuario>().Any())
                    {
                        var frmDatosUsuario = new frmDatosUsuario();
                        frmDatosUsuario.ShowDialog();
                    }
                }
            }
            catch (Exception ex)
            {
                if (cnnU.State == ConnectionState.Open) cnnU.Close();
                // Throw New Exception(", dentro de 'Cargar_Usuario': " + ex.Message.ToString)
            }
            finally
            {
                cmUsu.Dispose();
            }
        }

        public static bool ValidarCampoNumerico(string valor)
        {
            return !string.IsNullOrWhiteSpace(valor) && float.TryParse(valor.Trim(), out _);
        }

        public static bool ValidarCampos()
        {
            try
            {
                if (_frmDatosUsuario == null) throw new InvalidOperationException("Formulario de datos de usuario no inicializado");

                if (string.IsNullOrWhiteSpace(_frmDatosUsuario.txtCodCliente.Text))
                    return false;

                if (string.IsNullOrWhiteSpace(_frmDatosUsuario.txtCuitCliente.Text))
                    return false;

                if (!ValidarCampoNumerico(_frmDatosUsuario.txtDesc1.Text))
                    return false;

                if (!ValidarCampoNumerico(_frmDatosUsuario.txtDesc2.Text))
                    return false;

                if (!ValidarCampoNumerico(_frmDatosUsuario.txtBene.Text))
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en ValidarCampos: {ex.Message}");
            }
        }

        public static void CargarDatos()
        {
            try
            {
                if (_frmDatosUsuario == null) throw new InvalidOperationException("Formulario de datos de usuario no inicializado");

                ConectarDBUsuario();
                using var cmd = new OleDbCommand("SELECT US_CODI, US_CUIT, US_MAIL, US_DESCUENTO1, US_DESCUENTO2, US_BENEFICIO from USUARIO", cnnU);
                using var dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    _frmDatosUsuario.txtCodCliente.Text = dr["US_CODI"].ToString();
                    _frmDatosUsuario.txtCuitCliente.Text = dr["US_CUIT"].ToString();
                    _frmDatosUsuario.txtEmail.Text = dr["US_MAIL"].ToString();
                    _frmDatosUsuario.txtDesc1.Text = dr["US_DESCUENTO1"].ToString();
                    _frmDatosUsuario.txtDesc2.Text = dr["US_DESCUENTO2"].ToString();
                    _frmDatosUsuario.txtBene.Text = dr["US_BENEFICIO"].ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en CargarDatos: {ex.Message}");
            }
            finally
            {
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
            }
        }

        public static bool ValidarCodiCliente(string codi)
        {
            OleDbCommand cmaccess = new OleDbCommand();
            OleDbDataReader draccess = null;
            try
            {
                if (_frmDatosUsuario == null) throw new InvalidOperationException("Formulario de datos de usuario no inicializado");

                ConectarDBArti();
                cmaccess.Connection = cnn;
                cmaccess.CommandType = CommandType.Text;
                cmaccess.CommandText = $"select clcodi,clrazo,clciva,clvend,clcpag from clientes WHERE clcodi={codi}";
                draccess = cmaccess.ExecuteReader();

                if (draccess.Read())
                {
                    _frmDatosUsuario.lblUsuario.Text = draccess["clrazo"].ToString();
                    US_CLRAZO = draccess["clrazo"].ToString();

                    US_CLCIVA = !string.IsNullOrEmpty(draccess["clciva"].ToString()) ? 
                        Convert.ToInt32(draccess["clciva"]) : 0;

                    US_CLVEND = !string.IsNullOrEmpty(draccess["clvend"].ToString()) ? 
                        Convert.ToInt32(draccess["clvend"]) : 0;

                    US_CLCPAG = !string.IsNullOrEmpty(draccess["clcpag"].ToString()) ? 
                        Convert.ToInt32(draccess["clcpag"]) : 0;

                    ExisteCodi_Cliente = true;
                }
                else
                {
                    US_CLRAZO = string.Empty;
                    US_CLCIVA = 0;
                    US_CLVEND = 0;
                    US_CLCPAG = 0;
                    ExisteCodi_Cliente = false;
                }

                draccess.Close();
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();

                return ExisteCodi_Cliente;
            }
            catch (Exception)
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                return false;
            }
            finally
            {
                if (draccess != null && !draccess.IsClosed)
                    draccess.Close();
                cmaccess.Dispose();
            }
        }

        public static bool ValidarCuitCliente(string codi, string cuit)
        {
            OleDbCommand cmaccess = new OleDbCommand();
            OleDbDataReader draccess = null;
            try
            {
                ConectarDBArti();
                cmaccess.Connection = cnn;
                cmaccess.CommandType = CommandType.Text;
                cmaccess.CommandText = "SELECT clcodi, clrazo, clciva, clvend, clcpag, cldes1, cldes2 FROM clientes WHERE clcodi=@codi AND clcuit=@cuit";

                if (!int.TryParse((codi ?? string.Empty).Trim(), out int cod))
                    cod = 0;
                if (!double.TryParse((cuit ?? string.Empty).Trim(), out double cuitNum))
                    cuitNum = 0;

                cmaccess.Parameters.Clear();
                cmaccess.Parameters.AddWithValue("@codi", cod);
                cmaccess.Parameters.AddWithValue("@cuit", cuitNum);
                draccess = cmaccess.ExecuteReader();

                if (draccess.Read())
                {
                    US_CLRAZO = draccess["clrazo"].ToString();

                    US_CLCIVA = !string.IsNullOrEmpty(draccess["clciva"].ToString()) ? 
                        Convert.ToInt32(draccess["clciva"]) : 0;

                    US_CLVEND = !string.IsNullOrEmpty(draccess["clvend"].ToString()) ? 
                        Convert.ToInt32(draccess["clvend"]) : 0;

                    US_CLCPAG = !string.IsNullOrEmpty(draccess["clcpag"].ToString()) ? 
                        Convert.ToInt32(draccess["clcpag"]) : 0;

                    ClienteDescuento1 = draccess["cldes1"] != DBNull.Value ? Convert.ToDecimal(draccess["cldes1"]) : 0;
                    ClienteDescuento2 = draccess["cldes2"] != DBNull.Value ? Convert.ToDecimal(draccess["cldes2"]) : 0;
                    ClienteRestringido = ClienteDescuento1 == 0 && ClienteDescuento2 == 0;
                    ClienteAutorizado = !ClienteRestringido;
                    ClienteMensajeValidacion = ClienteRestringido
                        ? "Notifique a la Distribuidora que no tiene descuentos aplicados"
                        : string.Empty;

                    ExisteCuit_Cliente = true;
                }
                else
                {
                    US_CLRAZO = string.Empty;
                    US_CLCIVA = 0;
                    US_CLVEND = 0;
                    US_CLCPAG = 0;
                    ExisteCuit_Cliente = false;

                    ClienteDescuento1 = 0;
                    ClienteDescuento2 = 0;
                    ClienteAutorizado = false;
                    ClienteRestringido = false;
                    ClienteMensajeValidacion = "Datos inválidos";
                }

                draccess.Close();
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();

                return ExisteCuit_Cliente;
            }
            catch (Exception)
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                return false;
            }
            finally
            {
                if (draccess != null && !draccess.IsClosed)
                    draccess.Close();
                cmaccess.Dispose();
            }
        }

        public static void ActualizarFAMIVA()
        {
            var cmFamilias = new OleDbCommand();
            OleDbDataReader drFamilias = null;
            try
            {
                if (cnn.State == ConnectionState.Closed) ConectarDBArti();
                cmFamilias.Connection = cnn;
                cmFamilias.CommandType = CommandType.Text;
                cmFamilias.CommandText = "SELECT DISTINCT FACODI, FADESC, FAMIVA from FAMILIAS order by FACODI";
                drFamilias = cmFamilias.ExecuteReader();
                if (drFamilias.HasRows)
                {
                    ModuloPreparac.EjecutarQueryBDUsu("DELETE FROM FAMILIAS");
                    while (drFamilias.Read())
                    {
                        ModuloPreparac.EjecutarQueryBDUsu($"INSERT INTO FAMILIAS(FACODI, FADESC, FAMIVA) VALUES({drFamilias["FACODI"]},'{drFamilias["FADESC"]}',{drFamilias["FAMIVA"]})");
                    }
                }
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception($", dentro de la función 'ActualizarFAMIVA': {ex.Message}");
            }
            finally
            {
                if (drFamilias != null) drFamilias.Close();
                if (cnn.State == ConnectionState.Open) cnn.Close();
                cmFamilias.Dispose();
            }
        }

        public static void aplicar_desc_Cliente()
        {
            var cm = new OleDbCommand();
            OleDbDataReader recCliente = null;
            OleDbConnection cnnLocal = null;
            try
            {
                cnnLocal = new OleDbConnection($"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={ini_ruta_usu};");
                cnnLocal.Open();
                cm.Connection = cnnLocal;
                cm.CommandType = CommandType.Text;
                cm.CommandText = "select US_DESCUENTO1,US_DESCUENTO2,US_BENEFICIO from USUARIO";
                Application.DoEvents();
                recCliente = cm.ExecuteReader();
                Application.DoEvents();
                if (recCliente.Read())
                {
                        _frmPrincipal.txtDescuento1.Text = recCliente["US_DESCUENTO1"].ToString();
                        _frmPrincipal.txtDescuento2.Text = recCliente["US_DESCUENTO2"].ToString();
                        //_frmPrincipal.txtBeneficio.Text = recCliente["US_BENEFICIO"].ToString();
                }
                else
                {
                        _frmPrincipal.txtDescuento1.Text = "";
                        _frmPrincipal.txtDescuento2.Text = "";
                        //_frmPrincipal.txtBeneficio.Text = "";
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en la Funcion 'aplicar_desc_Cliente': {ex.Message}");
            }
            finally
            {
                if (recCliente != null) recCliente.Close();
                if (cnnLocal != null && cnnLocal.State == ConnectionState.Open) cnnLocal.Close();
                cnnLocal?.Dispose();
                cm.Dispose();
            }
        }

        public static void Funcion_Ofertas()
        {
            var da = new OleDbDataAdapter();
            try
            {
                ConectarDBArti();
                da = new OleDbDataAdapter("SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    //"l.lidesc as Linea," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    //" a.ARMARC as Marca2," +
                    " a.arpre5 as Preciopublico, a.artt as aroo" +
                    ", a.arff as Stock, a.artt, a.arubic1 from listas as a " +
                    "INNER JOIN proveedo AS p ON p.prcodi = a.arprov " +
                    " WHERE a.ARTT <> 0 and a.arprim<>'000' order by a.arterc", cnn);
                var ds = new DataSet();
                da.Fill(ds);
                _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                _frmPrincipal.DataGridView1.Refresh();
                SeleccionarGrilla();
                if (cnn.State == ConnectionState.Open) cnn.Close();
                //Definir_Grilla();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception("Error en la Funcion 'Funcion_Ofertas' " + ex.Message);
                da.Dispose();
            }
        }

        public static void LimpiarCampos(int p1)
        {
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");

                campo = string.Empty;
                valor = string.Empty;
                prcodi = string.Empty;
                _frmPrincipal.txtDescripcion.Text = string.Empty;
                _frmPrincipal.txtModelo.Text = string.Empty;
                _frmPrincipal.txtEquivalencia.Text = string.Empty;
                _frmPrincipal.txtCodigo.Text = string.Empty;
                if (p1 != 1) _frmPrincipal.cmbRubros.SelectedIndex = -1;
                if (p1 != 1) _frmPrincipal.cmbLineas.SelectedIndex = -1;
                // lblCodigoSelec.Text = string.Empty;
                if (p1 != 1)
                    consultas_Relacionales();
                Despintar_Ofertas();
            }
            catch (Exception ex)
            {
                GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Limpiar Campos': {ex.Message}");
            }
        }

        public static void CargarComboLineas(string _Armadora)
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");

                Application.DoEvents();
                DataTable Dt_lineas = new DataTable();
                Dt_lineas.Columns.Add("licodi", typeof(string));
                Dt_lineas.Columns.Add("lidesc", typeof(string));

                // Agregar TODOS como primera opción
                Dt_lineas.Rows.Add("0", "TODOS");

                // Construir SQL base y agregar filtro opcional por armadora
                bool filtrarPorArmadora = !string.IsNullOrEmpty(_Armadora) && _Armadora != "0";

                string sql =
                    "SELECT DISTINCT LI.licodi, LI.lidesc " +
                    // En Jet/Access, al encadenar INNER JOIN conviene usar paréntesis
                    "FROM (lineas AS LI INNER JOIN listas AS LT ON LI.lidesc = LT.ARMODE) ";

                if (filtrarPorArmadora)
                {
                    sql +=
                        "INNER JOIN RUBROS AS r ON LT.ARRUB2 = r.RBCODI " +
                        "WHERE r.RBDESC = ? ";
                }

                sql += "ORDER BY LI.lidesc";

                da = new OleDbDataAdapter(sql, cnn);

                if (filtrarPorArmadora)
                {
                    da.SelectCommand.Parameters.AddWithValue("@Rubro", _Armadora);
                }

                DataTable Dt_temp = new DataTable();
                da.Fill(Dt_temp);

                // Copiar los datos de la consulta
                foreach (DataRow row in Dt_temp.Rows)
                {
                    Dt_lineas.Rows.Add(row["licodi"], row["lidesc"]);
                }

                _frmPrincipal.cmbLineas.DataSource = null;
                _frmPrincipal.cmbLineas.DataSource = Dt_lineas;
                _frmPrincipal.cmbLineas.DisplayMember = "lidesc";
                _frmPrincipal.cmbLineas.ValueMember = "licodi";

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
            catch (Exception ex)
            {
                if (_frmPrincipal != null)
                {
                    //_frmPrincipal.cmbRubros.Visible = false;
                    //_frmPrincipal.cmbLineas.Visible = false;
                }

                if (cnn.State == ConnectionState.Open)
                    cnn.Close();

                throw new Exception($", dentro de 'CargarComboLineas': {ex}");
            }
            finally
            {
                da.Dispose();
                _frmPrincipal.cmbLineas.SelectedIndex = -1;
            }
        }

        public static void EliminarItemCarro(int item)
        {
            try
            {
                ConectarDBArti();
                string sql = "DELETE FROM preparac WHERE prnume=0 AND prusuario=@usuario AND item=@item";
                using var cmd = new OleDbCommand(sql, cnn);
                cmd.Parameters.AddWithValue("@usuario", ini_usuario);
                cmd.Parameters.AddWithValue("@item", item);
                cmd.ExecuteNonQuery();

                if (cnn.State == ConnectionState.Open)
                {
                    cnn.Close();
                }
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                {
                    cnn.Close();
                }
                throw new Exception($"Error al eliminar item del carrito: {ex.Message}");
            }
        }

        public static bool EliminarTodosItemsCarro()
        {
            try
            {
                ConectarDBArti();
                string sql = "DELETE FROM preparac WHERE prnume=0 AND prusuario=@usuario";
                using var cmd = new OleDbCommand(sql, cnn);
                cmd.Parameters.AddWithValue("@usuario", ini_usuario);
                return cmd.ExecuteNonQuery() > 0;
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                {
                    cnn.Close();
                }
                throw new Exception($"Error al eliminar todos los items del carrito: {ex.Message}");
            }
            finally
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
        }

        public static bool ExisteArticuloEnPedido(string ppacod3, string ppaflia, string ppaprov)
        {
            using var cmUsu = new OleDbCommand();
            bool existe = false;
            try
            {
                ConectarDBArti();
                cmUsu.Connection = cnn;
                cmUsu.CommandText = $"SELECT COUNT(*) FROM PREPARAC WHERE ARTERC='{ppacod3}' AND ARFLIA={ppaflia} AND ARPROV={ppaprov}";
                int count = Convert.ToInt32(cmUsu.ExecuteScalar());
                existe = count > 0;
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
                throw new Exception($", dentro de 'ExisteArticuloEnPedido': {ex.Message}");
            }
            finally
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
            return existe;
        }

        public static async Task PrepararArticulo(string accion)
        {
            try
            {
                string sql = $"SELECT * from listas a {ModuloPreparac.WHERE_PREPARAC}";
                
                using (var da = new OleDbDataAdapter(sql, cnnU))
                {
                    var ds = new DataSet();
                    da.Fill(ds);
                }

                if (ds.Tables[0].Rows.Count > 0)
                {
                    await ModuloPreparac.GrabarArticuloaPreparacAsync(ds, accion);
                }
                else
                {
                    MessageBox.Show("No se encontró el artículo", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($", dentro de 'PrepararArticulo': {ex.Message}");
            }
        }

        public static void PrepararArticulo(string codigo, string familia, string proveedor, decimal cantidad)
        {
            try
            {
                ModuloPreparac.WHERE_PREPARAC = $" Where a.ARTERC='{codigo}' AND a.ARPROV={proveedor} AND a.ARFLIA={familia}";
                ModuloPreparac.CANT_PREPARAC = cantidad;
            }
            catch (Exception ex)
            {
                GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Preparar Artículo': {ex.Message}");
            }
        }

        public static void BuscarArticuloParaPreparac(string accion)
        {
            try
            {
                string sql = "";
                if (accion == "INSERT")
                {
                    sql = "INSERT INTO PREPARAC (ARTERC, ARFLIA, ARPROV, CANTIDAD) " +
                          "SELECT a.ARTERC, a.ARFLIA, a.ARPROV, " + ModuloPreparac.CANT_PREPARAC.ToString() +
                          " from listas a " + ModuloPreparac.WHERE_PREPARAC;
                }
                else if (accion == "UPDATE")
                {
                    sql = "UPDATE PREPARAC SET CANTIDAD = " + ModuloPreparac.CANT_PREPARAC.ToString() +
                          " from listas a " + ModuloPreparac.WHERE_PREPARAC;
                }

                using (OleDbCommand cmd = new OleDbCommand(sql, cnn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Buscar Artículo Para Preparac': {ex.Message}");
            }
        }

        //public static bool DescargaArchivo(string nombreArchivo, string rutaDestino, string usuario, string password, string rutaFtp)
        //{
        //    try
        //    {
        //        string ftpPath = $"ftp://{rutaFtp}/{nombreArchivo}";
        //        FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpPath);
        //        request.Credentials = new NetworkCredential(usuario, password);
        //        request.Method = WebRequestMethods.Ftp.DownloadFile;
        //        using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
        //        using (Stream ftpStream = response.GetResponseStream())
        //        using (FileStream fileStream = File.Create(rutaDestino))
        //        {
        //            byte[] buffer = new byte[3071];
        //            int read;
        //            while ((read = ftpStream.Read(buffer, 0, buffer.Length)) > 0)
        //            {
        //                fileStream.Write(buffer, 0, read);
        //            }
        //        }
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Descarga Archivo': {ex.Message}");
        //        return false;
        //    }
        //}
        public static bool DescargaArchivo(string Filename, string NewFileName, string UserNameFtp, string PassFTP, string URI)
        {
            Application.DoEvents();
            FtpWebRequest Ftpconection = (FtpWebRequest)FtpWebRequest.Create(URI + "/" + Filename);
            // Creamos la nueva conexion del tipo web request donde URI es la direccion del servidor
            Ftpconection.Credentials = new NetworkCredential(UserNameFtp, PassFTP);
            //Asignamos las credenciales para el servidor
            Application.DoEvents();
            Ftpconection.Method = WebRequestMethods.Ftp.DownloadFile;
            //El metodo que utilizaremos como peticion, en este cado el de descargar
            try
            {
                System.Net.FtpWebResponse response = (System.Net.FtpWebResponse)Ftpconection.GetResponse();
                Stream responseStr = response.GetResponseStream();
                //guardamos el stream de datos
                FileStream fstr = new FileStream(NewFileName, FileMode.Create);
                byte[] buffer = new byte[3071]; //antes=2047 'Buffer de transferencia
                int ReadStream = 0;
                do
                {
                    Application.DoEvents();
                    ReadStream = responseStr.Read(buffer, 0, buffer.Length); //leemos el stream de los datos
                    fstr.Write(buffer, 0, ReadStream); //escribimos en el archivo
                } while (ReadStream != 0);
                Application.DoEvents();
                responseStr.Close(); //cerramos los archivos y las conexiones
                fstr.Flush();
                fstr.Close();
                responseStr.Close();
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'DescargaArchivo': {ex.Message}");
                return false;
            }
        }
        public static string TraerFechaActualizacion()
        {
            try
            {
                using (OleDbCommand cmd = new OleDbCommand("SELECT MAX(FECHA) FROM ACTUALIZACIONES", cnn))
                {
                    object result = cmd.ExecuteScalar();
                    return result != DBNull.Value ? Convert.ToDateTime(result).ToString("dd/MM/yyyy") : DateTime.Now.ToString("dd/MM/yyyy");
                }
            }
            catch (Exception ex)
            {
                GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Traer Fecha Actualización': {ex.Message}");
                return DateTime.Now.ToString("dd/MM/yyyy");
            }
        }

        public static void ValidarEquivalencias()
        {
            try
            {
                int numFila = _frmPrincipal.DataGridView1.CurrentCellAddress.Y;
                //_frmPrincipal.DataGridView1.Columns[7].Visible = true;

                if (!string.IsNullOrEmpty(_frmPrincipal.DataGridView1["AROBSE", numFila].Value?.ToString()))
                {
                    _frmPrincipal.btnEquivalencia.Enabled = true;
                    NumEqui = _frmPrincipal.DataGridView1["AROBSE", numFila].Value.ToString();
                }
                else
                {
                    _frmPrincipal.btnEquivalencia.Enabled = false;
                    NumEqui = string.Empty;
                }

                //_frmPrincipal.DataGridView1.Columns[7].Visible = false;
            }
            catch (Exception ex)
            {
                //_frmPrincipal.DataGridView1.Columns[7].Visible = false;
                NumEqui = string.Empty;
                throw new Exception($", dentro de 'ValidarEquivalencias': {ex.Message}");
            }
        }


        public static void Update_PrecioCosto()
        {
            using (OleDbCommand cmaccessUpdate = new OleDbCommand())
            {
                try
                {
                    ConectarDBArti();
                    cmaccessUpdate.Connection = cnn;
                    cmaccessUpdate.CommandType = CommandType.Text;

                    cmaccessUpdate.CommandText = "update listas set arpre1=aroo";
                    Application.DoEvents();
                    cmaccessUpdate.ExecuteNonQuery();

                    // Aplico los descuentos
                    if (decimal.TryParse(_frmPrincipal.txtDescuento1.Text, out decimal descuento1) && descuento1 != 0)
                    {
                        cmaccessUpdate.CommandText = "update listas as up " +
                                                   "set up.arpre0=format(up.aroo-(up.aroo*" + _frmPrincipal.txtDescuento1.Text + "/100),'Fixed')";
                        cmaccessUpdate.ExecuteNonQuery();

                        if (decimal.TryParse(_frmPrincipal.txtDescuento2.Text, out decimal descuento2) && descuento2 != 0)
                        {
                            cmaccessUpdate.CommandText = "update listas as up " +
                                                       "set up.arpre1=format(up.arpre0-(up.arpre0*" + _frmPrincipal.txtDescuento2.Text + "/100),'Fixed')";
                        }
                        else
                        {
                            cmaccessUpdate.CommandText = "update listas set arpre1=arpre0";
                        }
                        cmaccessUpdate.ExecuteNonQuery();
                    }

                    //using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FADESC FROM [familias] F INNER JOIN [listas] L ON L.ARFLIA = F.FACODI", cnn))
                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FADESC FROM [familias] F ", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        ConectarDBUsuario();
                        using var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA, FADES1 FROM FAMILIAS WHERE FADESC = ?", cnnU);
                        cmUsu.Parameters.Add("@p1", OleDbType.VarWChar);

                        while (drFam.Read())
                        {
                            string fad = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(fad))
                                continue;

                            // Releemos descuentos para decidir dónde aplicar IVA
                            decimal d1 = 0, d2 = 0;
                            decimal.TryParse(_frmPrincipal.txtDescuento1.Text, out d1);
                            decimal.TryParse(_frmPrincipal.txtDescuento2.Text, out d2);

                            decimal? faiva = null;
                            decimal? fades1 = null;
                            cmUsu.Parameters[0].Value = fad;
                            using (var rdr = cmUsu.ExecuteReader())
                            {
                                if (rdr.Read())
                                {
                                    if (rdr[0] != DBNull.Value && decimal.TryParse(rdr[0].ToString(), out var ivaTmp))
                                        faiva = ivaTmp;
                                    if (rdr[1] != DBNull.Value && decimal.TryParse(rdr[1].ToString(), out var desTmp))
                                        fades1 = desTmp;
                                }
                            }

                            if (fades1.HasValue && fades1.Value != 0)
                            {
                                string des = fades1.Value.ToString().Replace(",", ".");
                                if (d2 != 0)
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 - (L.arpre1*" + des + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre0 = FORMAT(L.arpre0 - (L.arpre0*" + des + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = L.arpre0 WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 - (L.arpre1*" + des + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                            }

                            if (faiva.HasValue)
                            {
                                string iva = faiva.Value.ToString().Replace(",", ".");
                                // Si hay desc2, el final está en arpre1 -> IVA sobre arpre1
                                if (d2 != 0)
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 + (L.arpre1*" + iva + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    // Solo desc1: IVA sobre arpre0 y luego reflejar en arpre1
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre0 = FORMAT(L.arpre0 + (L.arpre0*" + iva + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = L.arpre0 WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    // Sin descuentos: IVA directo sobre arpre1
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 + (L.arpre1*" + iva + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                //si no encuentra la familia o el iva esta en 0?
                                //MessageBox.Show(faiva.HasValue.ToString());
                            }
                        }

                        if (cnnU.State == ConnectionState.Open)
                            cnnU.Close();
                    }

                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                }
                catch (Exception ex)
                {
                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();

                    throw new Exception(", dentro de 'Update_PrecioCosto': " +
                                      (ex.InnerException?.ToString() ?? "") + " - " + ex.Message);
                }
            }
        }

        public static void Update_PrecioCosto_Ofertas()
        {
            using (OleDbCommand cmaccessUpdate = new OleDbCommand())
            {
                try
                {
                    ConectarDBArti();
                    cmaccessUpdate.Connection = cnn;
                    cmaccessUpdate.CommandType = CommandType.Text;

                    cmaccessUpdate.CommandText = "update listas set arpre4=artt";
                    Application.DoEvents();
                    cmaccessUpdate.ExecuteNonQuery();

                    // Aplico los descuentos
                    if (decimal.TryParse(_frmPrincipal.txtDescuento1.Text, out decimal descuento1) && descuento1 != 0)
                    {
                        cmaccessUpdate.CommandText = "update listas as up " +
                                                   "set up.arpre3=format(up.artt-(up.artt*" + _frmPrincipal.txtDescuento1.Text + "/100),'Fixed')";
                        cmaccessUpdate.ExecuteNonQuery();

                        if (decimal.TryParse(_frmPrincipal.txtDescuento2.Text, out decimal descuento2) && descuento2 != 0)
                        {
                            cmaccessUpdate.CommandText = "update listas as up " +
                                                       "set up.arpre4=format(up.arpre3-(up.arpre3*" + _frmPrincipal.txtDescuento2.Text + "/100),'Fixed')";
                        }
                        else
                        {
                            cmaccessUpdate.CommandText = "update listas set arpre4=arpre3";
                        }
                        cmaccessUpdate.ExecuteNonQuery();
                    }

                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FADESC FROM [familias] F INNER JOIN [listas] L ON L.ARFLIA = F.FACODI", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        ConectarDBUsuario();
                        using var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA, FADES1 FROM FAMILIAS WHERE FADESC = ?", cnnU);
                        cmUsu.Parameters.Add("@p1", OleDbType.VarWChar);

                        while (drFam.Read())
                        {
                            string fad = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(fad))
                                continue;

                            decimal d1 = 0, d2 = 0;
                            decimal.TryParse(_frmPrincipal.txtDescuento1.Text, out d1);
                            decimal.TryParse(_frmPrincipal.txtDescuento2.Text, out d2);

                            decimal? faiva = null;
                            decimal? fades1 = null;
                            cmUsu.Parameters[0].Value = fad;
                            using (var rdr = cmUsu.ExecuteReader())
                            {
                                if (rdr.Read())
                                {
                                    if (rdr[0] != DBNull.Value && decimal.TryParse(rdr[0].ToString(), out var ivaTmp))
                                        faiva = ivaTmp;
                                    if (rdr[1] != DBNull.Value && decimal.TryParse(rdr[1].ToString(), out var desTmp))
                                        fades1 = desTmp;
                                }
                            }

                            if (fades1.HasValue && fades1.Value != 0)
                            {
                                string des = fades1.Value.ToString().Replace(",", ".");
                                if (d2 != 0)
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 - (L.arpre4*" + des + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre3 = FORMAT(L.arpre3 - (L.arpre3*" + des + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = L.arpre3 WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 - (L.arpre4*" + des + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                            }

                            if (faiva.HasValue)
                            {
                                string iva = faiva.Value.ToString().Replace(",", ".");
                                if (d2 != 0)
                                {
                                    // Con desc2: final en arpre4
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 + (L.arpre4*" + iva + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    // Solo desc1: IVA sobre arpre3 y luego reflejar en arpre4
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre3 = FORMAT(L.arpre3 + (L.arpre3*" + iva + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = L.arpre3 WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    // Sin descuentos: IVA directo sobre arpre4
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 + (L.arpre4*" + iva + "/100),'Fixed') WHERE F.FADESC = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", fad);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                            }
                        }

                        if (cnnU.State == ConnectionState.Open)
                            cnnU.Close();
                    }

                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                }
                catch (Exception ex)
                {
                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();

                    throw new Exception(", dentro de 'Update_PrecioCosto_Ofertas': " +
                                      (ex.InnerException?.ToString() ?? "") + " - " + ex.Message);
                }
            }
        }

        public static void Update_PrecioSug()
        {
            using (OleDbCommand cmaccess = new OleDbCommand())
            {
                try
                {
                    ConectarDBArti();
                    cmaccess.Connection = cnn;
                    cmaccess.CommandType = CommandType.Text;

                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FADESC FROM [familias] F INNER JOIN [listas] L ON L.ARFLIA = F.FACODI", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        while (drFam.Read())
                        {
                            string fad = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(fad))
                                continue;

                            decimal? faben1 = null;
                            using (var cmUsu = new OleDbCommand("SELECT TOP 1 FABEN1 FROM FAMILIAS WHERE FADESC = ?", cnnU))
                            {
                                ConectarDBUsuario();
                                cmUsu.Parameters.AddWithValue("@p1", fad);
                                var obj = cmUsu.ExecuteScalar();
                                if (obj != null && obj != DBNull.Value)
                                {
                                    if (decimal.TryParse(obj.ToString(), out var benTmp))
                                        faben1 = benTmp;
                                }
                                if (cnnU.State == ConnectionState.Open)
                                    cnnU.Close();
                            }

                            if (faben1.HasValue && faben1.Value != 0)
                            {
                                string ben = faben1.Value.ToString().Replace(",", ".");
                                cmaccess.Parameters.Clear();
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre2 = FORMAT(L.arpre1 + (L.arpre1*" + ben + "/100),'Fixed') WHERE F.FADESC = ?";
                                cmaccess.Parameters.AddWithValue("@p1", fad);
                                cmaccess.ExecuteNonQuery();
                            }
                            else
                            {
                                cmaccess.Parameters.Clear();
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre2 = FORMAT(L.arpre1,'Fixed') WHERE F.FADESC = ?";
                                cmaccess.Parameters.AddWithValue("@p1", fad);
                                try { GrabarError($"Familia sin FABEN1 (o 0) al sugerir precio: '{fad}'"); } catch {}
                                cmaccess.ExecuteNonQuery();
                            }
                        }
                    }

                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                }
                catch (Exception ex)
                {
                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                    GrabarError($", dentro de 'Update_PrecioSug': " +
                  (ex.InnerException?.ToString() ?? "") + " - " + ex.Message);

                    throw new Exception(", dentro de 'Update_PrecioSug': " +
                                      (ex.InnerException?.ToString() ?? "") + " - " + ex.Message);
                }
            }
        }

        public static void Update_PrecioSug_Ofertas()
        {
            using (OleDbCommand cmaccess = new OleDbCommand())
            {
                try
                {
                    ConectarDBArti();
                    cmaccess.Connection = cnn;
                    cmaccess.CommandType = CommandType.Text;

                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FADESC FROM [familias] F INNER JOIN [listas] L ON L.ARFLIA = F.FACODI", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        while (drFam.Read())
                        {
                            string fad = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(fad))
                                continue;

                            decimal? faben1 = null;
                            using (var cmUsu = new OleDbCommand("SELECT TOP 1 FABEN1 FROM FAMILIAS WHERE FADESC = ?", cnnU))
                            {
                                ConectarDBUsuario();
                                cmUsu.Parameters.AddWithValue("@p1", fad);
                                var obj = cmUsu.ExecuteScalar();
                                if (obj != null && obj != DBNull.Value)
                                {
                                    if (decimal.TryParse(obj.ToString(), out var benTmp))
                                        faben1 = benTmp;
                                }
                                if (cnnU.State == ConnectionState.Open)
                                    cnnU.Close();
                            }

                            if (faben1.HasValue && faben1.Value != 0)
                            {
                                string ben = faben1.Value.ToString().Replace(",", ".");
                                cmaccess.Parameters.Clear();
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre5 = FORMAT(L.arpre4 + (L.arpre4*" + ben + "/100),'Fixed') WHERE F.FADESC = ?";
                                cmaccess.Parameters.AddWithValue("@p1", fad);
                                cmaccess.ExecuteNonQuery();
                            }
                            else
                            {
                                cmaccess.Parameters.Clear();
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre5 = FORMAT(L.arpre4,'Fixed') WHERE F.FADESC = ?";
                                cmaccess.Parameters.AddWithValue("@p1", fad);
                                try { GrabarError($"Familia sin FABEN1 (o 0) al sugerir precio ofertas: '{fad}'"); } catch {}
                                cmaccess.ExecuteNonQuery();
                            }
                        }
                    }

                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                }
                catch (Exception ex)
                {
                    if (cnn.State == ConnectionState.Open)
                        cnn.Close();
                    GrabarError($", dentro de 'Update_PrecioSug_Ofertas': " +
                                      (ex.InnerException?.ToString() ?? "") + " - " + ex.Message);
                    throw new Exception(", dentro de 'Update_PrecioSug_Ofertas': " +
                                      (ex.InnerException?.ToString() ?? "") + " - " + ex.Message);
                }
            }
        }

        public static void CargarDatosExtras(string StringSelect)
        {
            try
            {
                ConectarDBArti();
                using var cmd = new OleDbCommand(StringSelect, cnn);
                using var dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    _frmPrincipal.lblRubro.Text = "Rubro: " + dr["RBDESC"].ToString();
                    _frmPrincipal.lblOrigen.Text = "Origen: "+ dr["ARORIG"].ToString();
                    _frmPrincipal.lblObservaciones.Text = "Observación: " + dr["ARUBIC1"].ToString() + 
                        Environment.NewLine + "Ubicación: " + dr["AROBSE"].ToString(); //AROBSE
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en CargarDatosExtras: {ex.Message}");
            }
            finally
            {
                if (cnnU.State == ConnectionState.Open)
                    cnnU.Close();
            }
        }

        //public static void MostrarAlertaUpdate(string message)
        //{
        //    try
        //    {
        //        if (_frmPrincipal == null) { GrabarError(" - MostrarAlertaUpdate: _frmPrincipal es null"); return; }
        //        var f = _frmPrincipal;
        //        void DoUI()
        //        {
        //            try
        //            {
        //                var item = f.lblAlertaUpdate;
        //                if (item == null) { GrabarError(" - MostrarAlertaUpdate: lblAlertaUpdate es null"); return; }
        //                item.Text = message ?? string.Empty;
        //                item.Visible = true;
        //                var parent = item.GetCurrentParent();
        //                if (parent == null)
        //                {
        //                    var ss = f.Controls.OfType<System.Windows.Forms.StatusStrip>().FirstOrDefault();
        //                    if (ss != null)
        //                    {
        //                        ss.Items.Add(item);
        //                        parent = ss;
        //                    }
        //                }
        //                if (parent != null)
        //                {
        //                    parent.Visible = true;
        //                    parent.PerformLayout();
        //                    parent.Invalidate();
        //                    parent.Update();
        //                    parent.Refresh();
        //                    var bg = parent.BackColor;
        //                    float b = bg.GetBrightness();
        //                    item.ForeColor = (b < 0.5f) ? System.Drawing.Color.Yellow : System.Drawing.Color.Red;
        //                    item.BackColor = System.Drawing.Color.Transparent;
        //                }
        //                if (!item.Visible) GrabarError(" - MostrarAlertaUpdate: lblAlertaUpdate.Visible es false");
        //                if (parent == null) GrabarError(" - MostrarAlertaUpdate: GetCurrentParent() es null");
        //                else if (!parent.Visible) GrabarError(" - MostrarAlertaUpdate: parent.Visible es false");
        //                if (string.IsNullOrWhiteSpace(item.Text)) GrabarError(" - MostrarAlertaUpdate: lblAlertaUpdate.Text vacío");
        //            }
        //            catch (Exception exUI)
        //            {
        //                GrabarError($" - MostrarAlertaUpdate/UI: {exUI.Message}");
        //            }
        //        }
        //        if (f.InvokeRequired) f.Invoke((Action)DoUI); else DoUI();
        //    }
        //    catch (Exception ex)
        //    {
        //        GrabarError($" - MostrarAlertaUpdate: {ex.Message}");
        //    }
        //}
    }
}

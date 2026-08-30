using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
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
    /// <summary>
    /// Módulo estático central de la aplicación. Coordina todas las áreas funcionales:
    /// base de datos, configuración, cálculo de precios, búsquedas, carrito, FTP, PDF y correo.
    /// </summary>
    /// <remarks>
    /// <b>Bases de datos:</b>
    /// <list type="bullet">
    /// <item><c>cnn</c>: base del distribuidor (artículos, clientes, familias) → ruta <c>ini_ruta_base</c>.</item>
    /// <item><c>cnnU</c>: base local del usuario (<c>cliente.dat</c>) → descuentos, PREPARAC, configuración.</item>
    /// <item><c>cnnO</c>: base de origen temporal para el proceso de actualización FTP.</item>
    /// </list>
    /// <b>Cadena de precios en tabla LISTAS:</b>
    /// <list type="bullet">
    /// <item><c>aroo</c> = precio base de lista | <c>artt</c> = precio de oferta.</item>
    /// <item><c>arpre0</c> = aroo − desc1 &nbsp;|&nbsp; <c>arpre1</c> = arpre0 − desc2 − descFamilia + IVA (precio de costo).</item>
    /// <item><c>arpre2</c> = arpre1 + beneficio familiar = precio sugerido de venta al público.</item>
    /// <item><c>arpre3/arpre4/arpre5</c>: equivalentes para artículos en oferta (<c>artt</c>).</item>
    /// </list>
    /// <b>Flujo de inicio:</b> <see cref="Inicializacion"/> → <see cref="LeerConfig"/>
    /// → <see cref="Cargar_Usuario"/> → <see cref="funcion_inicial"/> → <see cref="aplicar_desc_Cliente"/>.<br/>
    /// <b>Actualización FTP:</b> <see cref="ActualizacionXFTP"/> descarga archivo, migra base y recalcula precios.<br/>
    /// <b>Pedido:</b> <see cref="ModuloPreparac"/> gestiona PREPARAC; <see cref="ProcesoCorreo"/> genera PDF y envía por email.
    /// </remarks>
    public static class Module
    {
        #region Variables Globales
        // --- Referencias a formularios (inyectadas por Set*Form; uso interno) ---
        // _frmPrincipal:    ventana principal de la app (grilla, precios, filtros).
        // _frmCarroCompras: formulario del carrito de compras (gridPedidos).
        // _frmDatosUsuario: formulario de datos del cliente (código, CUIT, descuentos).
        // _carroComprasForm/_datosUsuarioForm: instancias singleton lazy (ver Get*Form).
        private static frmPrincipal _frmPrincipal;
        private static frmCarroCompras _frmCarroCompras;
        private static frmDatosUsuario _frmDatosUsuario;
        private static frmCarroCompras _carroComprasForm;
        private static frmDatosUsuario _datosUsuarioForm;

        /// <summary>
        /// DataTable maestro resultado de la última consulta SQL (filtros superiores).
        /// Los filtros inferiores (combos) aplican DataView.RowFilter sobre esta tabla sin volver a consultar la BD.
        /// </summary>
        public static DataTable _dtMaestro;

        /// <summary>Registra la referencia al formulario principal. Debe llamarse en el constructor de <c>frmPrincipal</c>.</summary>
        /// <param name="form">Instancia de <c>frmPrincipal</c>; no puede ser null.</param>
        public static void SetMainForm(frmPrincipal form)
        {
            _frmPrincipal = form ?? throw new ArgumentNullException(nameof(form));
        }

        /// <summary>Registra la referencia al formulario del carrito. Debe llamarse al abrir <c>frmCarroCompras</c>.</summary>
        /// <param name="form">Instancia de <c>frmCarroCompras</c>; no puede ser null.</param>
        public static void SetCarroComprasForm(frmCarroCompras form)
        {
            _frmCarroCompras = form ?? throw new ArgumentNullException(nameof(form));
        }

        /// <summary>Registra la referencia al formulario de datos de usuario y propaga la referencia a <see cref="ModuloUsuario"/>.</summary>
        /// <param name="form">Instancia de <c>frmDatosUsuario</c>; no puede ser null.</param>
        public static void SetDatosUsuarioForm(frmDatosUsuario form)
        {
            if (form == null) throw new ArgumentNullException(nameof(form));
            _frmDatosUsuario = form;
            ModuloUsuario.SetDatosUsuarioForm(form); // Actualizar también la referencia en ModuloUsuario
        }

        /// <summary>
        /// Retorna la instancia singleton de <c>frmCarroCompras</c>, creándola si no existe o fue cerrada.
        /// </summary>
        /// <returns>Instancia activa (no dispuesta) de <c>frmCarroCompras</c>.</returns>
        public static frmCarroCompras GetCarroComprasForm()
        {
            if (_carroComprasForm == null || _carroComprasForm.IsDisposed)
            {
                _carroComprasForm = new frmCarroCompras();
            }
            return _carroComprasForm;
        }

        /// <summary>
        /// Carga en la grilla los artículos marcados como novedades (<c>ARFECASIG &lt;&gt; 0</c>),
        /// es decir, aquellos con una fecha de asignación reciente en la base del distribuidor.
        /// </summary>
        public static async Task Funcion_Novedades()
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
                    "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                    //"l.lidesc as Linea," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    //" a.ARMARC as Marca2," + 
                    "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from (listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) " +
                    "LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI " +
                    "WHERE a.ARFECASIG <> 0 and a.arprim<>'000' order by a.arterc", cnn);

                var ds = new DataSet();
                da.Fill(ds);

                // Una vez que obtengo los datos, los cargo en la grilla para hacerlos visible al usuario
                if (ds.Tables[0].Rows.Count > 0)
                {
                    _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                    _frmPrincipal.DataGridView1.Refresh();
                    await SeleccionarGrilla();
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

        /// <summary>
        /// Retorna la instancia singleton de <c>frmDatosUsuario</c>, creándola si no existe o fue cerrada.
        /// </summary>
        /// <returns>Instancia activa (no dispuesta) de <c>frmDatosUsuario</c>.</returns>
        public static frmDatosUsuario GetDatosUsuarioForm()
        {
            if (_datosUsuarioForm == null || _datosUsuarioForm.IsDisposed)
            {
                _datosUsuarioForm = new frmDatosUsuario();
            }
            return _datosUsuarioForm;
        }

        // --- Configuración de sesión ---
        // ini_usuario: nombre/código del usuario que inició sesión.
        // ini_fecha:   fecha de inicio de sesión.
        public static string ini_usuario { get; set; }
        public static string ini_fecha { get; set; }

        // --- Conexiones a bases de datos ---
        // cnn:   base del distribuidor (artículos, clientes, familias) → ini_ruta_base.
        // cnnU:  base local del usuario (cliente.dat) → ini_ruta_usu.
        // cnnO:  base de origen temporal durante el proceso de actualización FTP.
        // cm/dr/ds/dt: objetos ADO.NET reutilizados en métodos que no usan using (uso con cuidado).
        // campo/valor: variables auxiliares para construcción dinámica de WHERE en búsquedas.
        public static OleDbConnection cnn = new OleDbConnection();
        public static OleDbConnection cnnU = new OleDbConnection();
        public static OleDbConnection cnnO = new OleDbConnection();
        public static OleDbCommand cm;
        public static OleDbDataReader dr;
        public static DataSet ds = new DataSet();
        public static DataTable dt;
        public static string campo = "";
        public static string valor = "";
        
        // --- Rutas de bases de datos ---
        // ini_ruta_arti: no usado actualmente (reemplazado por ini_ruta_base).
        // ini_ruta_usu:  ruta completa a cliente.dat (base local del usuario).
        public static string ini_ruta_arti { get; set; }
        public static string ini_ruta_usu { get; set; }

        // --- Cadena de cálculo de precios (actualizados en Calculadora() por cada selección de artículo) ---
        // precio_lista:         aroo del artículo seleccionado (precio base sin descuentos).
        // precio_desc1:         precio_lista − descuento1 (%).
        // precio_desc2:         precio_desc1 − descuento2 (%).
        // precio_desc1_desc2:   alias de precio_desc2 (no usado actualmente).
        // precio_costo_sin_iva: precio_desc2 − descuento familia = precio de costo sin IVA (base para PREPARAC).
        // precio_final:         precio_costo_sin_iva + IVA familia = precio final al cliente con IVA.
        // PRECIO_DESC3:         monto del tercer descuento (­optional) sobre precio_final.
        // precio_final_desc3:   precio_final − desc3 = precio publicado si hay desc3 activo.
        // preciofinalxx:        copia de precio_final antes de aplicar desc3 (intermedio).
        // precio_con_desc3:     alias de precio_final_desc3.
        // precio_sin_desc3:     alias de precio_final cuando desc3 = 0.
        // precio_beneficio:     monto del beneficio familiar sobre el precio (no usado en flujo actual).
        // precio_final_beneficio: precio_final + beneficio familiar (no usado en flujo actual).
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
        // AccionDatosUsu: indica la acción a realizar en frmDatosUsuario (1=alta, 2=edición).
        // NumEqui:        código de equivalencia del artículo seleccionado (AROBSE); vacío si no tiene.
        // vectornomfoto:  array de nombres de archivo de fotos (no usado actualmente).
        // prcodi:         código de proveedor del artículo seleccionado para filtros.
        public static int AccionDatosUsu;
        public static string NumEqui = string.Empty;
        public static string[] vectornomfoto;
        public static string prcodi = "";

        // --- Configuración de Email (leídas de Config.ini) ---
        // ini_emisor_email:  dirección del remitente SMTP.
        // ini_pass_email:    contraseña del remitente.
        // ini_host_email:    servidor SMTP (ej: smtp.gmail.com).
        // ini_port_email:    puerto SMTP como string.
        // ini_enable_ssl:    indica si el SMTP requiere SSL.
        // ini_receptor_email: destinatario por defecto (distribuidora).
        // --- Configuración FTP (leídas de Config.ini) ---
        // ini_ruta_ftp:      URL del servidor FTP (ej: ftp://servidor.com).
        // ini_usu_ftp:       usuario FTP.
        // ini_pass_ftp:      contraseña FTP.
        // ini_arch_ftp:      nombre del archivo de actualización en el FTP.
        // ini_ruta_base:     ruta local a la base de artículos del distribuidor.
        // ini_ruta_base_descarga: ruta temporal donde se descarga la base actualizada.
        // ini_ruta_banner:   URL para el texto del banner superior de la app.
        // ini_ruta_gral:     ruta general de la instalación local.
        // ini_apli_ftp:      ruta del ejecutable instalador descargado del FTP.
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
        public static string ini_ruta_banner_grande;
        public static string ini_ruta_banner_boton;
        public static string ini_link_banner_grande;
        public static string ini_link_banner_boton;
        public static string ini_ruta_gral;
        public static string ini_apli_ftp;

        // --- Configuración FTP para fotos de artículos ---
        public static string ini_ftp_fotos_host;
        public static int ini_ftp_fotos_port = 21;
        public static string ini_ftp_fotos_user;
        public static string ini_ftp_fotos_pass;
        public static string ini_ftp_fotos_folder = "/FOTOCATALOGO";

        // Servicio de caché de fotos FTP (inicializado en LeerConfig)
        public static FtpPhotoCacheService _fotoService;

        public static string _textoBanner;


        // --- Semáforo de stock (leídos de Config.ini, usados en ColorearStock()) ---
        // ini_cond_rojo:     umbral máximo de stock para color ROJO (crítico).
        // ini_cond_amarillo: umbral para color AMARILLO (precaución).
        // ini_cond_verde:    umbral mínimo para color VERDE (stock suficiente).
        public static int ini_cond_rojo;
        public static int ini_cond_amarillo;
        public static int ini_cond_verde;

        // fecha:  almacena la fecha de última actualización de la base (leída de LISTAS.arpp).
        // Migra:  flag temporal que indica si la migración de base estuvo activa (usado en ActualizacionXFTP).
        public static string fecha = "";

        public static bool Migra = false;

        // --- Datos del cliente/usuario activo (cargados en Cargar_Usuario y ValidarCuitCliente) ---
        // US_CODI:   código de cliente en la base del distribuidor.
        // US_CLRAZO: razón social del cliente.
        // US_CUIT:   CUIT del cliente (usado para validación cruzada).
        // US_EMAIL:  dirección de email del cliente (se agrega como destinatario del pedido).
        // US_CLCIVA: condición de IVA del cliente (CLIENTES.CLCIVA).
        // US_CLVEND: código del vendedor asignado (CLIENTES.CLVEND).
        // US_CLCPAG: condición de pago del cliente (CLIENTES.CLCPAG).
        public static string US_CODI = "0";
        public static string US_CLRAZO = "0";
        public static string US_CUIT = "0";
        public static string US_EMAIL = "0";

        public static int US_CLCIVA = 0;
        public static int US_CLVEND = 0;
        public static int US_CLCPAG = 0;

        public static int US_DESCUENTO1 = 0;
        public static int US_DESCUENTO2 = 0;

        // --- Estado del artículo actualmente seleccionado en la grilla ---
        // _ARTI_ARFLIA: código de familia del artículo (LISTAS.ARFLIA); usado en Calculadora, aplicar_beneficio.
        private static int _ARTI_ARFLIA;

        // accionDatosUsu: duplicado de AccionDatosUsu (mantener consistencia al escribir).
        public static int accionDatosUsu;

        // --- Estado de autorización del cliente (actualizado en ValidarCuitCliente) ---
        // ExisteCodi_Cliente:     true si el código existe en CLIENTES de la base del distribuidor.
        // ExisteCuit_Cliente:     true si la combinación código + CUIT es válida.
        // ClienteAutorizado:      true cuando CLDES1 o CLDES2 > 0 (tiene descuentos aplicados).
        // ClienteRestringido:     true cuando ambos descuentos son 0 (sin habilitación del distribuidor).
        // ClienteDescuento1/2:    descuentos de la tabla CLIENTES (CLDES1/CLDES2), cargados en la validación.
        // ClienteMensajeValidacion: mensaje informativo de estado para mostrar al usuario.
        public static bool ExisteCodi_Cliente = false;
        public static bool ExisteCuit_Cliente = false;

        public static bool ClienteAutorizado = false;
        public static bool ClienteRestringido = false;
        public static decimal ClienteDescuento1 = 0;
        public static decimal ClienteDescuento2 = 0;
        public static string ClienteMensajeValidacion = string.Empty;

        // _descargarFichero: flag para controlar si se debe descargar el fichero FTP (uso interno).
        // _textoBanner:      contenido del banner descargado desde la URL configurada en Config.ini.
        public static bool _descargarFichero;

        /// <summary>
        /// Ejecuta un SELECT sobre la base del distribuidor y retorna los campos
        /// <c>ARFLIA</c> (familia) y <c>ARPROV</c> (proveedor) del artículo.
        /// </summary>
        /// <param name="stringSelect">Sentencia SELECT completa con WHERE.</param>
        /// <returns>Array de dos elementos: [0]=ARFLIA, [1]=ARPROV; o <c>["0"]</c> si no hay datos.</returns>
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

        /// <summary>Establece la cláusula WHERE dinámica de búsqueda para el próximo artículo a agregar al carrito.</summary>
        /// <param name="value">Cláusula WHERE completa (ej: <c>" WHERE A.ARCODI='X'..."</c>).</param>
        public static void SetWherePreparac(string value)
        {
            ModuloPreparac.WHERE_PREPARAC = value;
        }

        /// <summary>Retorna la cláusula WHERE actualmente configurada en <see cref="ModuloPreparac"/>.</summary>
        public static string GetWherePreparac()
        {
            return ModuloPreparac.WHERE_PREPARAC;
        }

        /// <summary>Establece la cantidad para la próxima operación de carrito, parseando el string como decimal.</summary>
        /// <param name="value">Cantidad como cadena; se normaliza el separador decimal antes de parsear.</param>
        /// <exception cref="ArgumentException">Si el valor no puede convertirse a decimal.</exception>
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

        /// <summary>Retorna la cantidad actualmente configurada en <see cref="ModuloPreparac"/>.</summary>
        public static decimal GetCantPreparac()
        {
            return ModuloPreparac.CANT_PREPARAC;
        }

        #region Variables Preparac
        // ini_codcli: código del cliente para el proceso de preparación (no usado en flujo principal actual).
        // ini_barr:   código de barrio del cliente (no usado en flujo principal actual).
        // ini_razon:  razón social del cliente para preparación (no usado en flujo principal actual).
        public static string ini_codcli = "";
        public static string ini_barr = "";
        public static string ini_razon = "";
        #endregion

        #endregion

        #region Métodos TXT
        /// <summary>
        /// Inicia el proceso de exportación del catálogo a Excel (.xlsx) mostrando la barra de progreso.
        /// Actualmente llama a <see cref="PasarArti_a_xls"/> en lugar de <see cref="PasarArti_a_TXT"/>.
        /// </summary>
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
                //throw new Exception(", dentro de 'Iniciar_Proceso_TXT': " + ex.Message.ToString());
            }
        }

        /// <summary>
        /// Exporta todos los artículos de LISTAS a un archivo de texto plano posicional
        /// (<c>lista_txt.txt</c>) y lo abre con la aplicación predeterminada del sistema.
        /// </summary>
        /// <remarks>Campos exportados con padding fijo: arterc(10), arcuar(4), precio(10), marc(15), ardesc(80), arobse(80).
        /// Actualmente no se llama (reemplazado por <see cref="PasarArti_a_xls"/>).</remarks>
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
        /// <summary>
        /// Exporta todos los artículos de LISTAS a un archivo Excel formateado
        /// (<c>lista_excel.xlsx</c>) usando la librería ClosedXML y lo abre con Excel.
        /// </summary>
        /// <remarks>Incluye encabezados en negrita con fondo gris, formato moneda en la columna de precio
        /// (<c>$#,##0.00</c>) y ajuste automático del ancho de columnas.</remarks>
        public static void PasarArti_a_xls()
        {
            OleDbDataAdapter da = new OleDbDataAdapter();
            try
            {
                string excelPath = @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\lista_excel.xlsx";

                //// SOLUCION: Cerrar Excel si está abierto con este archivo
                //try
                //{
                //    if (File.Exists(excelPath))
                //    {
                //        // Intentar eliminar el archivo existente
                //        File.Delete(excelPath);
                //    }
                //}
                //catch (Exception)
                //{
                //    // Si no se puede eliminar, usar un nombre alternativo
                //    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                //    excelPath = $@"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\lista_excel_{timestamp}.xlsx";
                //}

                ConectarDBArti();
                
                // TEST 1: Primero probamos el SELECT básico que funcionaba antes
                string testQuery = "SELECT " +
                    "iif(isnull(a.arprim),'',a.arprim) as arprim, " +
                    "iif(isnull(a.arterc),'',a.arterc) as arterc, " +
                    "iif(isnull(a.arcuar),'',a.arcuar) as arcuar, " +
                    "mid(a.armarc,1,15) as armarc, " +
                    "mid(iif(isnull(a.ardesc),'',a.ardesc),1,80) as ardesc, " +
                    "mid(iif(isnull(a.arobse),'',a.arobse),1,80) as arobse, " +
                    "iif(isnull(a.arpre0),0,a.arpre0) as costo_sin_iva, " +
                    "iif(isnull(a.arpre2),0,a.arpre2) as precio_lista " +
                    "FROM listas as a " +
                    "WHERE a.arprim<>'000' " +
                    "ORDER BY a.arcodi";
                
                da = new OleDbDataAdapter(testQuery, cnn);

                DataSet ds = new DataSet();
                da.Fill(ds, "articulo");
                
                // TEST: Verificar si obtuvimos datos
                if (ds.Tables["articulo"].Rows.Count == 0)
                {
                    throw new Exception("No se encontraron datos en la consulta");
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Lista de Artículos");

                    // Agregar encabezados básicos para el test
                    worksheet.Cell(1, 1).Value = "Primera";
                    worksheet.Cell(1, 2).Value = "Código";
                    worksheet.Cell(1, 3).Value = "Cuarto";
                    worksheet.Cell(1, 4).Value = "Marca";
                    worksheet.Cell(1, 5).Value = "Descripción";
                    worksheet.Cell(1, 6).Value = "Observaciones";
                    worksheet.Cell(1, 7).Value = "Costo sin IVA";
                    worksheet.Cell(1, 8).Value = "Costo con IVA (temp)";
                    //worksheet.Cell(1, 9).Value = "Precio de Lista";

                    // Dar formato al encabezado
                    var headerRow = worksheet.Row(1);
                    headerRow.Style.Font.Bold = true;
                    headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                    // Agregar datos - TEST SIN CALCULO DE IVA PRIMERO
                    int row = 2;
                    foreach (DataRow fila in ds.Tables["articulo"].Rows)
                    {
                        // Obtener valores base
                        double costoSinIva = 0;
                        double precioLista = 0;
                        
                        if (fila["costo_sin_iva"] != DBNull.Value)
                            costoSinIva = Convert.ToDouble(fila["costo_sin_iva"]);
                            
                        if (fila["precio_lista"] != DBNull.Value)
                            precioLista = Convert.ToDouble(fila["precio_lista"]);

                        // TEST: Cálculo de IVA temporal fijo de 21% por ahora
                        double costoConIva = costoSinIva * 1.21;

                        // Llenar las celdas
                        worksheet.Cell(row, 1).Value = fila["arprim"]?.ToString() ?? "";
                        worksheet.Cell(row, 2).Value = fila["arterc"]?.ToString() ?? "";
                        worksheet.Cell(row, 3).Value = fila["arcuar"]?.ToString() ?? "";
                        worksheet.Cell(row, 4).Value = fila["armarc"]?.ToString() ?? "";
                        worksheet.Cell(row, 5).Value = fila["ardesc"]?.ToString() ?? "";
                        worksheet.Cell(row, 6).Value = fila["arobse"]?.ToString() ?? "";
                        worksheet.Cell(row, 7).Value = costoSinIva;
                        worksheet.Cell(row, 8).Value = costoConIva;
                        //worksheet.Cell(row, 9).Value = precioLista;
                        row++;
                    }

                    // Ajustar el ancho de las columnas automáticamente
                    worksheet.Columns().AdjustToContents();

                    // Dar formato a las columnas de precios como moneda
                    var rangoCostoSinIva = worksheet.Column(7).CellsUsed();
                    rangoCostoSinIva.Style.NumberFormat.Format = "$#,##0.00";
                    
                    var rangoCostoConIva = worksheet.Column(8).CellsUsed();
                    rangoCostoConIva.Style.NumberFormat.Format = "$#,##0.00";
                    
                    var rangoPrecioLista = worksheet.Column(9).CellsUsed();
                    rangoPrecioLista.Style.NumberFormat.Format = "$#,##0.00";

                    // Guardar el archivo
                    workbook.SaveAs(excelPath);
                }

                ds.Dispose();
                if (cnn.State == ConnectionState.Open) cnn.Close();

                // CORREGIDO: Mensaje simple y abrir después de aceptar
                if (File.Exists(excelPath))
                {
                    // Mostrar mensaje de éxito simple
                    System.Windows.Forms.MessageBox.Show(
                        "¡Archivo Excel generado exitosamente!", 
                        "Excel Generado", 
                        System.Windows.Forms.MessageBoxButtons.OK, 
                        System.Windows.Forms.MessageBoxIcon.Information);
                    
                    // Después de aceptar el mensaje, abrir el archivo
                    try
                    {
                        var startInfo = new ProcessStartInfo
                        {
                            FileName = excelPath,
                            UseShellExecute = true
                        };
                        Process.Start(startInfo);
                    }
                    catch (Exception openEx)
                    {
                        System.Windows.Forms.MessageBox.Show("No se pudo abrir el archivo automáticamente.\n\n" +
                            "Ubicación: " + excelPath + "\n\nError: " + openEx.Message, 
                            "Error al abrir", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    throw new Exception("El archivo Excel no se generó correctamente en: " + excelPath);
                }
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                
                //// MEJORADO: Mostrar error más detallado para debug
                //string mensajeDetallado = $"Error al generar Excel:\n\n" +
                //    $"Mensaje: {ex.Message}\n\n" +
                //    $"Archivo destino: C:\\DistribuidoraPalomarNet\\DistribuidoraPalomarNet\\Descargas\\lista_excel.xlsx\n\n";

                //// Si es un error de archivo en uso, dar sugerencia
                //if (ex.Message.Contains("being used by another process") || ex.Message.Contains("cannot access the file"))
                //{
                //    mensajeDetallado += "SUGERENCIA: Cierre el archivo Excel si está abierto e intente nuevamente.";
                //}
                
                //System.Windows.Forms.MessageBox.Show(mensajeDetallado, "Error al generar Excel", 
                //    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                
                //throw new Exception("Error en PasarArti_a_xls: " + ex.Message);
            }
            finally
            {
                da.Dispose();
            }
        }
        #endregion
        #region Métodos Principales
        /// <summary>
        /// Carga la tabla <c>listass</c> completa en el DataGridView1 ordenada por <c>codigo</c>.
        /// Actualiza el contador de registros totales.
        /// </summary>
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

        /// <summary>
        /// Carga la tabla <c>listass</c> con filtros opcionales en el DataGridView1.
        /// </summary>
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
                else if (_frmPrincipal.cmbArmadora.SelectedIndex > -1)
                    sql += $" WHERE rubro = '{_frmPrincipal.cmbArmadora.Text}'";

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
        /// <summary>
        /// Abre la conexión <c>cnnO</c> a una base de datos Access de origen (usada durante el proceso de actualización FTP).
        /// </summary>
        /// <param name="dbOrigen">Ruta completa al archivo .mdb/.accdb de origen.</param>
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

        /// <summary>
        /// Abre la conexión <c>cnn</c> a la base de datos del distribuidor (artículos, clientes, familias).
        /// Si la conexión ya estaba abierta, la cierra primero para reutilizarla.
        /// </summary>
        /// <remarks>Ruta: <c>ini_ruta_base</c> (leída desde Config.ini).</remarks>
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

        /// <summary>
        /// Abre la conexión <c>cnnU</c> a la base de datos local del usuario (<c>cliente.dat</c>).
        /// Si la conexión ya estaba abierta, la cierra primero para reutilizarla.
        /// </summary>
        /// <remarks>Ruta: <c>ini_ruta_usu</c> = <c>C:\DistribuidoraPalomarNet\cliente.dat</c>.</remarks>
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

        /// <summary>
        /// Punto de entrada de la inicialización completa de la aplicación.
        /// Se llama al cargar <c>frmPrincipal</c>.
        /// </summary>
        /// <remarks>
        /// Secuencia:
        /// <list type="number">
        /// <item><see cref="LeerConfig"/>: carga Config.ini en variables <c>ini_*</c>.</item>
        /// <item><see cref="Cargar_Usuario"/>: carga datos del cliente desde <c>cnnU</c>.</item>
        /// <item><see cref="funcion_inicial"/>(0): inicializa combos y estructura de la grilla (vacía).</item>
        /// <item><see cref="aplicar_desc_Cliente"/>: carga descuentos 1 y 2 en los campos de la UI.</item>
        /// <item>Muestra la fecha de última actualización de la base.</item>
        /// </list>
        /// </remarks>
        public static async Task Inicializacion()
        {
            try
            {
                //DetectarAplicacion();
                Application.DoEvents();
                LeerConfig();
                Cargar_Usuario();
                //ActualizarFAMIVA();
                await funcion_inicial(0);
                aplicar_desc_Cliente(); // Cargo Descuentos 1 y 2
                _frmPrincipal.lblFechaAct.Text = $"Actualizado al {FechaActualizacion()}";
                Application.DoEvents();
                await SeleccionarGrilla();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error dentro de 'Inicializacion': {ex.Message}");
            }
        }

        /// <summary>
        /// Carga inicial de la grilla de artículos.
        /// </summary>
        /// <remarks>
        /// Comportamiento según <paramref name="opc"/>:
        /// <list type="bullet">
        /// <item><b>0</b>: solo inicialización visual — carga los combos (armadora, líneas, rubros, marcas)
        /// y crea una tabla vacía con las columnas correctas. No consulta artículos.
        /// Usado en la primera carga para mostrar la UI rápidamente.</item>
        /// <item><b>1</b>: ejecuta SELECT completo sobre LISTAS con JOIN a proveedo y carga todos
        /// los artículos en DataGridView1. Activa el timer de ofertas y simula clic en la primera fila.</item>
        /// </list>
        /// </remarks>
        /// <param name="opc">0 = inicialización visual; 1 = carga completa de artículos.</param>
        public static async Task funcion_inicial(int opc)
        {
            using var da = new OleDbDataAdapter();
            try
            {
                ConectarDBArti();
                
                if (opc == 0)
                {
                    // Solo inicialización visual: combos + estructura de columnas vacía.
                    // Se activa _aplicandoFiltros para evitar que los SelectedIndexChanged de los
                    // combos (disparados al setear SelectedIndex durante la carga, ej. cmbMarcas)
                    // ejecuten consultas_Relacionales/Ofertas en este punto: eso cargaba artículos
                    // reales en la grilla que quedaban inmediatamente pisados por la tabla vacía de abajo.
                    if (_frmPrincipal != null) _frmPrincipal._aplicandoFiltros = true;
                    try
                    {
                        CargarComboArmadora();
                        CargarComboLineas("0");

                        CargarComboRubros();
                        CargarComboMarcasPorRubro(""); // Cargar todas las marcas inicialmente
                    }
                    finally
                    {
                        if (_frmPrincipal != null) _frmPrincipal._aplicandoFiltros = false;
                    }

                    // Crear una tabla vacía con las columnas que se usan en la grilla
                    var tablaVacia = new DataTable();
                    tablaVacia.Columns.Add("Codigo", typeof(string));
                    tablaVacia.Columns.Add("Descripcion", typeof(string));
                    tablaVacia.Columns.Add("Marca", typeof(string));
                    tablaVacia.Columns.Add("Montadora", typeof(string));
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
                        "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                        //"a.armarc as Marca," +
                        //"l.lidesc as Linea," +
                        "a.armode as Linea," + 
                        " a.arinte as Posición," +
                        //" a.ARMARC as Marca2," + 
                        "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                        ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                        " from (listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) " +
                        "LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI " +
                        //" left join lineas as l on a.arcuar=l.licodi= " +
                        "where a.arprim<>'000' order by a.arterc", cnn);

                    var ds = new DataSet();
                    da.Fill(ds);

                    _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                    _frmPrincipal.DataGridView1.Refresh();

                    await SeleccionarGrilla();
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
        /// <summary>
        /// Configura el estilo visual, ancho proporcional (FillWeight) y visibilidad de las columnas
        /// de <c>DataGridView1</c> en <c>frmPrincipal</c>.
        /// </summary>
        /// <remarks>
        /// Columnas ocultas: <c>aroo</c>, <c>artt</c>, <c>arubic1</c>, <c>arobse</c>, <c>Posición</c>.
        /// Pesos proporcionales: Codigo=7, Marca=7, Descripcion=30, Linea=30, Preciopublico=15.
        /// </remarks>
        public static void Definir_Grilla()
        {
          try
          {
            var dgv = _frmPrincipal?.DataGridView1;
            if (dgv == null) return;

            // Si aún no hay columnas o falta alguna esperada, no reconfigurar (evita excepciones)
            string[] columnasRequeridas =
            {
                "Codigo", "Marca", "Descripcion", "Montadora", "Linea",
                "Posición", "Preciopublico", "Stock", "aroo", "artt", "arubic1", "arobse"
            };
            foreach (var nombre in columnasRequeridas)
            {
                if (!dgv.Columns.Contains(nombre))
                    return;
            }

            var miEstilo = new DataGridViewCellStyle(dgv.Columns["Preciopublico"].DefaultCellStyle)
            {
                Font = new Font("Arial", 12, FontStyle.Bold)
            };

            var miEstiloStock = new DataGridViewCellStyle(dgv.Columns["Stock"].DefaultCellStyle)
            {
                Font = new Font("Arial", 2, FontStyle.Regular)
            };

            dgv.Columns["Stock"].DefaultCellStyle = miEstiloStock;
            dgv.Columns["Preciopublico"].HeaderCell.Style = miEstilo;
            dgv.Columns["Preciopublico"].HeaderCell.Style.BackColor = Color.Black;
            dgv.Columns["Preciopublico"].HeaderCell.Style.ForeColor = Color.White;
            dgv.Columns["Preciopublico"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["Preciopublico"].HeaderText = "Precio";

            // Columnas ocultas
            dgv.Columns["aroo"].Visible    = false;
            dgv.Columns["artt"].Visible    = false;
            dgv.Columns["arubic1"].Visible = false;
            dgv.Columns["arobse"].Visible  = false;
            if (dgv.Columns.Contains("ProvCodi"))
                dgv.Columns["ProvCodi"].Visible = false;

            // Orden de columnas visibles: Codigo / Marca / Descripcion / Montadora / Linea / Posición / Precio / Stock
            dgv.Columns["Posición"].Visible  = true;
            dgv.Columns["Montadora"].Visible = true;

            dgv.Columns["Codigo"].DisplayIndex       = 0;
            dgv.Columns["Marca"].DisplayIndex        = 1;
            dgv.Columns["Descripcion"].DisplayIndex  = 2;
            dgv.Columns["Montadora"].DisplayIndex    = 3;
            dgv.Columns["Linea"].DisplayIndex        = 4;
            dgv.Columns["Posición"].DisplayIndex     = 5;
            dgv.Columns["Preciopublico"].DisplayIndex = 6;
            dgv.Columns["Stock"].DisplayIndex        = 7;

            // Ajuste proporcional de ancho
            dgv.Columns["Codigo"].AutoSizeMode        = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["Marca"].AutoSizeMode         = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["Descripcion"].AutoSizeMode   = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["Montadora"].AutoSizeMode     = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["Linea"].AutoSizeMode         = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["Posición"].AutoSizeMode      = DataGridViewAutoSizeColumnMode.Fill;
            dgv.Columns["Preciopublico"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            // Pesos proporcionales (total ~100)
            dgv.Columns["Codigo"].FillWeight        = 7;
            dgv.Columns["Marca"].FillWeight         = 8;
            dgv.Columns["Descripcion"].FillWeight   = 28;
            dgv.Columns["Montadora"].FillWeight     = 12;
            dgv.Columns["Linea"].FillWeight         = 12;
            dgv.Columns["Posición"].FillWeight      = 8;
            dgv.Columns["Preciopublico"].FillWeight = 15;
          }
          catch (Exception ex)
          {
            GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'Definir_Grilla': {ex.Message}");
          }
        }

        /// <summary>
        /// Actualiza el contador de artículos en el label y selecciona el primer registro de la grilla,
        /// disparando así el pipeline de cálculo de precios vía <see cref="AccionGrilla"/>.
        /// </summary>
        public static async Task SeleccionarGrilla()
        {
            try
            {
                if (_frmPrincipal.DataGridView1.RowCount != 0)
                {
                    _frmPrincipal.lblTotalRegistros.Text = $"ARTÍCULO: {_frmPrincipal.DataGridView1.RowCount}";
                    _frmPrincipal.DataGridView1.CurrentCell = _frmPrincipal.DataGridView1.Rows[0].Cells[0];
                    await AccionGrilla(_frmPrincipal.DataGridView1);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error dentro 'SeleccionarGrilla': {ex.Message}");
            }
        }

        /// <summary>
        /// Pipeline de procesamiento al seleccionar una fila en la grilla de artículos.
        /// </summary>
        /// <remarks>
        /// Secuencia de acciones:
        /// <list type="number">
        /// <item><see cref="Cargar_Imagen"/>: carga la foto del artículo o la imagen por defecto.</item>
        /// <item><see cref="ValidarEquivalencias"/>: habilita/deshabilita el botón de equivalencias.</item>
        /// <item><see cref="TraerDatosExtras"/>: carga ubicación, observación, stock y familia del artículo.</item>
        /// <item><see cref="ColorearStock"/>: colorea el campo de stock según semáforo.</item>
        /// <item><see cref="Calculadora"/>: calcula la cadena de precios (desc1 → desc2 → descFam → IVA → precio_final).</item>
        /// <item><see cref="aplicar_beneficio"/>: calcula el precio sugerido de venta.</item>
        /// <item>Muestra el descuento de familia e IVA de familia en la UI.</item>
        /// </list>
        /// </remarks>
        /// <param name="dataGrid">DataGridView que disparó el evento de selección.</param>
        public static async Task AccionGrilla(DataGridView dataGrid)
        {
            try
            {
                // Bloquear visualización de precios si el cliente no tiene descuentos configurados
                if (!DescuentosConfigurados())
                {
                    _frmPrincipal.txtPreciolista.Text      = string.Empty;
                    _frmPrincipal.txtDescuento1.Text       = string.Empty;
                    _frmPrincipal.txtDescuento2.Text       = string.Empty;
                    _frmPrincipal.txtPreciopublico.Text    = string.Empty;
                    _frmPrincipal.txtDescuentoFamilia.Text = string.Empty;
                    _frmPrincipal.txtIvaFlia.Text          = string.Empty;
                    _frmPrincipal.MostrarNotificacionDescuento();
                    return;
                }

                _frmPrincipal.OcultarNotificacionDescuento();
                // Cargar imagen en el hilo de UI (Cargar_Imagen toca controles).
                // La descarga FTP se offloadea internamente con su propio Task.Run,
                // por lo que la UI no se bloquea mientras se descarga.
                _ = Cargar_Imagen();
                // Por ahora comento las equivalencias
                if(dataGrid.Name== "DataGridView1") ValidarEquivalencias();
                
                if (dataGrid.CurrentCell == null)
                    return;

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

        /// <summary>
        /// Carga los datos complementarios del artículo seleccionado en la grilla:
        /// ubicación (<c>arubic1</c>), observación (<c>arobse</c>), stock (<c>arff</c>)
        /// y código de familia (<c>arflia</c>) en la variable interna <c>_ARTI_ARFLIA</c>.
        /// </summary>
        /// <param name="numFila">Índice de la fila actualmente seleccionada en DataGridView1.</param>
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

        /// <summary>
        /// Carga la imagen del artículo actualmente seleccionado en el PictureBox de la UI.
        /// Primero intenta cargar desde la carpeta local FOTOS.
        /// Si no existe localmente y el servicio FTP está configurado, descarga desde FTP con caché.
        /// Si no existe la imagen del artículo, muestra la imagen por defecto <c>SINFOTO.png</c>.
        /// </summary>
        public static async Task Cargar_Imagen()
        {
            try
            {
                int numFila = _frmPrincipal.DataGridView1.CurrentCellAddress.Y;
                if (numFila < 0)
                {
                    MostrarImagenPorDefecto();
                    return;
                }

                string codigoArticulo = _frmPrincipal.DataGridView1["Codigo", numFila].Value?.ToString();
                
                if (string.IsNullOrWhiteSpace(codigoArticulo))
                {
                    MostrarImagenPorDefecto();
                    return;
                }

                string basePath = ObtenerCarpetaFotosBase();
                
                // 1. Intentar cargar desde carpeta local primero (más rápido) - probar .jpg y .png
                string[] extensionesLocales = { ".jpg", ".png" };
                foreach (var ext in extensionesLocales)
                {
                    string localImagePath = Path.Combine(basePath, $"{codigoArticulo}{ext}");
                    if (File.Exists(localImagePath))
                    {
                        using (var fs = new FileStream(localImagePath, FileMode.Open, FileAccess.Read))
                        {
                            _frmPrincipal.ptbImagen.Image = Image.FromStream(fs);
                        }
                        return;
                    }
                }

                // 2. Si no existe localmente y el servicio FTP está configurado, intentar descargar
                // El servicio ya busca automáticamente en .jpg y .png
                if (_fotoService != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[Cargar_Imagen] Solicitando foto FTP para: {codigoArticulo}");
                    string rutaFotoCache = await Task.Run(async () => await _fotoService.ObtenerRutaLocalFotoAsync(codigoArticulo));
                    System.Diagnostics.Debug.WriteLine($"[Cargar_Imagen] Resultado FTP: {(rutaFotoCache ?? "null")}");

                    if (rutaFotoCache != null && File.Exists(rutaFotoCache))
                    {
                        string rutaFotoBase = null;
                        try
                        {
                            Directory.CreateDirectory(basePath);
                            rutaFotoBase = Path.Combine(basePath, Path.GetFileName(rutaFotoCache));
                            bool debeCopiar = true;
                            if (File.Exists(rutaFotoBase))
                            {
                                DateTime baseWrite = File.GetLastWriteTime(rutaFotoBase);
                                DateTime cacheWrite = File.GetLastWriteTime(rutaFotoCache);
                                debeCopiar = cacheWrite > baseWrite;
                            }

                            if (debeCopiar)
                                File.Copy(rutaFotoCache, rutaFotoBase, overwrite: true);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[Cargar_Imagen] No se pudo copiar foto desde caché a base: {ex.GetType().Name}: {ex.Message}");
                            rutaFotoBase = null;
                        }

                        string rutaParaMostrar = (rutaFotoBase != null && File.Exists(rutaFotoBase))
                            ? rutaFotoBase
                            : rutaFotoCache;

                        using (var fs = new FileStream(rutaParaMostrar, FileMode.Open, FileAccess.Read))
                        {
                            _frmPrincipal.ptbImagen.Image = Image.FromStream(fs);
                        }
                        return;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[Cargar_Imagen] _fotoService es null, no se intentará FTP");
                }

                // 3. Si no se encontró la imagen, mostrar imagen por defecto
                MostrarImagenPorDefecto();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Cargar_Imagen] Error: {ex.Message}");
                MostrarImagenPorDefecto();
            }
        }

        /// <summary>
        /// Muestra la imagen por defecto SINFOTO.png en el PictureBox.
        /// </summary>
        private static void MostrarImagenPorDefecto()
        {
            try
            {
                string defaultImagePath = ObtenerRutaImagenPorDefecto();
                if (File.Exists(defaultImagePath))
                {
                    using (var fs = new FileStream(defaultImagePath, FileMode.Open, FileAccess.Read))
                    {
                        _frmPrincipal.ptbImagen.Image = Image.FromStream(fs);
                    }
                }
                else
                {
                    _frmPrincipal.ptbImagen.Image = null;
                }
            }
            catch
            {
                _frmPrincipal.ptbImagen.Image = null;
            }
        }

        private static string ObtenerCarpetaImagenes()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                var dir = new DirectoryInfo(baseDir);
                while (dir != null)
                {
                    string candidata = Path.Combine(dir.FullName, "imagenes");
                    if (Directory.Exists(candidata))
                        return candidata;

                    dir = dir.Parent;
                }
            }
            catch
            {
            }

            return Path.Combine(baseDir, "imagenes");
        }

        private static string ObtenerCarpetaFotosBase()
        {
            return Path.Combine(ObtenerCarpetaImagenes(), "FOTOS");
        }

        public static string ObtenerCarpetaFotosBasePublica()
        {
            return ObtenerCarpetaFotosBase();
        }

        private static string ObtenerRutaImagenPorDefecto()
        {
            return Path.Combine(ObtenerCarpetaImagenes(), "SINFOTO.png");
        }

        public static string ObtenerRutaImagenPorDefectoPublica()
        {
            return ObtenerRutaImagenPorDefecto();
        }

        /// <summary>
        /// Método de prueba para cargar la foto de un artículo específico.
        /// Simula la carga de foto sin necesidad de seleccionar en el DataGridView.
        /// </summary>
        /// <param name="codigoArticulo">Código del artículo a buscar (ej: RE4214)</param>
        public static async void TestCargarFotoArticulo(string codigoArticulo)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"========================================");
                System.Diagnostics.Debug.WriteLine($"[TEST] Iniciando prueba para artículo: {codigoArticulo}");
                System.Diagnostics.Debug.WriteLine($"========================================");

                if (string.IsNullOrWhiteSpace(codigoArticulo))
                {
                    System.Diagnostics.Debug.WriteLine($"[TEST] ERROR: Código de artículo vacío");
                    MessageBox.Show("Código de artículo vacío", "Error de Test", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 1. Verificar si el servicio FTP está inicializado
                if (_fotoService == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[TEST] ADVERTENCIA: Servicio FTP no inicializado");
                    MessageBox.Show("Servicio FTP no inicializado.\nVerifica Config.ini", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[TEST] ✓ Servicio FTP inicializado correctamente");
                }

                // 2. Intentar cargar desde carpeta local primero - buscar .jpg y .png
                string basePath = ObtenerCarpetaFotosBase();
                string[] extensiones = { ".jpg", ".png" };
                
                foreach (var ext in extensiones)
                {
                    string localImagePath = Path.Combine(basePath, $"{codigoArticulo}{ext}");
                    System.Diagnostics.Debug.WriteLine($"[TEST] Buscando en carpeta local: {localImagePath}");
                    
                    if (File.Exists(localImagePath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[TEST] ✓ Foto encontrada en carpeta local ({ext})");
                        using (var fs = new FileStream(localImagePath, FileMode.Open, FileAccess.Read))
                        {
                            _frmPrincipal.ptbImagen.Image = Image.FromStream(fs);
                        }
                        MessageBox.Show(
                            $"Foto cargada desde carpeta local:\n{localImagePath}",
                            "Test Exitoso - Local",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[TEST] ✗ Foto NO encontrada en carpeta local (.jpg ni .png)");

                // 3. Si no existe localmente y el servicio FTP está configurado, intentar descargar
                if (_fotoService != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[TEST] Intentando descargar desde FTP...");
                    System.Diagnostics.Debug.WriteLine($"[TEST] Host: {ini_ftp_fotos_host}");
                    System.Diagnostics.Debug.WriteLine($"[TEST] Puerto: {ini_ftp_fotos_port}");
                    System.Diagnostics.Debug.WriteLine($"[TEST] Usuario: {ini_ftp_fotos_user}");
                    System.Diagnostics.Debug.WriteLine($"[TEST] Carpeta remota: {ini_ftp_fotos_folder}");
                    System.Diagnostics.Debug.WriteLine($"[TEST] Buscando extensiones: .jpg y .png");

                    string rutaFotoCache = await _fotoService.ObtenerRutaLocalFotoAsync(codigoArticulo);

                    if (rutaFotoCache != null && File.Exists(rutaFotoCache))
                    {
                        System.Diagnostics.Debug.WriteLine($"[TEST] ✓ Foto descargada/encontrada en caché: {rutaFotoCache}");
                        using (var fs = new FileStream(rutaFotoCache, FileMode.Open, FileAccess.Read))
                        {
                            _frmPrincipal.ptbImagen.Image = Image.FromStream(fs);
                        }

                        // Verificar estadísticas de caché
                        int cantidadCache = _fotoService.ObtenerCantidadFotosEnCache();
                        long tamanoCache = _fotoService.ObtenerTamanoCache();
                        string tamanoMB = (tamanoCache / 1024.0 / 1024.0).ToString("F2");

                        MessageBox.Show(
                            $"Foto descargada desde FTP y almacenada en caché:\n{rutaFotoCache}\n\n" +
                            $"Estadísticas de caché:\n" +
                            $"- Fotos en caché: {cantidadCache}\n" +
                            $"- Tamaño total: {tamanoMB} MB",
                            "Test Exitoso - FTP",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[TEST] ✗ Foto NO encontrada en FTP");
                    }
                }

                // 4. Si no se encontró la imagen, mostrar imagen por defecto
                System.Diagnostics.Debug.WriteLine($"[TEST] Mostrando imagen por defecto (SINFOTO.png)");
                MostrarImagenPorDefecto();
                MessageBox.Show(
                    $"Foto no encontrada para el artículo: {codigoArticulo}\n\n" +
                    "Verificaciones realizadas:\n" +
                    "✗ Carpeta local\n" +
                    "✗ Servidor FTP\n\n" +
                    "Mostrando imagen por defecto (SINFOTO.png)",
                    "Test - Foto No Encontrada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                System.Diagnostics.Debug.WriteLine($"========================================");
                System.Diagnostics.Debug.WriteLine($"[TEST] Fin de la prueba");
                System.Diagnostics.Debug.WriteLine($"========================================");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TEST] ERROR: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[TEST] StackTrace: {ex.StackTrace}");
                MessageBox.Show(
                    $"Error durante el test:\n{ex.Message}",
                    "Error de Test",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                MostrarImagenPorDefecto();
            }
        }

        /// <summary>
        /// Colorea el campo de stock (<c>txtStock</c>) en la UI según un semáforo de tres niveles:
        /// Rojo (&lt;= 1), Amarillo (= 2) o Verde (&gt;= 3).
        /// </summary>
        /// <remarks>
        /// Los umbrales están hardcodeados actualmente en el código.
        /// Las variables <c>ini_cond_rojo/amarillo/verde</c> están disponibles pero no se usan aquí.
        /// </remarks>
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

        /// <summary>
        /// Calcula y muestra el precio sugerido de venta al público aplicando el beneficio
        /// familiar (<c>FABEN1</c>) del artículo seleccionado.
        /// </summary>
        /// <remarks>
        /// Flujo:
        /// <list type="number">
        /// <item>Busca <c>FABEN1</c> directamente en la tabla <c>FAMILIAS</c> de <c>cnnU</c> (base local) por <c>FACODI = _ARTI_ARFLIA</c>.</item>
        /// <item>Si hay beneficio: <c>precioSugerido = baseSugerido + (baseSugerido × FABEN1 / 100)</c>.</item>
        /// </list>
        /// La base del sugerido es <c>precio_final_desc3</c> si hay desc3 activo; sino <c>precio_final</c>.
        /// <b>Nota:</b> se filtra por <c>FACODI</c> (no por <c>FADESC</c>) porque puede haber familias con
        /// igual descripción pero distinto proveedor (<c>FAPROV</c>).
        /// </remarks>
        public static void aplicar_beneficio()
        {
            try
            {
                // Base para sugerido: si hay desc3 aplicado, usar precio_final_desc3; si no, precio_final
                decimal baseSugerido = precio_final_desc3 != 0 && _frmPrincipal.txtDescuento3.Text != "0"
                    ? precio_final_desc3
                    : precio_final;

                // Buscar FABEN1 directamente por FACODI en base de usuario
                decimal? faben1 = null;
                if (_ARTI_ARFLIA != 0)
                {
                    using (var cmUsu = new OleDbCommand("SELECT TOP 1 FABEN1 FROM FAMILIAS WHERE FACODI = ?", cnnU))
                    {
                        try
                        {
                            ConectarDBUsuario();
                            cmUsu.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
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

        /// <summary>
        /// Aplica el tercer descuento opcional (<c>txtDescuento3</c>) sobre <c>precio_final</c>.
        /// Actualiza <c>txtPreciopublico</c> con el resultado.
        /// </summary>
        /// <remarks>
        /// Solo se ejecuta si hay valor en <c>txtDescuento1</c> y <c>txtDescuento3</c>.
        /// Si desc3 = 0, muestra <c>precio_final</c> directamente.
        /// </remarks>
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

        /// <summary>
        /// Núcleo del cálculo de precios. Se ejecuta cada vez que se selecciona un artículo en la grilla.
        /// Calcula la cadena completa de precios y actualiza las variables globales y la UI.
        /// </summary>
        /// <remarks>
        /// Cadena cuando <c>Descuento1 &gt; 0</c>:
        /// <code>
        /// precio_desc1         = precio_lista − (precio_lista × desc1 / 100)
        /// precio_desc2         = precio_desc1 − (precio_desc1 × desc2 / 100)
        /// precio_desc2        -= precio_desc2 × descFamilia / 100   [si FADES1 != 0]
        /// precio_costo_sin_iva = precio_desc2
        /// precio_final         = precio_desc2 + (precio_desc2 × IVA familia / 100)
        /// </code>
        /// Cadena cuando <c>Descuento1 = 0</c>: se usa <c>precio_lista</c> como base y aplica solo IVA de familia.
        /// Los valores de IVA y descuento de familia se consultan en <c>cnnU</c> (tabla FAMILIAS local).
        /// </remarks>
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

                            // aplicar IVA por familia si existe (consulta directa por FACODI: evita
                            // ambigüedad cuando dos familias comparten FADESC con distinto FAPROV)
                            decimal precioConIva = precio_desc2;
                            if (_ARTI_ARFLIA != 0)
                            {
                                using (var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA FROM FAMILIAS WHERE FACODI = ?", cnnU))
                                {
                                    try
                                    {
                                        ConectarDBUsuario();
                                        cmUsu.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
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

                            // aplicar IVA por familia también cuando no hay descuentos (consulta directa
                            // por FACODI: evita ambigüedad cuando dos familias comparten FADESC con distinto FAPROV)
                            decimal precioConIva = precio_costo_sin_iva;
                            if (_ARTI_ARFLIA != 0)
                            {
                                using (var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA FROM FAMILIAS WHERE FACODI = ?", cnnU))
                                {
                                    try
                                    {
                                        ConectarDBUsuario();
                                        cmUsu.Parameters.AddWithValue("@p1", _ARTI_ARFLIA);
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

        /// <summary>
        /// Retorna el porcentaje de descuento de familia (<c>FADES1</c>) del artículo actualmente seleccionado.
        /// </summary>
        /// <remarks>
        /// Consulta la tabla <c>FAMILIAS</c> en <c>cnnU</c> filtrando por <c>_ARTI_ARFLIA</c>.
        /// Si no hay familia o la consulta falla, retorna 0.
        /// </remarks>
        /// <returns>Porcentaje de descuento familiar (ej: 5.0 para el 5%), o 0 si no aplica.</returns>
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
        /// <summary>
        /// Retorna el porcentaje de IVA de familia (<c>FAIVA</c>) del artículo actualmente seleccionado.
        /// </summary>
        /// <remarks>
        /// Consulta la tabla <c>FAMILIAS</c> en <c>cnnU</c> filtrando por <c>_ARTI_ARFLIA</c>.
        /// Si no hay familia o la consulta falla, retorna 0.
        /// </remarks>
        /// <returns>Porcentaje de IVA familiar (ej: 21.0 para el 21%), o 0 si no aplica.</returns>
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
        /// <summary>
        /// Carga o actualiza la grilla del carrito (<c>gridPedidos</c>) con todos los ítems
        /// de la vista <c>PedidosRealizados</c> y calcula el total sin IVA.
        /// </summary>
        /// <remarks>
        /// Se llama automáticamente después de cada operación sobre PREPARAC
        /// (agregar, actualizar cantidad, eliminar) para mantener la UI sincronizada.
        /// Lee de <c>cnnU</c> usando una vista o consulta que expone:
        /// número, código, marca, descripción, código origen, precio sin IVA, cantidad y subtotal.
        /// </remarks>
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
        //                await SeleccionarGrilla();
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

        /// <summary>
        /// Ejecuta la búsqueda dinámica de artículos en la base del distribuidor
        /// aplicando todos los filtros activos simultáneamente.
        /// </summary>
        /// <remarks>
        /// Filtros disponibles (se combinan con AND):
        /// <list type="bullet">
        /// <item><c>txtDescripcion</c>: búsqueda multipalabra sobre <c>ARDESC</c>.</item>
        /// <item><c>txtModelo</c>: búsqueda multipalabra sobre <c>ARUBIC1</c>.</item>
        /// <item><c>cmbArmadora</c>: filtro por RUBROS (JOIN con LISTAS.ARRUB2).</item>
        /// <item><c>cmbMarcas</c>: filtro por proveedor (<c>ARPROV = prcodi</c>).</item>
        /// <item><c>cmbRubros</c>: filtro por modelo/línea (<c>ARMODE = rubro</c>).</item>
        /// </list>
        /// La condición multipalabra es generada por <see cref="GenerarCondicionMultipalabra"/>.
        /// Límite de resultados: 10 000 filas. Límite de longitud de query: 4 000 caracteres.
        /// </remarks>
        /// <summary>
        /// Ejecuta la consulta SQL a la BD Access usando SOLO los filtros superiores
        /// (Código, Descripción, Observación/Código Universal) concatenados con WHERE AND.
        /// El resultado se almacena en <see cref="_dtMaestro"/> para que los filtros inferiores
        /// (Rubro, Armadora, Líneas, Marca) filtren sobre estos datos sin volver a consultar la BD.
        /// </summary>
        public static async Task consultas_Relacionales()
        {
            if (_frmPrincipal == null)
                throw new ArgumentNullException("El formulario principal no puede ser nulo");
            if (cnn == null)
                throw new InvalidOperationException("La conexión a la base de datos no está disponible");

            using var da = new OleDbDataAdapter();
            string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    "a.arpre2 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    ", a.arprov as ProvCodi " +
                    " from ((listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI) ";

            try
            {
                using var ds = new DataSet();
                // --- FILTROS SUPERIORES: generan condiciones WHERE contra la BD ---
                var conditions = new System.Collections.Generic.List<string>();
                conditions.Add("a.arprim<>'000'");

                // txtCodigo: busca por concatenación arprim+arterc
                string txtCod = _frmPrincipal.txtCodigo?.Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(txtCod))
                    conditions.Add($"a.arprim+a.arterc like '%{txtCod}%'");

                // txtDescripcion: búsqueda multipalabra sobre ARDESC
                string condDesc = GenerarCondicionMultipalabra("a.ARDESC", _frmPrincipal.txtDescripcion.Text);
                if (!string.IsNullOrEmpty(condDesc))
                    conditions.Add(condDesc);

                // txtModelo (Observación/Código Universal): búsqueda multipalabra sobre ARUBIC1
                string condModelo = GenerarCondicionMultipalabra("a.ARUBIC1", _frmPrincipal.txtModelo.Text);
                if (!string.IsNullOrEmpty(condModelo))
                    conditions.Add(condModelo);

                // cmbRubros: ahora representa una FAMILIA (FADESC). Como una misma FADESC puede
                // corresponder a varios FACODI (ej: "ROTULA Y EXTREMO" -> 3, 4, 6), se resuelven todos
                // y se filtra por a.ARFLIA IN (...) en vez de buscar texto en la descripción.
                string rubroSel = _frmPrincipal.cmbRubros?.Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(rubroSel) && rubroSel != "TODOS"
                    && !rubroSel.Contains("DataRowView"))
                {
                    var facodisRubro = ObtenerFacodisPorFamilia(rubroSel);
                    if (facodisRubro.Count > 0)
                        conditions.Add(ConstruirClausulaIn("a.ARFLIA", facodisRubro));
                }

                // cmbArmadora: filtra por r.rbdesc (JOIN RUBROS ya presente en el FROM)
                string armadoraSel = _frmPrincipal.cmbArmadora?.Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(armadoraSel) && armadoraSel != "TODOS" && armadoraSel != " TODOS"
                    && !armadoraSel.Contains("DataRowView"))
                    conditions.Add($"r.rbdesc = '{armadoraSel.Replace("'", "''")}'");

                // cmbMarcas: filtra por a.arprov = prcodi (SelectedValue del combo)
                if (_frmPrincipal.cmbMarcas?.SelectedValue is int marcaVal && marcaVal != 0)
                    conditions.Add($"a.arprov = {marcaVal}");

                // cmbLineas: filtra por a.armode = lidesc (Text del combo)
                string lineaSel = _frmPrincipal.cmbLineas?.Text?.Trim() ?? string.Empty;
                string lineaValStr = _frmPrincipal.cmbLineas?.SelectedValue?.ToString() ?? "0";
                if (!string.IsNullOrEmpty(lineaSel) && lineaSel != "TODOS" && lineaValStr != "0"
                    && !lineaSel.Contains("DataRowView"))
                    conditions.Add($"a.armode = '{lineaSel.Replace("'", "''")}'");

                string WhereQuery = " WHERE " + string.Join(" AND ", conditions);
                string QueryFinal = SelectQuery + WhereQuery + " order by a.arterc";

                if (QueryFinal.Length > 4000)
                    throw new InvalidOperationException("La consulta es demasiado larga");

                ConectarDBArti();
                da.SelectCommand = new OleDbCommand(QueryFinal, cnn);
                da.Fill(ds);

                if (ds?.Tables?.Count > 0 && ds.Tables[0]?.Rows?.Count > 0)
                {
                    // Guardar DataTable maestro (referencia para uso futuro de filtrado en memoria)
                    _dtMaestro = ds.Tables[0];
                    _dtMaestro.TableName = "Articulos";

                    _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                    _frmPrincipal.DataGridView1.Refresh();
                    Definir_Grilla();

                    // NOTA: filtrado en memoria comentado para uso futuro (ver AplicarFiltrosLocales)
                    // AplicarFiltrosLocales();

                    if (_frmPrincipal.DataGridView1.RowCount > 0)
                        await SeleccionarGrilla();
                }
                else
                {
                    _dtMaestro = null;
                    _frmPrincipal.DataGridView1.DataSource = null;
                    _frmPrincipal.DataGridView1.Rows.Clear();
                    _frmPrincipal.DataGridView1.Refresh();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en consultas_Relacionales: {ex.Message}");
            }
            finally
            {
                if (_frmPrincipal?.TimerVerOfertas != null && !_frmPrincipal.TimerVerOfertas.Enabled)
                    _frmPrincipal.TimerVerOfertas.Enabled = true;
            }
        }

        // ============================================================================
        // AplicarFiltrosLocales: filtrado en memoria sobre _dtMaestro (DataView.RowFilter).
        // Comentado a pedido del usuario (30/08/2026): los combos (Armadora, Rubro, Marca,
        // Línea) ahora filtran directo contra la BD vía consultas_Relacionales/consultas_Ofertas,
        // igual que los TextBox. Se deja el código para retomar el enfoque de filtrado en
        // memoria en una futura ocasión si se requiere.
        // ============================================================================
        ///// <summary>
        ///// Aplica los filtros inferiores (Rubro/Armadora/Líneas/Marca) sobre el DataTable maestro
        ///// ya cargado en memoria, sin volver a consultar la BD Access.
        ///// Usa DataView.RowFilter para filtrar los datos visibles en la grilla.
        ///// </summary>
        //public static async void AplicarFiltrosLocales()
        //{
        //    if (_frmPrincipal == null) return;

        //    try
        //    {
        //        // Si no hay DataTable maestro (no se hizo búsqueda previa), traer todos los datos
        //        if (_dtMaestro == null)
        //        {
        //            await consultas_Relacionales();
        //            return; // consultas_Relacionales ya llama a AplicarFiltrosLocales internamente
        //        }

        //        var filtros = new System.Collections.Generic.List<string>();

        //        // Filtro por Armadora (Montadora = rbdesc de RUBROS con RBCODI > 100)
        //        string armadora = _frmPrincipal.cmbArmadora?.Text?.Trim() ?? string.Empty;
        //        if (!string.IsNullOrEmpty(armadora) && armadora != "TODOS" && armadora != " TODOS"
        //            && !armadora.Contains("DataRowView"))
        //            filtros.Add($"Montadora = '{armadora.Replace("'", "''")}'");

        //        // Nota: cmbRubros NO filtra localmente — va directo a la BD via consultas_Relacionales()

        //        // Filtro por Marca (ProvCodi = SelectedValue del combo cmbMarcas)
        //        if (_frmPrincipal.cmbMarcas?.SelectedValue is int marcaVal && marcaVal != 0)
        //            filtros.Add($"ProvCodi = {marcaVal}");

        //        // Filtro por Línea (cmbLineas → lidesc = armode = columna Linea)
        //        // SelectedValue es string (licodi); "0" = TODOS
        //        string lineaSel = _frmPrincipal.cmbLineas?.Text?.Trim() ?? string.Empty;
        //        string lineaValStr = _frmPrincipal.cmbLineas?.SelectedValue?.ToString() ?? "0";
        //        if (!string.IsNullOrEmpty(lineaSel) && lineaSel != "TODOS" && lineaValStr != "0"
        //            && !lineaSel.Contains("DataRowView"))
        //            filtros.Add($"Linea = '{lineaSel.Replace("'", "''")}'");

        //        // Aplicar el RowFilter al DataView
        //        string rowFilter = filtros.Count > 0 ? string.Join(" AND ", filtros) : string.Empty;

        //        DataView dv = _dtMaestro.DefaultView;
        //        dv.RowFilter = rowFilter;

        //        _frmPrincipal.DataGridView1.DataSource = dv;
        //        _frmPrincipal.DataGridView1.Refresh();
        //        Definir_Grilla();

        //        // Pintar semáforo de Stock
        //        _frmPrincipal.ColorearGrillaxOfertas(_frmPrincipal.DataGridView1);

        //        // Actualizar contador de registros
        //        _frmPrincipal.lblTotalRegistros.Text = "ARTÍCULO: " + _frmPrincipal.DataGridView1.RowCount.ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        GrabarError($" - {DateTime.Now:dd/MM/yyyy HH:mm:ss} 'AplicarFiltrosLocales': {ex.Message}");
        //    }
        //}

        /// <summary>
        /// Igual a <see cref="consultas_Relacionales"/> pero sobre artículos en oferta (<c>ARTT &lt;&gt; 0</c>).
        /// Muestra el precio de oferta (<c>arpre5</c>) en lugar del precio de lista.
        /// </summary>
        /// <remarks>
        /// Aplica los mismos filtros combinados que <see cref="consultas_Relacionales"/>.
        /// La condición adicional <c>a.ARTT &lt;&gt; 0</c> garantiza que solo se muestren artículos con precio de oferta.
        /// </remarks>
        public static async Task consultas_Ofertas()
        {
            if (_frmPrincipal == null)
                throw new ArgumentNullException("El formulario principal no puede ser nulo");
            if (cnn == null)
                throw new InvalidOperationException("La conexión a la base de datos no está disponible");

            using var da = new OleDbDataAdapter();
            string SelectQuery = "SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    "a.arpre5 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo " +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    ", a.arprov as ProvCodi " +
                    " from ((listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI) ";

            try
            {
                using var ds = new DataSet();
                // --- FILTROS SUPERIORES: generan condiciones WHERE contra la BD ---
                var conditions = new System.Collections.Generic.List<string>();
                conditions.Add("a.ARTT<>0");
                conditions.Add("a.arprim<>'000'");

                // txtCodigo: busca por concatenación arprim+arterc
                string txtCod = _frmPrincipal.txtCodigo?.Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(txtCod))
                    conditions.Add($"a.arprim+a.arterc like '%{txtCod}%'");

                // txtDescripcion: búsqueda multipalabra sobre ARDESC
                string condDesc = GenerarCondicionMultipalabra("a.ARDESC", _frmPrincipal.txtDescripcion.Text);
                if (!string.IsNullOrEmpty(condDesc))
                    conditions.Add(condDesc);

                // txtModelo (Observación/Código Universal): búsqueda multipalabra sobre ARUBIC1
                string condModelo = GenerarCondicionMultipalabra("a.ARUBIC1", _frmPrincipal.txtModelo.Text);
                if (!string.IsNullOrEmpty(condModelo))
                    conditions.Add(condModelo);

                // cmbRubros: ahora representa una FAMILIA (FADESC). Como una misma FADESC puede
                // corresponder a varios FACODI (ej: "ROTULA Y EXTREMO" -> 3, 4, 6), se resuelven todos
                // y se filtra por a.ARFLIA IN (...) en vez de buscar texto en la descripción.
                string rubroSelOf = _frmPrincipal.cmbRubros?.Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(rubroSelOf) && rubroSelOf != "TODOS"
                    && !rubroSelOf.Contains("DataRowView"))
                {
                    var facodisRubroOf = ObtenerFacodisPorFamilia(rubroSelOf);
                    if (facodisRubroOf.Count > 0)
                        conditions.Add(ConstruirClausulaIn("a.ARFLIA", facodisRubroOf));
                }

                // cmbArmadora: filtra por r.rbdesc (JOIN RUBROS ya presente en el FROM)
                string armadoraSelOf = _frmPrincipal.cmbArmadora?.Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(armadoraSelOf) && armadoraSelOf != "TODOS" && armadoraSelOf != " TODOS"
                    && !armadoraSelOf.Contains("DataRowView"))
                    conditions.Add($"r.rbdesc = '{armadoraSelOf.Replace("'", "''")}'");

                // cmbMarcas: filtra por a.arprov = prcodi (SelectedValue del combo)
                if (_frmPrincipal.cmbMarcas?.SelectedValue is int marcaValOf && marcaValOf != 0)
                    conditions.Add($"a.arprov = {marcaValOf}");

                // cmbLineas: filtra por a.armode = lidesc (Text del combo)
                string lineaSelOf = _frmPrincipal.cmbLineas?.Text?.Trim() ?? string.Empty;
                string lineaValStrOf = _frmPrincipal.cmbLineas?.SelectedValue?.ToString() ?? "0";
                if (!string.IsNullOrEmpty(lineaSelOf) && lineaSelOf != "TODOS" && lineaValStrOf != "0"
                    && !lineaSelOf.Contains("DataRowView"))
                    conditions.Add($"a.armode = '{lineaSelOf.Replace("'", "''")}'");

                string WhereQuery = " WHERE " + string.Join(" AND ", conditions);
                string QueryFinal = SelectQuery + WhereQuery + " order by a.arterc";

                if (QueryFinal.Length > 4000)
                    throw new InvalidOperationException("La consulta es demasiado larga");

                ConectarDBArti();
                da.SelectCommand = new OleDbCommand(QueryFinal, cnn);
                da.Fill(ds);

                if (ds?.Tables?.Count > 0 && ds.Tables[0]?.Rows?.Count > 0)
                {
                    // Guardar DataTable maestro (referencia para uso futuro de filtrado en memoria)
                    _dtMaestro = ds.Tables[0];
                    _dtMaestro.TableName = "Articulos";

                    _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                    _frmPrincipal.DataGridView1.Refresh();
                    Definir_Grilla();

                    // NOTA: filtrado en memoria comentado para uso futuro (ver AplicarFiltrosLocales)
                    // AplicarFiltrosLocales();

                    if (_frmPrincipal.DataGridView1.RowCount > 0)
                        await SeleccionarGrilla();
                }
                else
                {
                    _dtMaestro = null;
                    _frmPrincipal.DataGridView1.DataSource = null;
                    _frmPrincipal.DataGridView1.Rows.Clear();
                    _frmPrincipal.DataGridView1.Refresh();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en consultas_Ofertas: {ex.Message}");
            }
            finally
            {
                if (_frmPrincipal?.TimerVerOfertas != null && !_frmPrincipal.TimerVerOfertas.Enabled)
                    _frmPrincipal.TimerVerOfertas.Enabled = true;
            }
        }


        /// <summary>
        /// Genera una condición SQL LIKE multipalabra para el campo especificado.
        /// Cada palabra del texto se convierte en una condición <c>LIKE '%palabra%'</c> unidas por AND.
        /// </summary>
        /// <param name="campo">Nombre del campo SQL (ej: <c>"a.ARDESC"</c>).</param>
        /// <param name="texto">Texto de búsqueda; si está vacío retorna cadena vacía.</param>
        /// <returns>Condición SQL entre paréntesis, o cadena vacía si no hay palabras.</returns>
        /// <example>GenerarCondicionMultipalabra("a.ARDESC", "filtro aceite") →
        /// <c>(a.ARDESC like '%filtro%' AND a.ARDESC like '%aceite%')</c></example>
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
        /// <summary>
        /// Lee el archivo <c>C:\DistribuidoraPalomarNet\Config.ini</c> línea a línea
        /// y carga todas las variables <c>ini_*</c> del módulo.
        /// </summary>
        /// <remarks>
        /// El archivo INI tiene un formato posicional (sin secciones); cada línea comienza con
        /// una clave prefija (ej: <c>"Emisor "</c>, <c>"RutaFTP "</c>) seguida del valor.
        /// El orden de las líneas es fijo y debe coincidir con el esperado por este método.
        /// Si el archivo no existe, muestra un aviso y retorna sin cargar nada.
        /// <b>Nota:</b> <c>ini_ruta_usu</c> siempre se inicializa a
        /// <c>C:\DistribuidoraPalomarNet\cliente.dat</c> independientemente del INI.
        /// </remarks>
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
                        // RutaBannerGrande
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 16 && cadena.Substring(0, 16) == "RutaBannerGrande")
                        {
                            largoCadena = cadena.Length;
                            ini_ruta_banner_grande = cadena.Substring(17, largoCadena - 17);
                        }
                        // RutaBannerBoton
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 15 && cadena.Substring(0, 15) == "RutaBannerBoton")
                        {
                            largoCadena = cadena.Length;
                            ini_ruta_banner_boton = cadena.Substring(16, largoCadena - 16);
                        }
                        // LinkBannerGrande
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 16 && cadena.Substring(0, 16) == "LinkBannerGrande")
                        {
                            largoCadena = cadena.Length;
                            ini_link_banner_grande = cadena.Substring(17, largoCadena - 17);
                        }
                        // LinkBannerBoton
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 15 && cadena.Substring(0, 15) == "LinkBannerBoton")
                        {
                            largoCadena = cadena.Length;
                            ini_link_banner_boton = cadena.Substring(16, largoCadena - 16);
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

                        // *******************Configuración FTP para Fotos*********************
                        // FtpFotosHost
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 12 && cadena.Substring(0, 12) == "FtpFotosHost")
                        {
                            largoCadena = cadena.Length;
                            ini_ftp_fotos_host = cadena.Substring(13, largoCadena - 13);
                        }

                        // FtpFotosPort
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 12 && cadena.Substring(0, 12) == "FtpFotosPort")
                        {
                            largoCadena = cadena.Length;
                            string portStr = cadena.Substring(13, largoCadena - 13);
                            if (int.TryParse(portStr, out int port))
                                ini_ftp_fotos_port = port;
                        }

                        // FtpFotosUser
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 12 && cadena.Substring(0, 12) == "FtpFotosUser")
                        {
                            largoCadena = cadena.Length;
                            ini_ftp_fotos_user = cadena.Substring(13, largoCadena - 13);
                        }

                        // FtpFotosPass
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 12 && cadena.Substring(0, 12) == "FtpFotosPass")
                        {
                            largoCadena = cadena.Length;
                            ini_ftp_fotos_pass = cadena.Substring(13, largoCadena - 13);
                        }

                        // FtpFotosFolder (opcional, tiene valor por defecto)
                        cadena = leer.ReadLine();
                        if (cadena != null && cadena.Length > 14 && cadena.Substring(0, 14) == "FtpFotosFolder")
                        {
                            largoCadena = cadena.Length;
                            ini_ftp_fotos_folder = cadena.Substring(15, largoCadena - 15);
                        }
                    }

                    // Inicializar servicio de caché de fotos FTP si hay configuración válida
                    if (!string.IsNullOrWhiteSpace(ini_ftp_fotos_host) && 
                        !string.IsNullOrWhiteSpace(ini_ftp_fotos_user) && 
                        !string.IsNullOrWhiteSpace(ini_ftp_fotos_pass))
                    {
                        _fotoService = new FtpPhotoCacheService(
                            ini_ftp_fotos_host,
                            ini_ftp_fotos_port,
                            ini_ftp_fotos_user,
                            ini_ftp_fotos_pass,
                            ini_ftp_fotos_folder
                        );
                        System.Diagnostics.Debug.WriteLine($"[Module] Servicio FTP de fotos inicializado: Host={ini_ftp_fotos_host}:{ini_ftp_fotos_port} Folder={ini_ftp_fotos_folder}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[Module] Configuración FTP de fotos incompleta, servicio no inicializado");
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

        /// <summary>
        /// Descarga asincrónicamente el texto del banner desde la URL configurada
        /// y lo aplica al formulario principal mediante <c>ActualizarTexto</c>.
        /// </summary>
        /// <param name="ini_ruta_banner">URL del recurso de texto (puede ser un archivo de texto plano en un servidor web).</param>
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
        /// <summary>Realiza una petición HTTP GET y retorna el contenido de la respuesta como string.</summary>
        /// <param name="url">URL del recurso a descargar.</param>
        /// <returns>Contenido del recurso como texto plano.</returns>
        /// <exception cref="HttpRequestException">Si el servidor retorna un código de error HTTP.</exception>
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
        /// <summary>
        /// Retorna la fecha de la última actualización de la base de artículos,
        /// leyendo el campo <c>arpp</c> de la tabla LISTAS.
        /// </summary>
        /// <returns>Fecha formateada como <c>"dd/MM/yyyy"</c>, o cadena vacía si no se encuentra.</returns>
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

        /// <summary>
        /// Obtiene la fecha de última modificación de un archivo en el servidor FTP.
        /// </summary>
        /// <param name="Filename">Nombre del archivo en el servidor FTP.</param>
        /// <param name="NewFileName">Ruta local (no usado en este método, solo por firma).</param>
        /// <param name="UserNameFtp">Usuario FTP.</param>
        /// <param name="PassFTP">Contraseña FTP.</param>
        /// <param name="URI">URL base del servidor FTP (ej: <c>ftp://servidor.com</c>).</param>
        /// <returns>Fecha formateada como <c>"dd/MM/yyyy"</c>, o <c>"01/01/1900"</c> si hay error.</returns>
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

        /// <summary>
        /// Actualiza la cantidad, precio unitario y precio total de un artículo en la tabla <c>PEDIDOS</c>
        /// de la base local del usuario (tabla diferente de PREPARAC).
        /// </summary>
        /// <param name="cod">Código del artículo en PEDIDOS.</param>
        /// <param name="pun">Nuevo precio unitario.</param>
        /// <param name="Cantidad">Nueva cantidad como cadena.</param>
        /// <param name="pt">Nuevo precio total.</param>
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

        /// <summary>
        /// Proceso completo de actualización de la base de artículos vía FTP.
        /// </summary>
        /// <remarks>
        /// Muestra <see cref="frmProcesoActualizacion"/> (WebView2 con HTML moderno de 7 pasos)
        /// durante todo el proceso. Cada paso se marca como activo/completado/error en tiempo real.
        /// <list type="number">
        /// <item><b>Paso 0</b> — Descarga el archivo de actualización del FTP con <see cref="DescargaArchivo"/>.</item>
        /// <item><b>Paso 1</b> — Ejecuta el instalador/extractor descargado (<c>Process.Start</c>).</item>
        /// <item><b>Paso 2</b> — Verifica que el archivo de base exista en <c>ini_ruta_base_descarga</c>.</item>
        /// <item><b>Paso 3</b> — Migra la base con <see cref="UpdateBase"/> (copia origen → destino).</item>
        /// <item><b>Paso 4</b> — <see cref="SincronizarDescuentosCliente"/>: actualiza <c>USUARIO.US_DESCUENTO1/2</c>
        ///              con los valores frescos de <c>CLIENTES.CLDES1/2</c>, luego llama <see cref="Cargar_Usuario"/>.</item>
        /// <item><b>Paso 5</b> — Si el cliente está autorizado: <see cref="Update_Precios_Gral"/> →
        ///              <see cref="aplicar_desc_Cliente"/> → <see cref="funcion_inicial"/>(1).</item>
        /// <item><b>Paso 6</b> — Finalización: actualiza fecha y oculta alerta de update disponible.</item>
        /// </list>
        /// En cualquier error, el paso correspondiente se marca en rojo y se muestra el mensaje en la barra inferior.
        /// </remarks>
        public static void ActualizacionXFTP()
        {
            bool ret;
            frmProcesoActualizacion frmProceso = null;

            try
            {
                _frmPrincipal.Timer1.Start();
                _frmPrincipal.toolStripProgressBar1.Visible = true;
                Application.DoEvents();

                // Mostrar el formulario de progreso HTML/WebView2
                frmProceso = new frmProcesoActualizacion();
                frmProceso.Show();
                Application.DoEvents();
                frmProceso.WaitReady(); // espera hasta que WebView2 termine de cargar el HTML

                // ── Paso 0: Descargar del servidor FTP ────────────────────────────
                frmProceso.SetStep(0, "active");
                frmProceso.SetMensaje("Conectando y descargando desde el servidor FTP...");

                ret = DescargaArchivo(ini_arch_ftp,
                                    @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\" + ini_arch_ftp,
                                    ini_usu_ftp,
                                    ini_pass_ftp,
                                    ini_ruta_ftp);

                if (ret)
                {
                    frmProceso.SetStep(0, "done");

                    // ── Paso 1: Extraer archivos del ZIP ──────────────────────
                    frmProceso.SetStep(1, "active");
                    frmProceso.SetMensaje("Extrayendo archivos del paquete ZIP descargado...");

                    // NUEVO: Extraer el ZIP usando System.IO.Compression
                    string rutaDescargaZip = @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\" + ini_arch_ftp;
                    bool zipExtraido = ExtractZipFile(rutaDescargaZip, @"C:\DistribuidoraPalomarNet\DistribuidoraPalomarNet\Descargas\");
                    Application.DoEvents();

                    if (zipExtraido && File.Exists(ini_ruta_base_descarga))
                    {
                        frmProceso.SetStep(1, "done");

                        // ── Paso 2: Verificar archivos descargados ────────────────
                        frmProceso.SetStep(2, "active");
                        frmProceso.SetMensaje("Verificando integridad de los archivos...");
                        Application.DoEvents();
                        frmProceso.SetStep(2, "done");

                        // ── Paso 3: Migrar / actualizar la base de datos ──────────
                        frmProceso.SetStep(3, "active");
                        frmProceso.SetMensaje("Actualizando base de datos local...");

                        Migra = UpdateBase(ini_ruta_base_descarga, ini_ruta_base);

                        if (Migra)
                        {
                            frmProceso.SetStep(3, "done");

                            // ── Paso 4: Sincronizar descuentos del cliente ────────
                            frmProceso.SetStep(4, "active");
                            frmProceso.SetMensaje("Sincronizando descuentos desde la base del distribuidor...");

                            Update_FechaDB();
                            Application.DoEvents();
                            // Lee CLDES1/CLDES2 desde la nueva CLIENTES y los persiste en USUARIO (cliente.dat)
                            SincronizarDescuentosCliente();
                            Application.DoEvents();
                            Cargar_Usuario();
                            _frmPrincipal.ApplyEstadoClienteToUI();

                            frmProceso.SetStep(4, "done");

                            if (!ClienteAutorizado)
                            {
                                // El cliente no tiene descuentos habilitados: finalizar sin recalcular
                                _frmPrincipal.ckbOfertas.Checked = false;
                                _frmPrincipal.lblFechaAct.Text = "Actualizado al  " + FechaActualizacion();
                                BuscarActualizacionDisponible();
                                _frmPrincipal.cmbArmadora.DropDownWidth = 280;
                                _frmPrincipal.lblAlertaUpdate.Visible = false;
                                frmProceso.SetStep(6, "done");
                                frmProceso.SetFinalizado();
                                frmProceso.SetMensaje("Base actualizada (cliente sin descuentos habilitados).");
                                System.Threading.Thread.Sleep(1800);
                                Application.DoEvents();
                                Migra = false;
                                return;
                            }

                            // ── Paso 5: Recalcular precios ────────────────────────
                            frmProceso.SetStep(5, "active");
                            frmProceso.SetMensaje("Recalculando precios con los nuevos descuentos...");

                            Update_Precios_Gral(); //recalculo los precios
                            aplicar_desc_Cliente(); //Cargo Descuentos 1 y 2 (antes de recargar grilla)
                            _frmPrincipal.ckbOfertas.Checked = false;
                            //funcion_inicial(1);
                            //funcion_inicial_carrito();
                            _frmPrincipal.lblFechaAct.Text = "Actualizado al  " + FechaActualizacion();
                            BuscarActualizacionDisponible();
                            _frmPrincipal.cmbArmadora.DropDownWidth = 280;

                            frmProceso.SetStep(5, "done");

                            // ── Paso 6: Proceso completado ────────────────────────
                            frmProceso.SetStep(6, "done");
                            frmProceso.SetFinalizado();
                            frmProceso.SetMensaje("¡Actualización completada con éxito!");
                            System.Threading.Thread.Sleep(1800);
                            Application.DoEvents();

                            _frmPrincipal.lblAlertaUpdate.Visible = false;
                            Migra = false;
                        }
                        else
                        {
                            frmProceso.SetStep(3, "error");
                            frmProceso.SetError("No se pudo terminar la actualización entre bases.");
                            System.Threading.Thread.Sleep(2000);
                            Application.DoEvents();
                            MessageBox.Show("No Se pudo Terminar la Actualizacion entre Bases");
                        }
                    }
                    else
                    {
                        frmProceso.SetStep(1, "error");
                        frmProceso.SetError("No se pudo extraer el archivo ZIP o no se encontró el archivo de base de datos.");
                        System.Threading.Thread.Sleep(2000);
                        Application.DoEvents();
                        MessageBox.Show("No se pudo extraer el archivo ZIP o no se encontró BASEPALOMARNET.mdb en el ZIP");
                    }
                }
                else
                {
                    frmProceso.SetStep(0, "error");
                    frmProceso.SetError("No se pudo descargar el archivo de actualización del FTP.");
                    System.Threading.Thread.Sleep(2000);
                    Application.DoEvents();
                    MessageBox.Show("No se Pudo descargar el archivo de Actualizacion");
                }

                Application.DoEvents();
                _frmPrincipal.toolStripProgressBar1.Visible = false;
                _frmPrincipal.Timer1.Stop();
            }
            catch (Exception ex)
            {
                if (frmProceso != null)
                {
                    frmProceso.SetError("Error: " + ex.Message);
                    System.Threading.Thread.Sleep(1500);
                    Application.DoEvents();
                }
                _frmPrincipal.toolStripProgressBar1.Visible = false;
                _frmPrincipal.Timer1.Stop();
                MessageBox.Show("Error dentro de Menu:Archivo-->Actualizar Base" + ex.Message,
                               "Alerta",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
            finally
            {
                if (frmProceso != null)
                {
                    try
                    {
                        frmProceso.Close();
                        frmProceso.Dispose();
                    }
                    catch { }
                }
            }
        }
        /// <summary>
        /// Reemplaza la base de datos de destino con la base de origen descargada del FTP.
        /// Elimina el archivo de destino si existe antes de copiar.
        /// </summary>
        /// <param name="Origen">Ruta del archivo fuente (base descargada, <c>ini_ruta_base_descarga</c>).</param>
        /// <param name="Destino">Ruta del archivo destino (base activa, <c>ini_ruta_base</c>).</param>
        /// <returns><c>true</c> si la copia fue exitosa; <c>false</c> ante cualquier error.</returns>
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

        /// <summary>
        /// Extrae un archivo ZIP a la carpeta de destino especificada.
        /// </summary>
        /// <param name="zipFilePath">Ruta completa del archivo ZIP a extraer.</param>
        /// <param name="extractPath">Carpeta destino donde extraer los archivos.</param>
        /// <returns>true si la extracción fue exitosa, false en caso contrario.</returns>
        public static bool ExtractZipFile(string zipFilePath, string extractPath)
        {
            try
            {
                if (!File.Exists(zipFilePath))
                {
                    GrabarError($"Archivo ZIP no encontrado: {zipFilePath}");
                    return false;
                }

                // Crear la carpeta de destino si no existe
                if (!Directory.Exists(extractPath))
                {
                    Directory.CreateDirectory(extractPath);
                }

                // Extraer el archivo ZIP sobrescribiendo archivos existentes
                ZipFile.ExtractToDirectory(zipFilePath, extractPath, overwriteFiles: true);
                
                return true;
            }
            catch (Exception ex)
            {
                GrabarError($"Error al extraer ZIP {zipFilePath}: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// Escribe la fecha del archivo FTP en el campo <c>arpp</c> de todos los registros de LISTAS.
        /// Esto permite que <see cref="FechaActualizacion"/> y <see cref="BuscarActualizacionDisponible"/>
        /// muestren la fecha correcta de la última actualización.
        /// </summary>
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

        /// <summary>
        /// Compara la fecha del archivo de actualización en el FTP con la fecha de la base local.
        /// Si hay una versión más nueva, muestra la alerta en la UI.
        /// </summary>
        /// <returns><c>true</c> si hay actualización disponible; <c>false</c> si la base está al día.</returns>
        /// <remarks>
        /// <b>Advertencia:</b> el bloque <c>catch</c> contiene código inalcanzable (<c>ret=false</c>
        /// después del <c>throw</c>). En caso de error, la función siempre relanza la excepción.
        /// </remarks>
        /// <summary>
        /// Verifica de forma asíncrona si hay una actualización disponible en el servidor FTP sin bloquear la UI.
        /// </summary>
        public static async Task<bool> BuscarActualizacionDisponibleAsync()
        {
            try
            {
                string fechaFTP = await Task.Run(() => 
                    FechaArchivoFTP(ini_arch_ftp, ini_ruta_gral + "\\DistribuidoraPalomarNet\\Descargas\\" + ini_arch_ftp, 
                                   ini_usu_ftp, ini_pass_ftp, ini_ruta_ftp)).ConfigureAwait(false);

                if (Convert.ToDateTime(fechaFTP) != Convert.ToDateTime("01/01/1900"))
                {
                    if (Convert.ToDateTime(fecha) < Convert.ToDateTime(fechaFTP))
                    {
                        _frmPrincipal.Invoke((Action)(() =>
                        {
                            _frmPrincipal.lblAlertaUpdate.Visible = true;
                            _frmPrincipal.lblAlertaUpdate.Text = "Hay una Actualizacion disponible, por favor actualice la base de datos";
                        }));
                        return true;
                    }
                    else
                    {
                        _frmPrincipal.Invoke((Action)(() => _frmPrincipal.lblAlertaUpdate.Visible = false));
                        return false;
                    }
                }
                else
                {
                    _frmPrincipal.Invoke((Action)(() => _frmPrincipal.lblAlertaUpdate.Visible = false));
                    return false;
                }
            }
            catch (Exception ex)
            {
                _frmPrincipal.Invoke((Action)(() => _frmPrincipal.lblAlertaUpdate.Visible = false));
                GrabarError($" - {DateTime.Now:dd/MM/yyyy HH:mm:ss} 'BuscarActualizacionDisponibleAsync': {ex.Message}");
                return false;
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

        /// <summary>
        /// Genera el PDF del pedido actual y lo envía por correo electrónico
        /// a la distribuidora y, opcionalmente, al cliente (si tiene email válido).
        /// </summary>
        /// <remarks>
        /// Flujo:
        /// <list type="number">
        /// <item>Genera el PDF en <c>C:\...\pedidos\Remito {US_CODI} {timestamp}.pdf</c>.</item>
        /// <item>Si <c>cboxNoEnviar</c> no está marcado, envía el PDF por SMTP.</item>
        /// <item>El asunto incluye la nota del pedido si <c>txtNotaPedido</c> tiene texto.</item>
        /// <item>Si el email del cliente es válido (<see cref="ValidaEMail"/>), se agrega como CC.</item>
        /// </list>
        /// </remarks>
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
        /// <summary>
        /// Valida el formato de una dirección de email usando reglas básicas.
        /// </summary>
        /// <param name="EMail">Dirección de email a validar.</param>
        /// <returns><c>true</c> si contiene exactamente un <c>@</c> y el dominio tiene al menos un punto.</returns>
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
        /// <summary>
        /// Construye y agrega al documento iTextSharp el contenido del pedido:
        /// logo, encabezado, fecha y tabla de artículos desde <c>gridPedidos</c>.
        /// </summary>
        /// <remarks>
        /// Excluye las columnas 0 (número interno) y 4 (precio) de la tabla PDF.
        /// Los anchos de columna son fijos (ver <see cref="GetColumnasSize"/>).
        /// </remarks>
        /// <param name="document">Documento iTextSharp ya abierto donde se agregan los elementos.</param>
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

                // **NUEVO: Crear lista de filas ordenadas por PPACODI (columna "Codigo" que está en posición 1)**
                var filasOrdenadas = new List<DataGridViewRow>();
                
                // Agregar todas las filas del grid a la lista
                for (int i = 0; i < _frmCarroCompras.gridPedidos.RowCount; i++)
                {
                    filasOrdenadas.Add(_frmCarroCompras.gridPedidos.Rows[i]);
                }
                
                // Ordenar por la columna "Codigo" (posición 1) que contiene PPACODI
                filasOrdenadas.Sort((x, y) =>
                {
                    string codigoX = x.Cells[1].Value?.ToString() ?? "";
                    string codigoY = y.Cells[1].Value?.ToString() ?? "";
                    return string.Compare(codigoX, codigoY, StringComparison.OrdinalIgnoreCase);
                });

                // Se generan las columnas del DataGridView ordenadas por PPACODI
                foreach (var fila in filasOrdenadas)
                {
                    for (int j = 0; j < _frmCarroCompras.gridPedidos.ColumnCount; j++)
                    {
                        if (j != 0 && j != 4) // saco columnas de id y precios#lo agregue gmedina
                        {
                            string valor = fila.Cells[j].Value?.ToString() ?? "";
                            datatable.AddCell(new PdfPCell(new Paragraph(valor, FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.ITALIC))));
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
        /// <summary>
        /// Retorna un array de anchos de columna fijos para la tabla PDF del pedido.
        /// </summary>
        /// <remarks>Los valores están hardcodeados: {105, 95, 420, 215, 125, 175}.
        /// El loop dinámico basado en el DataGridView está comentado.</remarks>
        /// <param name="dg">DataGridView de referencia (solo se usa para determinar el tamaño del array).</param>
        /// <returns>Array float[] con anchos de columna en puntos para la tabla PDF.</returns>
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
        /// <summary>
        /// Orquesta el recálculo completo de precios tras una actualización de la base.
        /// Deshabilita la grilla durante el proceso para evitar interacciones.
        /// </summary>
        /// <remarks>Llama a <see cref="Cargar_Usuario"/>, <see cref="Update_Precios"/>,
        /// <see cref="Update_Precios_Ofertas"/> y <see cref="funcion_inicial"/>(0) en ese orden.</remarks>
        public static async Task Update_Precios_Gral()
        {
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");
                
                _frmPrincipal.DataGridView1.Enabled = false;
                Cargar_Usuario();
                Update_Precios();
                Update_Precios_Ofertas();
                await funcion_inicial(0);
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

        /// <summary>Recalcula los precios de lista para todos los artículos normales.
        /// Llama a <see cref="Update_PrecioCosto"/> (cadena desc1→desc2→IVA) y <see cref="Update_PrecioSug"/> (precio sugerido).
        /// </summary>
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
        /// <summary>Recalcula los precios para los artículos en oferta.
        /// Llama a <see cref="Update_PrecioCosto_Ofertas"/> y <see cref="Update_PrecioSug_Ofertas"/>.
        /// </summary>
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

        /// <summary>Colorea de amarillo las filas de la grilla cuyos artículos tienen precio de oferta (<c>aroo = "1"</c>).</summary>
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

        /// <summary>Restaura el color de fondo blanco en todas las filas de la grilla de artículos.</summary>
        public static void Despintar_Ofertas()
        {
            if (_frmPrincipal.DataGridView1.RowCount > 0)
            {
                _frmPrincipal.DataGridView1.DefaultCellStyle.BackColor = Color.White;
                _frmPrincipal.DataGridView1.DefaultCellStyle.ForeColor = Color.Black;
                _frmPrincipal.DataGridView1.Refresh();
            }
        }

        /// <summary>Agrega un mensaje de error al archivo <c>error.log</c> en el directorio actual de la aplicación.
        /// Silencia los errores de escritura para no interrumpir el flujo principal.</summary>
        /// <param name="mensaje">Texto del error a registrar.</param>
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

        /// <summary>
        /// Ejecuta una sentencia SQL de modificación sobre la base del <b>distribuidor</b> (<c>cnn</c>) de forma asíncrona.
        /// </summary>
        /// <remarks><b>Nota:</b> a pesar del nombre "DBLocal", opera sobre <c>cnn</c> (base del distribuidor), no sobre <c>cnnU</c>.</remarks>
        /// <param name="query">Sentencia SQL a ejecutar (UPDATE, INSERT, DELETE).</param>
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

        /// <summary>
        /// Ejecuta una sentencia SQL de modificación sobre la base local del usuario (<c>cnnU</c>).
        /// </summary>
        /// <remarks>
        /// <b>Bug conocido:</b> en el bloque finally cierra <c>cnn</c> en lugar de <c>cnnU</c>.
        /// </remarks>
        /// <param name="query">Sentencia SQL a ejecutar (UPDATE, INSERT, DELETE).</param>
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

        /// <summary>
        /// Carga en <c>cmbArmadora</c> los rubros principales (código &gt; 100) que tienen
        /// artículos asociados en LISTAS, ordenados alfabéticamente.
        /// Agrega «TODOS» como primera opción.
        /// </summary>
        /// <param name="facodisFiltro">Opcional. Si se provee (con la familia/rubro seleccionada en
        /// <c>cmbRubros</c>), restringe las armadoras a las que tienen artículos con
        /// <c>ARFLIA</c> dentro de esta lista.</param>
        public static void CargarComboArmadora(List<string> facodisFiltro = null)
        {
            Application.DoEvents();
            ConectarDBArti();
            string sql;
            if (facodisFiltro != null && facodisFiltro.Count > 0)
            {
                sql = "SELECT DISTINCT R.rbcodi, R.rbdesc " +
                    "FROM rubros AS R INNER JOIN listas AS LT ON LT.ARRUB2 = R.RBCODI " +
                    "WHERE R.RBCODI > 100 AND " + ConstruirClausulaIn("LT.ARFLIA", facodisFiltro) +
                    " ORDER BY R.rbdesc";
            }
            else
            {
                sql = "SELECT DISTINCT R.rbcodi, R.rbdesc " +
                    "FROM rubros AS R " +
                    "WHERE R.RBCODI > 100 ORDER BY R.rbdesc";
            }
            var da = new OleDbDataAdapter(sql, cnn);
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");
                
                _frmPrincipal.cmbArmadora.Items.Clear();
                var ds = new DataSet();
                Application.DoEvents();
                da.Fill(ds, "rubros");
                
                foreach (DataRow fila in ds.Tables["rubros"].Rows)
                {
                    Application.DoEvents();
                    _frmPrincipal.cmbArmadora.Items.Add(fila["rbdesc"]);
                }
                _frmPrincipal.cmbArmadora.Items.Insert(0, "TODOS");
                ds.Dispose();
                if (cnn.State == ConnectionState.Open) cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception($", dentro de 'CargarComboArmadora': {ex.Message}");
            }
            finally
            {
                da.Dispose();
            }
        }
        /// <summary>
        /// Carga en <c>cmbLineas</c> las líneas (<c>lineas.lidesc</c>) disponibles,
        /// opcionalmente filtradas por una armadora/rubro específico.
        /// </summary>
        /// <param name="_Armadora">Nombre del rubro para filtrar (ej: valor de <c>cmbArmadora</c>)
        /// o <c>"0"</c> para cargar todas las líneas sin filtro.</param>
        /// <param name="facodisFiltro">Opcional. Si se provee (con la familia/rubro seleccionada en
        /// <c>cmbRubros</c>), restringe además las líneas a las que tienen artículos con
        /// <c>ARFLIA</c> dentro de esta lista.</param>
        public static void CargarComboLineas(string _Armadora, List<string> facodisFiltro = null)
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

                // Construir SQL base y agregar filtros opcionales por armadora y/o familia (ARFLIA)
                bool filtrarPorArmadora = !string.IsNullOrEmpty(_Armadora) && _Armadora != "0";
                bool filtrarPorFamilia = facodisFiltro != null && facodisFiltro.Count > 0;

                string sql = filtrarPorArmadora
                    ? "SELECT DISTINCT LI.licodi, LI.lidesc " +
                      "FROM ((lineas AS LI INNER JOIN listas AS LT ON LI.lidesc = LT.ARMODE) " +
                      "INNER JOIN RUBROS AS r ON LT.ARRUB2 = r.RBCODI) "
                    : "SELECT DISTINCT LI.licodi, LI.lidesc " +
                      "FROM (lineas AS LI INNER JOIN listas AS LT ON LI.lidesc = LT.ARMODE) ";

                var wheres = new List<string>();
                if (filtrarPorArmadora)
                    wheres.Add("r.RBDESC = ?");
                if (filtrarPorFamilia)
                    wheres.Add(ConstruirClausulaIn("LT.ARFLIA", facodisFiltro));

                if (wheres.Count > 0)
                    sql += "WHERE " + string.Join(" AND ", wheres) + " ";

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
        /// <summary>
        /// Carga en <c>cmbRubros</c> las descripciones de familia (<c>familias.FADESC</c>, base <c>cnn</c>)
        /// en lugar de <c>rubros.rbdesc</c>. Se mantiene el nombre "Rubro" en la UI por decisión del negocio,
        /// pero el combo ahora representa familias de artículos.
        /// </summary>
        private static void CargarComboRubros()
        {
            Application.DoEvents();
            ConectarDBArti();
            string sql = "SELECT DISTINCT FADESC FROM familias " +
                "WHERE FADESC IS NOT NULL AND FADESC <> '' ORDER BY FADESC";
            var da = new OleDbDataAdapter(sql, cnn);
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");

                _frmPrincipal.cmbRubros.Items.Clear();
                var ds = new DataSet();
                Application.DoEvents();
                da.Fill(ds, "familias");

                foreach (DataRow fila in ds.Tables["familias"].Rows)
                {
                    Application.DoEvents();
                    _frmPrincipal.cmbRubros.Items.Add(fila["FADESC"]);
                }
                _frmPrincipal.cmbRubros.Items.Insert(0, "TODOS");
                ds.Dispose();
                if (cnn.State == ConnectionState.Open) cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception($", dentro de 'CargarComboRubros': {ex.Message}");
            }
            finally
            {
                da.Dispose();
            }
        }

        /// <summary>
        /// Resuelve todos los <c>FACODI</c> (base <c>cnn</c>, tabla <c>familias</c>) que comparten
        /// una misma descripción de familia (<c>FADESC</c>). Es necesario porque una misma descripción
        /// puede tener varios códigos de familia distintos (ej: "ROTULA Y EXTREMO" → FACODI 3, 4 y 6).
        /// </summary>
        /// <param name="fadesc">Descripción de familia seleccionada en <c>cmbRubros</c>.</param>
        /// <returns>Lista de FACODI (como string) asociados a esa FADESC. Vacía si <paramref name="fadesc"/>
        /// está vacío, es "TODOS", o no se encuentra ninguna coincidencia.</returns>
        public static List<string> ObtenerFacodisPorFamilia(string fadesc)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(fadesc) || fadesc.Trim() == "TODOS" || fadesc.Contains("DataRowView"))
                return result;

            try
            {
                ConectarDBArti();
                using var cmd = new OleDbCommand("SELECT FACODI FROM familias WHERE FADESC = ?", cnn);
                cmd.Parameters.AddWithValue("@p1", fadesc.Trim());
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rdr[0] != DBNull.Value)
                    {
                        string facodi = rdr[0].ToString();
                        if (!string.IsNullOrEmpty(facodi) && !result.Contains(facodi))
                            result.Add(facodi);
                    }
                }
            }
            finally
            {
                if (cnn.State == ConnectionState.Open)
                    cnn.Close();
            }
            return result;
        }

        /// <summary>
        /// Construye una cláusula SQL <c>columna IN (v1,v2,...)</c> a partir de una lista de códigos.
        /// ARFLIA/FACODI son campos numéricos en Access: los valores NO se entrecomillan porque el
        /// motor Jet/ACE lanza "Data type mismatch in criteria expression" al comparar una columna
        /// numérica contra literales de texto. Se filtran valores no numéricos por seguridad.
        /// </summary>
        private static string ConstruirClausulaIn(string columna, List<string> valores)
        {
            var valoresNumericos = valores.Where(v => long.TryParse(v, out _)).ToList();
            string inList = string.Join(",", valoresNumericos);
            return $"{columna} IN ({inList})";
        }

        private static void CargarComboMarcas()
        {
            // Consulta: trae código (prcodi) y nombre de marca (prmarc) de la tabla proveedo.
            // Solo incluye proveedores que sean "Repuestos" (prcrub = 'R'), ordenados alfabéticamente.
            // prcodi  → se usará como VALUE del combo (lo que se lee al seleccionar)
            // prmarc  → se usará como DISPLAY del combo (lo que ve el usuario)
            Application.DoEvents();
            ConectarDBArti();
            string sql = "SELECT DISTINCT prcodi, prmarc FROM proveedo WHERE prcrub = 'R' ORDER BY prmarc";
            var da = new OleDbDataAdapter(sql, cnn);
            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");

                // Crear DataTable local para poder agregar la fila "TODOS" antes de bindear
                var Dt_marcas = new DataTable();
                Dt_marcas.Columns.Add("prcodi", typeof(int));    // valor interno (clave)
                Dt_marcas.Columns.Add("prmarc", typeof(string)); // texto visible en el combo

                // Fila especial "TODOS": prcodi = 0 indica sin filtro activo
                Dt_marcas.Rows.Add(0, "TODOS");

                Application.DoEvents();

                // Llenar con datos reales de la base
                var ds = new DataSet();
                da.Fill(ds, "marcas"); // nombre de tabla interno del DataSet
                Application.DoEvents();

                foreach (DataRow fila in ds.Tables["marcas"].Rows)
                {
                    // Agregar cada marca: prcodi como valor, prmarc como texto
                    Dt_marcas.Rows.Add(
                        fila["prcodi"],  // valor que se leerá con cmbMarcas.SelectedValue
                        fila["prmarc"]   // texto que verá el usuario en la lista
                    );
                    Application.DoEvents();
                }

                // Bindear al combo usando DisplayMember y ValueMember
                // - DisplayMember: columna que se muestra en la lista desplegable
                // - ValueMember:   columna que se retorna con .SelectedValue (prcodi)
                _frmPrincipal.cmbMarcas.DataSource = Dt_marcas;
                _frmPrincipal.cmbMarcas.DisplayMember = "prmarc";
                _frmPrincipal.cmbMarcas.ValueMember = "prcodi";

                // Dejar sin selección al inicializar (ninguna marca activa por defecto)
                _frmPrincipal.cmbMarcas.SelectedIndex = -1;

                ds.Dispose();
                if (cnn.State == ConnectionState.Open) cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception($"Error dentro de 'CargarComboMarcas': {ex.Message}");
            }
            finally
            {
                da.Dispose();
            }
        }

        /// <summary>
        /// Carga las marcas (proveedo) filtradas por la familia (rubro) seleccionada.
        /// Si rubroSeleccionado es "TODOS" o vacío, carga todas las marcas.
        /// </summary>
        /// <param name="rubroSeleccionado">Descripción de familia (<c>FADESC</c>) seleccionada en
        /// <c>cmbRubros</c> para filtrar las marcas.</param>
        public static void CargarComboMarcasPorRubro(string rubroSeleccionado = "")
        {
            Application.DoEvents();
            ConectarDBArti();

            var facodisRubro = ObtenerFacodisPorFamilia(rubroSeleccionado);

            string sql;
            if (facodisRubro.Count == 0)
            {
                // Sin familia seleccionada (o sin coincidencias): mostrar todas las marcas
                sql = "SELECT DISTINCT p.prcodi, p.prmarc FROM proveedo p WHERE p.prcrub = 'R' ORDER BY p.prmarc";
            }
            else
            {
                // Filtrar marcas por familia: la relación pasa por listas (arprov→prcodi, ARFLIA IN facodis)
                sql = "SELECT DISTINCT p.prcodi, p.prmarc " +
                      "FROM (proveedo p INNER JOIN listas a ON a.arprov = p.prcodi) " +
                      "WHERE p.prcrub = 'R' AND " + ConstruirClausulaIn("a.ARFLIA", facodisRubro) + " " +
                      "ORDER BY p.prmarc";
            }

            var da = new OleDbDataAdapter(sql, cnn);

            try
            {
                if (_frmPrincipal == null) throw new InvalidOperationException("Formulario principal no inicializado");

                // Crear DataTable local para poder agregar la fila "TODOS"
                var Dt_marcas = new DataTable();
                Dt_marcas.Columns.Add("prcodi", typeof(int));    // valor interno (clave)
                Dt_marcas.Columns.Add("prmarc", typeof(string)); // texto visible en el combo

                // Fila especial "TODOS": prcodi = 0 indica sin filtro activo
                Dt_marcas.Rows.Add(0, "TODOS");

                Application.DoEvents();

                // Llenar con datos filtrados por rubro
                var ds = new DataSet();
                da.Fill(ds, "marcas");
                Application.DoEvents();

                foreach (DataRow fila in ds.Tables["marcas"].Rows)
                {
                    // Agregar cada marca: prcodi como valor, prmarc como texto
                    Dt_marcas.Rows.Add(
                        fila["prcodi"],  // valor que se leerá con cmbMarcas.SelectedValue
                        fila["prmarc"]   // texto que verá el usuario en la lista
                    );
                    Application.DoEvents();
                }

                // Bindear al combo
                _frmPrincipal.cmbMarcas.DataSource = Dt_marcas;
                _frmPrincipal.cmbMarcas.DisplayMember = "prmarc";
                _frmPrincipal.cmbMarcas.ValueMember = "prcodi";

                // Seleccionar "TODOS" por defecto al cambiar de rubro
                _frmPrincipal.cmbMarcas.SelectedIndex = 0;

                ds.Dispose();
                if (cnn.State == ConnectionState.Open) cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception($"Error dentro de 'CargarComboMarcasPorRubro': {ex.Message}");
            }
            finally
            {
                da.Dispose();
            }
        }
        /// <summary>
        /// Lee los datos del usuario/cliente desde la tabla <c>USUARIO</c> de la base local
        /// y los carga en las variables globales <c>US_*</c>.
        /// </summary>
        /// <remarks>
        /// Si no se encuentran datos (tabla vacía), limpia las variables <c>US_*</c> y,
        /// si <paramref name="promptIfMissing"/> es <c>true</c>, abre <c>frmDatosUsuario</c> para que
        /// el usuario complete su configuración.
        /// Además llama a <see cref="ValidarCuitCliente"/> para verificar el cliente contra
        /// la base del distribuidor y cargar los descuentos.
        /// </remarks>
        /// <param name="promptIfMissing">Si <c>true</c> (por defecto), abre <c>frmDatosUsuario</c>
        /// cuando no hay datos configurados.</param>
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
                cmUsu.CommandText = "SELECT US_CODI, US_CLRAZO, US_CUIT, US_MAIL, US_CLCIVA, US_CLVEND, US_CLCPAG, US_DESCUENTO1, US_DESCUENTO2 from USUARIO";
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
                    US_DESCUENTO1 = Convert.ToInt32(drUsu["US_DESCUENTO1"]);
                    US_DESCUENTO2 = Convert.ToInt32(drUsu["US_DESCUENTO2"]);

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
                    US_DESCUENTO1 = 0;
                    US_DESCUENTO2 = 0;
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

        /// <summary>Verifica que un string sea no nulo, no vacío y parseble como número flotante.</summary>
        /// <param name="valor">Cadena a validar.</param>
        /// <returns><c>true</c> si el valor es numérico; <c>false</c> si es nulo, vacío o no numérico.</returns>
        public static bool ValidarCampoNumerico(string valor)
        {
            return !string.IsNullOrWhiteSpace(valor) && float.TryParse(valor.Trim(), out _);
        }

        /// <summary>
        /// Valida que todos los campos requeridos en <c>frmDatosUsuario</c> estén completos
        /// y sean válidos antes de grabar el usuario.
        /// </summary>
        /// <returns><c>true</c> si los campos código cliente, CUIT, desc1, desc2 y beneficio son válidos.</returns>
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

        /// <summary>
        /// Lee los datos del usuario desde la tabla <c>USUARIO</c> de <c>cnnU</c>
        /// y los carga en los controles de <c>frmDatosUsuario</c>.
        /// </summary>
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

        /// <summary>
        /// Valida que el código de cliente exista en la tabla <c>CLIENTES</c> de la base del distribuidor.
        /// Si existe, carga en las variables <c>US_*</c> los datos del cliente.
        /// </summary>
        /// <param name="codi">Código numérico del cliente a validar.</param>
        /// <returns><c>true</c> si el código existe; <c>false</c> si no existe o hay error.</returns>
        /// <remarks>Actualiza <see cref="ExisteCodi_Cliente"/> como efecto secundario.</remarks>
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

        /// <summary>
        /// Valida la combinación código + CUIT contra la tabla <c>CLIENTES</c> de la base del distribuidor.
        /// Carga los descuentos, la condición de IVA y determina si el cliente está autorizado.
        /// </summary>
        /// <remarks>
        /// Efecto secundario clave: establece
        /// <see cref="ClienteAutorizado"/>, <see cref="ClienteRestringido"/>,
        /// <see cref="ClienteDescuento1"/> y <see cref="ClienteDescuento2"/>.
        /// Un cliente está <b>restringido</b> cuando CLDES1 = CLDES2 = 0
        /// (sin descuentos cargados por la distribuidora).
        /// </remarks>
        /// <param name="codi">Código del cliente.</param>
        /// <param name="cuit">CUIT del cliente (se parsea a double para la comparación).</param>
        /// <returns><c>true</c> si la combinación es válida; <c>false</c> si no existe o hay error.</returns>
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

        /// <summary>
        /// Sincroniza la tabla <c>FAMILIAS</c> de la base local del usuario (<c>cnnU</c>)
        /// con los datos de familias de la base del distribuidor (<c>cnn</c>).
        /// </summary>
        /// <remarks>
        /// Proceso: borra todos los registros de FAMILIAS en <c>cnnU</c> y los re-inserta
        /// desde la tabla FAMILIAS de <c>cnn</c>, copiando FACODI, FADESC y FAMIVA.
        /// Este método actualmente está comentado en <see cref="Inicializacion"/>.
        /// </remarks>
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

        /// <summary>
        /// Indica si el cliente tiene descuentos configurados en la distribuidora.
        /// </summary>
        /// <returns>
        /// <c>true</c> si <see cref="ClienteDescuento1"/> es distinto de cero;
        /// <c>false</c> cuando CLDES1 = 0 (el distribuidor aún no asignó descuentos).
        /// </returns>
        public static bool DescuentosConfigurados() => ClienteDescuento1 != 0;

        /// <summary>
        /// Sincroniza los descuentos del cliente desde la base del distribuidor hacia la base local.
        /// Lee <c>CLDES1</c> y <c>CLDES2</c> de la tabla <c>CLIENTES</c> (base del distribuidor, <c>cnn</c>)
        /// para el cliente identificado por <see cref="US_CODI"/> y los persiste en
        /// <c>US_DESCUENTO1</c> / <c>US_DESCUENTO2</c> de la tabla <c>USUARIO</c> en <c>cliente.dat</c>.
        /// </summary>
        /// <remarks>
        /// Debe llamarse después de migrar la base del distribuidor (<see cref="UpdateBase"/>)
        /// y <b>antes</b> de <see cref="Cargar_Usuario"/>, para que el recálculo de precios
        /// utilice los descuentos más recientes publicados por la distribuidora.
        /// Si el cliente no se encuentra o hay un error, se registra en el log y se continúa
        /// sin modificar los descuentos existentes.
        /// </remarks>
        public static void SincronizarDescuentosCliente()
        {
            try
            {
                // Leer CLDES1 y CLDES2 desde la base del distribuidor (recién actualizada por FTP)
                decimal cldes1 = 0;
                decimal cldes2 = 0;
                bool encontrado = false;

                ConectarDBArti();
                using (var cmdLeer = new OleDbCommand(
                    "SELECT CLDES1, CLDES2 FROM CLIENTES WHERE CLCODI = ?", cnn))
                {
                    cmdLeer.Parameters.AddWithValue("@codi", US_CODI);
                    using (var dr = cmdLeer.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            cldes1 = dr["CLDES1"] != DBNull.Value ? Convert.ToDecimal(dr["CLDES1"]) : 0;
                            cldes2 = dr["CLDES2"] != DBNull.Value ? Convert.ToDecimal(dr["CLDES2"]) : 0;
                            encontrado = true;
                        }
                    }
                }
                if (cnn.State == ConnectionState.Open) cnn.Close();

                if (!encontrado)
                {
                    GrabarError($" - {DateTime.Now:dd/MM/yyyy HH:mm:ss} 'SincronizarDescuentosCliente': cliente '{US_CODI}' no encontrado en CLIENTES.");
                    return;
                }

                // Actualizar US_DESCUENTO1 / US_DESCUENTO2 en la base local del usuario
                ConectarDBUsuario();
                using (var cmdUpd = new OleDbCommand(
                    "UPDATE USUARIO SET US_DESCUENTO1 = ?, US_DESCUENTO2 = ?", cnnU))
                {
                    cmdUpd.Parameters.AddWithValue("@desc1", cldes1);
                    cmdUpd.Parameters.AddWithValue("@desc2", cldes2);
                    cmdUpd.ExecuteNonQuery();
                }
                if (cnnU.State == ConnectionState.Open) cnnU.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                if (cnnU.State == ConnectionState.Open) cnnU.Close();
                GrabarError($" - {DateTime.Now:dd/MM/yyyy HH:mm:ss} 'SincronizarDescuentosCliente': {ex.Message}");
            }
        }

        /// <summary>
        /// Carga los descuentos del cliente (<c>US_DESCUENTO1</c>, <c>US_DESCUENTO2</c>)
        /// desde la tabla <c>USUARIO</c> de la base local y los muestra en <c>txtDescuento1</c>
        /// y <c>txtDescuento2</c> del formulario principal.
        /// </summary>
        /// <remarks>
        /// Se llama en <see cref="Inicializacion"/> y en <see cref="ActualizacionXFTP"/> (antes de
        /// recargar la grilla) para garantizar que los cálculos de precios usen los descuentos actualizados.
        /// Usa una conexión local independiente (<c>cnnLocal</c>) para evitar conflictos de estado.
        /// </remarks>
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

        /// <summary>
        /// Carga en la grilla todos los artículos que tienen precio de oferta (<c>ARTT &lt;&gt; 0</c>),
        /// mostrando el precio de oferta calculado (<c>arpre5</c>).
        /// </summary>
        public static async Task Funcion_Ofertas()
        {
            var da = new OleDbDataAdapter();
            try
            {
                ConectarDBArti();
                da = new OleDbDataAdapter("SELECT a.arterc as Codigo, a.ardesc as Descripcion, " +
                    "p.PRMARC as Marca," +
                    "IIF(r.rbdesc IS NOT NULL, r.rbdesc, '') as Montadora," +
                    //"l.lidesc as Linea," +
                    "a.armode as Linea," +
                    " a.arinte as Posición," +
                    //" a.ARMARC as Marca2," +
                    "a.arpre5 as Preciopublico, IIF(a.artt IS NOT NULL AND a.artt <> 0, a.artt, a.aroo) as aroo" +
                    ", a.arff as Stock, a.artt, IIF(a.arubic1 IS NOT NULL, a.arubic1, '') as arubic1, IIF(a.arobse IS NOT NULL, a.arobse, '') as arobse " +
                    " from (listas as a INNER JOIN proveedo AS p ON p.prcodi = a.arprov) " +
                    "LEFT JOIN rubros AS r ON a.ARRUB2 = r.RBCODI " +
                    " WHERE a.ARTT <> 0 and a.arprim<>'000' order by a.arterc", cnn);
                var ds = new DataSet();
                da.Fill(ds);
                _frmPrincipal.DataGridView1.DataSource = ds.Tables[0];
                _frmPrincipal.DataGridView1.Refresh();
                Definir_Grilla();
                await SeleccionarGrilla();
                if (cnn.State == ConnectionState.Open) cnn.Close();
            }
            catch (Exception ex)
            {
                if (cnn.State == ConnectionState.Open) cnn.Close();
                throw new Exception("Error en la Funcion 'Funcion_Ofertas' " + ex.Message);
                da.Dispose();
            }
        }

        /// <summary>
        /// Limpia todos los campos de búsqueda del formulario principal y, si no es modo interno,
        /// relanza la consulta base y restaura el estado visual de la grilla.
        /// </summary>
        /// <param name="p1">Si es 1, no limpia los combos <c>cmbArmadora</c> / <c>cmbLineas</c>
        /// ni relanza <see cref="consultas_Relacionales"/>. Usar 0 para limpiar todo.</param>
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
                if (p1 != 1) _frmPrincipal.cmbArmadora.SelectedIndex = -1;
                if (p1 != 1) _frmPrincipal.cmbLineas.SelectedIndex = -1;
                if (p1 != 1) _frmPrincipal.cmbRubros.SelectedIndex = -1;
                
                // Cuando se limpia rubros, actualizar marcas para mostrar todas
                if (p1 != 1) 
                {
                    CargarComboMarcasPorRubro(""); // Cargar todas las marcas
                    _frmPrincipal.cmbMarcas.SelectedIndex = -1; // Sin selección
                }
                
                // Resetear filtros de NOVEDADES y OFERTAS
                if (p1 != 1 && _frmPrincipal.ckbOfertas != null)
                    _frmPrincipal.ckbOfertas.Checked = false;
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

        /// <summary>
        /// [<b>Obsoleto</b> — usa <c>cnn</c> en lugar de <c>cnnU</c>] Elimina un ítem del carrito por su número de ítem.
        /// Para el flujo actual usar <see cref="ModuloPreparac.EliminarItemCarro"/>.
        /// </summary>
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

        /// <summary>
        /// [<b>Obsoleto</b> — usa <c>cnn</c> en lugar de <c>cnnU</c>] Elimina todos los ítems del carrito del usuario actual.
        /// Para el flujo actual usar <see cref="ModuloPreparac.EliminarItemsCarro"/>.
        /// </summary>
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

        /// <summary>
        /// [<b>Obsoleto</b> — usa <c>cnn</c> en lugar de <c>cnnU</c>] Verifica si un artículo ya existe en PREPARAC.
        /// Para el flujo actual usar <see cref="ModuloPreparac.ExisteARTICULOPedido"/>.
        /// </summary>
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

        /// <summary>
        /// Wrapper asíncrono que delega en <see cref="ModuloPreparac.GrabarArticuloaPreparacAsync"/>
        /// usando el WHERE configurado en <see cref="ModuloPreparac.WHERE_PREPARAC"/>.
        /// </summary>
        /// <param name="accion">"INSERT" o "UPDATE".</param>
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

        /// <summary>
        /// Configura <c>WHERE_PREPARAC</c> y <c>CANT_PREPARAC</c> en <see cref="ModuloPreparac"/>
        /// para una operación posterior de agregar al carrito.
        /// </summary>
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

        /// <summary>
        /// [<b>Obsoleto</b> — usa SQL dinámico concatenado, sin parámetros]
        /// Ejecuta un INSERT o UPDATE directo sobre PREPARAC usando las variables de <see cref="ModuloPreparac"/>.
        /// Para el flujo actual usar <see cref="ModuloPreparac.GrabarArticuloaPreparacAsync"/>.
        /// </summary>
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
        /// <summary>
        /// Descarga un archivo del servidor FTP y lo guarda en la ruta local especificada.
        /// Lee el stream por bloques de 3071 bytes.
        /// </summary>
        /// <param name="Filename">Nombre del archivo en el servidor FTP.</param>
        /// <param name="NewFileName">Ruta local completa donde guardar el archivo descargado.</param>
        /// <param name="UserNameFtp">Usuario FTP.</param>
        /// <param name="PassFTP">Contraseña FTP.</param>
        /// <param name="URI">URL base del servidor FTP.</param>
        /// <returns><c>true</c> si la descarga fue exitosa; <c>false</c> si hubo error.</returns>
        /// <summary>
        /// Descarga un archivo desde FTP de forma asíncrona sin bloquear la UI.
        /// </summary>
        public static async Task<bool> DescargaArchivoAsync(string Filename, string NewFileName, string UserNameFtp, string PassFTP, string URI)
        {
            try
            {
                FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create(URI + "/" + Filename);
                ftpRequest.Credentials = new NetworkCredential(UserNameFtp, PassFTP);
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                ftpRequest.UseBinary = true;
                ftpRequest.KeepAlive = false;

                using (FtpWebResponse response = (FtpWebResponse)await ftpRequest.GetResponseAsync().ConfigureAwait(false))
                using (Stream responseStream = response.GetResponseStream())
                using (FileStream fileStream = new FileStream(NewFileName, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                {
                    byte[] buffer = new byte[8192];
                    int bytesRead;
                    while ((bytesRead = await responseStream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
                    {
                        await fileStream.WriteAsync(buffer, 0, bytesRead).ConfigureAwait(false);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                GrabarError($" - {DateTime.Now.ToLongDateString()} => {DateTime.Now.ToString("HH:mm:ss")} 'DescargaArchivoAsync': {ex.Message}");
                return false;
            }
        }

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
        /// <summary>
        /// Retorna la fecha máxima de la tabla <c>ACTUALIZACIONES</c> de la base del distribuidor.
        /// Si no hay registros, retorna la fecha actual como fallback.
        /// </summary>
        /// <returns>Fecha de actualización formateada como <c>"dd/MM/yyyy"</c>.</returns>
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

        /// <summary>
        /// Verifica si el artículo seleccionado en la grilla tiene equivalencias disponibles
        /// (campo <c>AROBSE</c> no vacío) y habilita/deshabilita el botón de equivalencias.
        /// Almacena el valor en <see cref="NumEqui"/> para uso posterior.
        /// </summary>
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


        /// <summary>
        /// Recalcula masivamente los precios de costo (<c>arpre0</c>, <c>arpre1</c>) en la tabla LISTAS
        /// aplicando la cadena completa: desc1 → desc2 → descuento de familia (<c>FADES1</c>) → IVA de familia (<c>FAIVA</c>).
        /// </summary>
        /// <remarks>
        /// Algoritmo por familia (itera todas las FAMILIAS con artículos en LISTAS):
        /// <list type="bullet">
        /// <item>Si desc2 &gt; 0: <c>arpre1 = arpre1 − FADES1 + FAIVA</c>.</item>
        /// <item>Si solo desc1 &gt; 0: <c>arpre0 = arpre0 − FADES1 + FAIVA</c>; luego <c>arpre1 = arpre0</c>.</item>
        /// <item>Sin descuentos: <c>arpre1 = arpre1 − FADES1 + FAIVA</c> directo.</item>
        /// </list>
        /// Los valores de FADES1 y FAIVA se leen de la tabla <c>FAMILIAS</c> en <c>cnnU</c>.
        /// </remarks>
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

                    _frmPrincipal.txtDescuento1.Text = US_DESCUENTO1.ToString();
                    _frmPrincipal.txtDescuento2.Text = US_DESCUENTO2.ToString();

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

                    // Se itera por FACODI (no FADESC) porque una misma descripción de familia puede
                    // repetirse con distinto proveedor (FAPROV); FACODI es la clave real y única.
                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FACODI FROM [familias] F ", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        ConectarDBUsuario();
                        using var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA, FADES1 FROM FAMILIAS WHERE FACODI = ?", cnnU);
                        cmUsu.Parameters.Add("@p1", OleDbType.VarWChar);

                        while (drFam.Read())
                        {
                            string facodi = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(facodi))
                                continue;

                            // Releemos descuentos para decidir dónde aplicar IVA
                            decimal d1 = 0, d2 = 0;
                            decimal.TryParse(_frmPrincipal.txtDescuento1.Text, out d1);
                            decimal.TryParse(_frmPrincipal.txtDescuento2.Text, out d2);

                            decimal? faiva = null;
                            decimal? fades1 = null;
                            cmUsu.Parameters[0].Value = facodi;
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
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 - (L.arpre1*" + des + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre0 = FORMAT(L.arpre0 - (L.arpre0*" + des + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = L.arpre0 WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 - (L.arpre1*" + des + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
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
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 + (L.arpre1*" + iva + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    // Solo desc1: IVA sobre arpre0 y luego reflejar en arpre1
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre0 = FORMAT(L.arpre0 + (L.arpre0*" + iva + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = L.arpre0 WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    // Sin descuentos: IVA directo sobre arpre1
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre1 = FORMAT(L.arpre1 + (L.arpre1*" + iva + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
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

        /// <summary>
        /// Equivalente a <see cref="Update_PrecioCosto"/> pero para los precios de oferta:
        /// recalcula <c>arpre3</c> y <c>arpre4</c> a partir de <c>artt</c> (precio de oferta base).
        /// </summary>
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

                    // Se itera por FACODI (no FADESC) porque una misma descripción de familia puede
                    // repetirse con distinto proveedor (FAPROV); FACODI es la clave real y única.
                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FACODI FROM [familias] F INNER JOIN [listas] L ON L.ARFLIA = F.FACODI", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        ConectarDBUsuario();
                        using var cmUsu = new OleDbCommand("SELECT TOP 1 FAIVA, FADES1 FROM FAMILIAS WHERE FACODI = ?", cnnU);
                        cmUsu.Parameters.Add("@p1", OleDbType.VarWChar);

                        while (drFam.Read())
                        {
                            string facodi = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(facodi))
                                continue;

                            decimal d1 = 0, d2 = 0;
                            decimal.TryParse(_frmPrincipal.txtDescuento1.Text, out d1);
                            decimal.TryParse(_frmPrincipal.txtDescuento2.Text, out d2);

                            decimal? faiva = null;
                            decimal? fades1 = null;
                            cmUsu.Parameters[0].Value = facodi;
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
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 - (L.arpre4*" + des + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre3 = FORMAT(L.arpre3 - (L.arpre3*" + des + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = L.arpre3 WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 - (L.arpre4*" + des + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
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
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 + (L.arpre4*" + iva + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else if (d1 != 0)
                                {
                                    // Solo desc1: IVA sobre arpre3 y luego reflejar en arpre4
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre3 = FORMAT(L.arpre3 + (L.arpre3*" + iva + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();

                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = L.arpre3 WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
                                    cmaccessUpdate.ExecuteNonQuery();
                                }
                                else
                                {
                                    // Sin descuentos: IVA directo sobre arpre4
                                    cmaccessUpdate.Parameters.Clear();
                                    cmaccessUpdate.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre4 = FORMAT(L.arpre4 + (L.arpre4*" + iva + "/100),'Fixed') WHERE F.FACODI = ?";
                                    cmaccessUpdate.Parameters.AddWithValue("@p1", facodi);
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

        /// <summary>
        /// Calcula el precio sugerido de venta al público (<c>arpre2</c>) para todos los artículos
        /// aplicando el beneficio familiar (<c>FABEN1</c>) sobre el precio de costo calculado (<c>arpre1</c>).
        /// </summary>
        /// <remarks>
        /// Fórmula: <c>arpre2 = arpre1 + (arpre1 × FABEN1 / 100)</c>.
        /// Si la familia no tiene <c>FABEN1</c>, copia <c>arpre1</c> a <c>arpre2</c> directamente.
        /// </remarks>
        public static void Update_PrecioSug()
        {
            using (OleDbCommand cmaccess = new OleDbCommand())
            {
                try
                {
                    ConectarDBArti();
                    cmaccess.Connection = cnn;
                    cmaccess.CommandType = CommandType.Text;

                    // Se itera por FACODI (no FADESC) porque una misma descripción de familia puede
                    // repetirse con distinto proveedor (FAPROV); FACODI es la clave real y única.
                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FACODI FROM [familias] F INNER JOIN [listas] L ON L.ARFLIA = F.FACODI", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        while (drFam.Read())
                        {
                            string facodi = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(facodi))
                                continue;

                            decimal? faben1 = null;
                            using (var cmUsu = new OleDbCommand("SELECT TOP 1 FABEN1 FROM FAMILIAS WHERE FACODI = ?", cnnU))
                            {
                                ConectarDBUsuario();
                                cmUsu.Parameters.AddWithValue("@p1", facodi);
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
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre2 = FORMAT(L.arpre1 + (L.arpre1*" + ben + "/100),'Fixed') WHERE F.FACODI = ?";
                                cmaccess.Parameters.AddWithValue("@p1", facodi);
                                cmaccess.ExecuteNonQuery();
                            }
                            else
                            {
                                cmaccess.Parameters.Clear();
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre2 = FORMAT(L.arpre1,'Fixed') WHERE F.FACODI = ?";
                                cmaccess.Parameters.AddWithValue("@p1", facodi);
                                try { GrabarError($"Familia sin FABEN1 (o 0) al sugerir precio: FACODI='{facodi}'"); } catch {}
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

        /// <summary>
        /// Equivalente a <see cref="Update_PrecioSug"/> para ofertas: calcula <c>arpre5</c>
        /// aplicando <c>FABEN1</c> sobre <c>arpre4</c>.
        /// </summary>
        public static void Update_PrecioSug_Ofertas()
        {
            using (OleDbCommand cmaccess = new OleDbCommand())
            {
                try
                {
                    ConectarDBArti();
                    cmaccess.Connection = cnn;
                    cmaccess.CommandType = CommandType.Text;

                    // Se itera por FACODI (no FADESC) porque una misma descripción de familia puede
                    // repetirse con distinto proveedor (FAPROV); FACODI es la clave real y única.
                    using (var cmdFam = new OleDbCommand("SELECT DISTINCT F.FACODI FROM [familias] F INNER JOIN [listas] L ON L.ARFLIA = F.FACODI", cnn))
                    using (var drFam = cmdFam.ExecuteReader())
                    {
                        while (drFam.Read())
                        {
                            string facodi = drFam[0]?.ToString() ?? string.Empty;
                            if (string.IsNullOrEmpty(facodi))
                                continue;

                            decimal? faben1 = null;
                            using (var cmUsu = new OleDbCommand("SELECT TOP 1 FABEN1 FROM FAMILIAS WHERE FACODI = ?", cnnU))
                            {
                                ConectarDBUsuario();
                                cmUsu.Parameters.AddWithValue("@p1", facodi);
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
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre5 = FORMAT(L.arpre4 + (L.arpre4*" + ben + "/100),'Fixed') WHERE F.FACODI = ?";
                                cmaccess.Parameters.AddWithValue("@p1", facodi);
                                cmaccess.ExecuteNonQuery();
                            }
                            else
                            {
                                cmaccess.Parameters.Clear();
                                cmaccess.CommandText = "UPDATE listas AS L INNER JOIN familias AS F ON L.ARFLIA = F.FACODI SET L.arpre5 = FORMAT(L.arpre4,'Fixed') WHERE F.FACODI = ?";
                                cmaccess.Parameters.AddWithValue("@p1", facodi);
                                try { GrabarError($"Familia sin FABEN1 (o 0) al sugerir precio ofertas: FACODI='{facodi}'"); } catch {}
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

        /// <summary>
        /// Ejecuta una consulta SELECT sobre la base del distribuidor y muestra el rubro,
        /// origen y observación del artículo en las etiquetas del formulario principal.
        /// </summary>
        /// <param name="StringSelect">SELECT que retorna RBDESC, ARORIG, ARUBIC1 y AROBSE.</param>
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
                        Environment.NewLine + "Equivalencia: " + dr["AROBSE"].ToString();
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

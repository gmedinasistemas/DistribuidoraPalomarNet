using ServiceReferenceWSPalomar;
using System;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    /// <summary>
    /// Módulo estático que gestiona el carrito de compras (tabla <c>PREPARAC</c>) en la base local del usuario.
    /// </summary>
    /// <remarks>
    /// Centraliza las variables del registro PREPARAC y todas las operaciones sobre él:
    /// <para><b>Agregar/Actualizar artículos:</b> <see cref="BuscarArticuloParaPreparacAsync"/> →
    /// <see cref="GrabarArticuloaPreparacAsync"/> construyen y persisten el registro.</para>
    /// <para><b>Modificar cantidad:</b> <see cref="Update_item"/> recalcula precios y actualiza el ítem.</para>
    /// <para><b>Eliminar ítems:</b> <see cref="EliminarItemCarro"/> (uno) / <see cref="EliminarItemsCarro"/> (todos).</para>
    /// <para><b>Envío al distribuidor:</b> <see cref="ProcesoWS"/> obtiene número de pedido del WS
    /// y llama a <see cref="InsertWSPreparac"/> para enviar cada fila al web service WSPalomar.</para>
    /// <para>Opera sobre <b>Module.cnnU</b> (base local), excepto <see cref="BuscarArticuloParaPreparacAsync"/>
    /// que lee de <b>Module.cnn</b> (base del distribuidor).</para>
    /// </remarks>
    public static class ModuloPreparac
    {
        #region Variables Preparac
        // --- Temporales de operación ---
        // Codi: código compuesto para display ("PRIMERA - TERCERA"), no se graba en PREPARAC.
        // Precio: precio de costo sin IVA del artículo en procesamiento (Module.precio_costo_sin_iva).
        // Cant: cantidad seleccionada; CantExistente: cantidad ya en PREPARAC si el artículo existe.
        // SUBT, Iva: valores intermedios de cálculo durante la grabación.
        public static string Codi = "";
        public static decimal Precio = 0;
        public static decimal Cant; // As Integer en VB
        public static decimal CantExistente=0;
        public static decimal SUBT, Iva;
        // --- Descripción y observaciones del artículo ---
        // _PPADESC (máx. 40 car.): descripción del artículo → columna PPADESC.
        // _PPAOBSE (máx. 255 car.): observaciones (campo AROBSE de LISTAS) → columna PPAOBSE.
        public static string _PPADESC = ""; // 40
        public static string _PPAOBSE = ""; // 255
        // --- Precios del artículo ---
        // _PPAORIG:        precio original con IVA (AROO + IVA de familia), formato 'Currency' de Access.
        // _PPAPREC:        precio de lista base (AROO sin IVA); también asignado a _PPABASE.
        // _PPABASE:        igual a _PPAPREC; base para cálculos posteriores en el WS.
        // _PPACIVA:        precio_final × cantidad (monto total IVA incluido por la línea).
        // _PPAPRECUNITARIO:precio unitario final con descuentos y beneficio aplicados (Module.precio_final).
        // _PPAPRECCOSTO:   precio de costo sin IVA redondeado (Module.precio_costo_sin_iva).
        // _FAMIVA:         porcentaje de IVA de la familia del artículo (leído de FAMILIAS.FAIVA).
        public static decimal _PPAORIG = 0;
        public static decimal _PPAPREC = 0;
        public static decimal _PPABASE = 0;
        public static decimal _PPACIVA = 0;
        public static decimal _PPAPRECUNITARIO = 0;
        public static decimal _PPAPRECCOSTO = 0;
        public static decimal _FAMIVA = 0;

        // --- Subtotales sin IVA ---
        // _SUBTSIVA: subtotal sin IVA = precio_costo_sin_iva × cantidad → columna SUBTSIVA.
        // _TOTALSIVA: total acumulado sin IVA (no utilizado en el flujo principal actual).
        public static decimal _SUBTSIVA = 0;
        public static decimal _TOTALSIVA = 0;

        // --- Identificadores y códigos del artículo ---
        // _PPABARR:  código de barras del artículo (LISTAS.ARBARR).
        // _PPACODx:  partes del código compuesto: COD1(3)=ARPRIM, COD2(4)=ARSEGU, COD3(10)=ARTERC, COD4(4)=ARCUAR.
        //            _PPACOD3 es el identificador principal usado como clave en UPDATE/EXISTS.
        // _PPACODI (3): código interno del artículo (LISTAS.ARCODI).
        public static decimal _PPABARR = 0;
        public static string _PPACOD1 = ""; // 3
        public static string _PPACOD2 = ""; // 4
        public static string _PPACOD3 = ""; // 10
        public static string _PPACOD4 = ""; // 4
        public static string _PPACODI = ""; // 3
        // --- Encabezado del pedido (datos del cliente y condiciones) ---
        // _PPANUME: número de pedido; inicia en 0 y es asignado por el WS en ProcesoWS → TraerNPedido().
        // _PPATPRE (1): tipo de precio. _PPAFECH: fecha del pedido (inicializada a la fecha actual).
        // _PPARAZO (30): razón social del cliente (Module.US_CLRAZO).
        // _PPASITI: condición IVA del cliente (Module.US_CLCIVA).
        // _PPAQOT (1): flag de cotización. _PPAPPA: pago a cuenta. _PPADESK: descuento aplicado.
        // _PPATARJ (25): tarjeta. _PPACUOT: cuotas. _PPAVEND: código vendedor. _PPAFLIA: familia del artículo.
        public static decimal _PPANUME = 0;
        public static string _PPATPRE = ""; // 1
        public static string _PPAFECH = DateTime.Now.ToShortDateString();
        public static string _PPARAZO = ""; // 30
        public static decimal _PPASITI = 0;
        public static string _PPAQOT = ""; // 1
        public static decimal _PPAPPA = 0;
        public static decimal _PPADESK = 0;
        public static string _PPATARJ = ""; // 25
        public static decimal _PPACUOT = 0;
        public static decimal _PPAVEND = 0;
        public static decimal _PPAFLIA = 0;
        // --- Proveedor, origen e IVA ---
        // _PPAIMPO (1): indicador de origen/importación (LISTAS.ARORIG).
        // _PPAPROV: código de proveedor (LISTAS.ARPROV). _PPAAUTO: autonumérico; por defecto 1 (Access asigna PPAAUTN).
        // _PPAARFF (10): campo auxiliar vacío en el flujo actual.
        // _PPAFIVA: porcentaje de IVA a aplicar; si la familia retorna 0 se asume 21% (estándar Argentina).
        public static string _PPAIMPO = ""; // 1
        public static decimal _PPAPROV = 0;
        public static decimal _PPAAUTO = 1;
        public static string _PPAARFF = ""; // 10
        public static decimal _PPAFIVA = 0;
        // --- Operador, empresa y logística de despacho ---
        // _PPAUSER: código de usuario interno. _PPACOMP (30): nombre de empresa/compañía.
        // _PPABULT: bultos. _PPAKILO: peso en kilos. _PPASEGU: seguro de envío.
        // _PPAREMF: flete del remito. _PPAREMN: número de remito. _PPAPIEZ: piezas por unidad (por defecto 1).
        public static decimal _PPAUSER = 0;
        public static string _PPACOMP = ""; // 30
        public static decimal _PPABULT = 0;
        public static decimal _PPAKILO = 0;
        public static decimal _PPASEGU = 0;
        public static decimal _PPAREMF = 0;
        public static decimal _PPAREMN = 0;
        public static decimal _PPAPIEZ = 1;
        // _PPAMARC (15): código de marca del artículo (proveedo.PRMARC, columna PPAMARC).
        // _PRMARC: descripción de marca para el PDF; campo de "fantasía", no tiene columna propia en PREPARAC.
        public static string _PPAMARC = ""; // 15
        public static string _PRMARC = ""; // Campo de fantasía, solo para mostrar en el pdf
        // --- Porcentajes adicionales, fechas y unidades de entrega ---
        // _PPAPOR1, _PPAPOR2: porcentajes adicionales de ajuste.
        // _PPACONT: contador de ítems. _PPAFAPL: fecha de aplicación de precio. _PPAFORI: fecha de origen.
        // _PPAUNID: unidades (por defecto 1). _PPADEPO: depósito de origen. _PPADIRE: código de dirección.
        // _PPACPAG: condición de pago del cliente (Module.US_CLCPAG).
        public static decimal _PPAPOR1 = 0;
        public static decimal _PPAPOR2 = 0;
        public static decimal _PPACONT = 0;
        public static string _PPAFAPL = DateTime.Now.ToShortDateString();
        public static string _PPAFORI = DateTime.Now.ToShortDateString();
        public static decimal _PPAUNID = 1;
        public static decimal _PPADEPO = 0;
        public static decimal _PPADIRE = 0;
        public static decimal _PPACPAG = 0;

        // --- Variables de control de operación ---
        // WHERE_PREPARAC: cláusula WHERE dinámica construida externamente antes de llamar a BuscarArticuloParaPreparacAsync.
        //                 Ejemplo: " WHERE A.ARCODI='123' AND A.ARFLIA=5 AND A.ARPROV=2"
        // CANT_PREPARAC:  cantidad a insertar/actualizar; también se asigna externamente antes de la operación.
        public static string WHERE_PREPARAC = "";
        public static decimal CANT_PREPARAC = 0;
        #endregion

        /// <summary>
        /// Limpia un string de moneda (puede contener prefijos como XDR, $, separador de miles '.',
        /// separador decimal ',') y lo convierte a <c>decimal</c> usando <c>InvariantCulture</c>.
        /// </summary>
        /// <param name="valor">String crudo del campo monetario.</param>
        /// <returns>Decimal parseado o 0 si no es válido.</returns>
        private static decimal ParseMoneda(object valor)
        {
            if (valor == null || string.IsNullOrWhiteSpace(valor.ToString()))
                return 0;

            string limpio = valor.ToString().Trim();

            // Eliminar todo carácter que no sea dígito, coma o punto
            limpio = Regex.Replace(limpio, @"[^0-9.,]", "");

            if (string.IsNullOrEmpty(limpio))
                return 0;

            // Detectar el separador decimal: es el ÚLTIMO punto o coma en el string.
            // Los anteriores son separadores de miles y se eliminan.
            int lastDot = limpio.LastIndexOf('.');
            int lastComma = limpio.LastIndexOf(',');
            int lastSep = Math.Max(lastDot, lastComma);

            if (lastSep > 0)
            {
                string enteros = limpio.Substring(0, lastSep).Replace(".", "").Replace(",", "");
                string decimales = limpio.Substring(lastSep + 1);
                limpio = enteros + "." + decimales;
            }
            else if (lastSep == 0)
            {
                // Empieza con separador (ej ".50") — dejarlo como está para TryParse
                limpio = "0" + limpio.Replace(",", ".");
            }

            if (decimal.TryParse(limpio, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal resultado))
                return resultado;

            return 0;
        }

        /// <summary>
        /// Procesa las filas del <c>DataSet</c> recibido y graba o actualiza el artículo en la tabla
        /// <c>PREPARAC</c> de la base local, según la acción indicada.
        /// </summary>
        /// <remarks>
        /// Flujo interno:
        /// <list type="number">
        /// <item>Extrae el precio de costo sin IVA desde <c>Module.precio_costo_sin_iva</c>.</item>
        /// <item>Parsea los campos de precio del DataSet (formato 'Currency' de Access: <c>"$1.234"</c>)
        /// eliminando <c>$</c> y <c>.</c> (separador de miles) antes de convertir a decimal.</item>
        /// <item>Consulta el IVA de la familia del artículo (<see cref="TraerFAMIVAAsync"/>);
        /// si retorna 0 asume 21% estándar.</item>
        /// <item>Construye y ejecuta INSERT (nuevo artículo) o UPDATE (artículo ya existente) sobre PREPARAC.</item>
        /// <item>Refresca el carrito llamando a <c>Module.funcion_inicial_carrito()</c>.</item>
        /// </list>
        /// </remarks>
        /// <param name="DTS">DataSet con una tabla que contiene los campos del artículo seleccionado.</param>
        /// <param name="Accion">"INSERT" para nuevo artículo; cualquier otro valor ejecuta UPDATE.</param>
        public static async Task GrabarArticuloaPreparacAsync(DataSet DTS, string Accion)
        {
            string strQuery = "";
            try
            {
                foreach (DataRow fila in DTS.Tables[0].Rows)
                {
                    string Codi = $"{fila["PRIMERA"].ToString().Trim()} - {fila["TERCERA"].ToString().Trim()}";
                    // Precio de costo sin IVA: viene de Module.precio_costo_sin_iva (calculado en AccionGrilla/Calculadora)
                    decimal Precio = Math.Round(Module.precio_costo_sin_iva, 2, MidpointRounding.ToEven);
                    _PPAPRECCOSTO = Math.Round(Module.precio_costo_sin_iva, 2, MidpointRounding.ToEven);

                    if (Precio != 0)
                    {
                        decimal SUBT = 0, Iva = 0;
                        string Cant = ModuloPreparac.CANT_PREPARAC.ToString().Replace(".", ",");

                        if (Cant != "0")
                        {
                            // Calcular subtotal base: precio_costo × cantidad
                            if (decimal.TryParse(Cant, out decimal cantidadDecimal))
                            {
                                SUBT = Math.Round(Precio * cantidadDecimal, 2, MidpointRounding.ToEven);
                            }

                            _PPADESC = !string.IsNullOrEmpty(fila["DESCRIPCION"].ToString()) ? fila["DESCRIPCION"].ToString() : "";
                            _PPAOBSE = !string.IsNullOrEmpty(fila["OBSERVACION"].ToString()) ? fila["OBSERVACION"].ToString() : "";
                            //_PPAORIG = !string.IsNullOrEmpty(fila["PPAORIG"].ToString())
                            //? Convert.ToDecimal(fila["PPAORIG"].ToString().Replace("$", "").Replace(".", "").Replace(",", "."))
                            //: 0;

                            _PPAORIG = ParseMoneda(fila["PPAORIG"]);
                            //_PPAORIG = !string.IsNullOrEmpty(fila["PPAORIG"].ToString()) ? Convert.ToDecimal(fila["PPAORIG"]) : 0;

                            _PPAPREC = ParseMoneda(fila["PPAPREC"]);
                            _PPABASE = ParseMoneda(fila["PPAPREC"]);

                            //_PPAPREC = Convert.ToDecimal(fila["PPAPREC"]);
                            //_PPABASE = Convert.ToDecimal(fila["PPAPREC"]);
                            _PPACIVA = ParseMoneda(fila["PPACIVA"]);
                            _PPAPRECUNITARIO = ParseMoneda(fila["PPAPRECUNITARIO"]);

                            //_PPACIVA = !string.IsNullOrEmpty(fila["PPACIVA"].ToString()) ? Convert.ToDecimal(fila["PPACIVA"]) : 0;
                            //_PPAPRECUNITARIO = !string.IsNullOrEmpty(fila["PPAPRECUNITARIO"].ToString()) ? Convert.ToDecimal(fila["PPAPRECUNITARIO"]) : 0;

                            _PPABARR = !string.IsNullOrEmpty(fila["ARBARR"].ToString()) ? Convert.ToDecimal(fila["ARBARR"]) : 0;
                            _PPACODI = !string.IsNullOrEmpty(fila["CODIGO"].ToString()) ? fila["CODIGO"].ToString() : "";
                            _PPAFLIA = !string.IsNullOrEmpty(fila["FAMILIA"].ToString()) ? Convert.ToDecimal(fila["FAMILIA"]) : 0;
                            _PPAIMPO = !string.IsNullOrEmpty(fila["ORIGEN"].ToString()) ? fila["ORIGEN"].ToString() : "";
                            _PPAPROV = !string.IsNullOrEmpty(fila["PROVEEDOR"].ToString()) ? Convert.ToDecimal(fila["PROVEEDOR"]) : 0;
                            _PPAARFF = "";
                            _PPAMARC = !string.IsNullOrEmpty(fila["MARCA"].ToString()) ? fila["MARCA"].ToString() : "";
                            _PRMARC = !string.IsNullOrEmpty(fila["PRMARC"].ToString()) ? fila["PRMARC"].ToString() : "";
                            _PPADEPO = !string.IsNullOrEmpty(fila["ARDEPO"].ToString()) ? Convert.ToDecimal(fila["ARDEPO"]) : 0;
                            _PPACOD1 = !string.IsNullOrEmpty(fila["PRIMERA"].ToString()) ? fila["PRIMERA"].ToString() : "";
                            _PPACOD2 = !string.IsNullOrEmpty(fila["SEGUNDA"].ToString()) ? fila["SEGUNDA"].ToString() : "";
                            _PPACOD3 = !string.IsNullOrEmpty(fila["TERCERA"].ToString()) ? fila["TERCERA"].ToString() : "";
                            _PPACOD4 = !string.IsNullOrEmpty(fila["CUARTA"].ToString()) ? fila["CUARTA"].ToString() : "";
                            // Datos del cliente tomados de las variables globales de Module (cargadas en ValidarCodiCliente)
                            _PPARAZO = Module.US_CLRAZO;
                            _PPASITI = Module.US_CLCIVA;
                            _PPAVEND = Module.US_CLVEND;

                            string SelectFAMIVA = $"SELECT familias.FAIVA FROM familias INNER JOIN PREPARAC ON familias.FACODI = PREPARAC.PPAFLIA WHERE PREPARAC.PPACODI='{_PPACODI}'";
                            // IVA de la familia: si no está configurado en FAMILIAS se asume 21% (alícuota general Argentina)
                            _PPAFIVA = await TraerFAMIVAAsync(SelectFAMIVA);
                            if (_PPAFIVA == 0) _PPAFIVA = 21;

                            _SUBTSIVA = Math.Round(Module.precio_costo_sin_iva * cantidadDecimal, 2, MidpointRounding.ToEven);
                            _PPACPAG = Module.US_CLCPAG;
                            _PPANUME = 0; // Se asigna el número real de pedido luego en ProcesoWS → TraerNPedido()

                            // Construir la sentencia SQL según si el artículo es nuevo en el carrito o ya existe
                            if (Accion == "INSERT")
                            {
                                strQuery = $"INSERT INTO PREPARAC(" +
                                    $"PPANUME,PPAESTA,PPACLIE,PPABARR,PPACOD1,PPACOD2,PPACOD3,PPACOD4," +
                                    $"PPADESC,PPACANT,PPAFECH,PPATPRE,PPAPREC,PPACIVA,PPARAZO,PPASITI,PPAQOT,PPAPPA,PPADESK," +
                                    $"PPAORIG,PPACODI,PPATARJ,PPACUOT,PPAVEND,PPAFLIA,PPAIMPO,PPAPROV,PPAAUTO,PPAARFF,PPAFIVA,PPAOBSE," +
                                    $"PPAUSER,PPACOMP,PPABULT,PPAKILO,PPASEGU,PPAREMF,PPAREMN,PPAPIEZ,PPAMARC,PPAPOR1,PPAPOR2,PPABASE," +
                                    $"PPACONT,PPAFAPL,PPAFORI,PPAUNID,PPADEPO,PPADIRE,PPACPAG,SUBTSIVA,PRMARC,PPAPRECUNITARIO,PPAPRECCOSTO)" +
                                    $" VALUES(" +
                                    $"{_PPANUME},1,{Module.US_CODI},{_PPABARR},'{_PPACOD1}','{_PPACOD2}','{_PPACOD3}','{_PPACOD4}'," +
                                    $"'{_PPADESC}','{Cant}','{_PPAFECH}','{_PPATPRE}','{_PPAPREC}','{_PPACIVA}','{_PPARAZO}'," +
                                    $"'{_PPASITI}','{_PPAQOT}',{_PPAPPA},'{_PPADESK}','{_PPAORIG}','{_PPACODI}','{_PPATARJ}'," +
                                    $"{_PPACUOT},{_PPAVEND},'{_PPAFLIA}','{_PPAIMPO}',{_PPAPROV},{_PPAAUTO},'{_PPAARFF}'," +
                                    $"{_PPAFIVA},'{_PPAOBSE}',{_PPAUSER},'{_PPACOMP}',{_PPABULT},{_PPAKILO},{_PPASEGU}," +
                                    $"{_PPAREMF},{_PPAREMN},{_PPAPIEZ},'{_PPAMARC}',{_PPAPOR1},{_PPAPOR2},'{_PPABASE}'," +
                                    $"{_PPACONT},'{_PPAFAPL}','{_PPAFORI}',{_PPAUNID},{_PPADEPO},{_PPADIRE},{_PPACPAG}," +
                                    $"'{_SUBTSIVA}','{_PRMARC}','{_PPAPRECUNITARIO}','{_PPAPRECCOSTO}')";
                            }
                            else
                            {
                                _SUBTSIVA = Math.Round(Module.precio_costo_sin_iva * cantidadDecimal, 2, MidpointRounding.ToEven);

                                strQuery = $"UPDATE PREPARAC SET" +
                                    $" PPAPREC='{_PPAPREC}', PPACANT='{Cant}', PPACIVA='{_PPACIVA}'," +
                                    $" SUBTSIVA='{_SUBTSIVA}', PPAPRECUNITARIO='{_PPAPRECUNITARIO}'," +
                                    $" PPAPRECCOSTO='{_PPAPRECCOSTO}'" +
                                    $" WHERE PPACOD3='{_PPACOD3}' AND PPAFLIA={_PPAFLIA} AND PPAPROV={_PPAPROV}";
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Artículo sin precio. Controlar en el Sistema", "No lo puede cargar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // Eliminar el espacio posterior al símbolo $ que puede quedar de Format('Currency') antes de ejecutar
                if (!string.IsNullOrEmpty(strQuery))
                {
                    EjecutarQueryBDUsu(strQuery.Replace("$ ", ""));
                }

                Module.funcion_inicial_carrito();
            }
            catch (Exception ex)
            {
                Module.funcion_inicial_carrito();
                throw new Exception($", dentro de 'GrabarArticuloaPreparac': {ex.Message}");
            }
        }
        /// <summary>
        /// Ejecutor genérico de sentencias SQL sobre la base local del usuario (<c>Module.cnnU</c>).
        /// Abre la conexión, ejecuta el query y la cierra. No retorna resultados.
        /// </summary>
        /// <param name="query">Sentencia SQL a ejecutar (INSERT, UPDATE, DELETE).</param>
        /// <exception cref="Exception">Relanza con prefijo <c>'EjecutarQueryBDUsu'</c> ante cualquier error.</exception>
        public static void EjecutarQueryBDUsu(string query)
        {
            using (var accesscm = new OleDbCommand())
            {
                try
                {
                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();

                    Module.ConectarDBUsuario();
                    accesscm.Connection = Module.cnnU;
                    accesscm.CommandType = CommandType.Text;
                    accesscm.CommandText = query;
                    accesscm.ExecuteNonQuery();

                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();
                }
                catch (Exception ex)
                {
                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();
                    throw new Exception($", dentro de 'EjecutarQueryBDUsu': {ex.Message}");
                }
            }
        }
        /// <summary>
        /// Busca el artículo especificado en la base del distribuidor y llama a
        /// <see cref="GrabarArticuloaPreparacAsync"/> para grabarlo en PREPARAC.
        /// </summary>
        /// <remarks>
        /// <b>Prerrequisitos:</b> antes de llamar a este método se deben asignar externamente:
        /// <list type="bullet">
        /// <item><c>WHERE_PREPARAC</c>: cláusula WHERE del artículo a buscar (ej: <c>" WHERE A.ARCODI='X'..."</c>).</item>
        /// <item><c>CANT_PREPARAC</c>: cantidad a insertar/actualizar.</item>
        /// </list>
        /// La consulta une las tablas <c>LISTAS</c>, <c>FAMILIAS</c> y <c>proveedo</c> de la base del distribuidor
        /// y usa <c>SELECT TOP 1</c> para obtener el primer artículo coincidente con el filtro.
        /// Los precios se calculan inline con la función <c>Format(..., 'Currency')</c> de Access.
        /// </remarks>
        /// <param name="accion">Se pasa a <see cref="GrabarArticuloaPreparacAsync"/>: "INSERT" o "UPDATE".</param>
        public static async Task BuscarArticuloParaPreparacAsync(string accion)
        {
            using var da = new OleDbDataAdapter();
            string sql = "SELECT TOP 1 A.ARPRIM AS PRIMERA, A.ARSEGU AS SEGUNDA, A.ARTERC AS TERCERA, A.ARCUAR AS CUARTA, A.ARCODI AS CODIGO, A.ARDESC AS DESCRIPCION, a.armode as Linea " +
                        ", P.PRMARC AS MARCA, A.ARFF AS STOCK" +
                        ", format (A.AROO+(A.AROO*F.FAMIVA/100),'Currency') as PPAORIG" +
                        $", format ({Module.precio_costo_sin_iva.ToString().Replace(",", ".")}+({Module.precio_costo_sin_iva.ToString().Replace(",", ".")}*F.FAMIVA/100),'Currency') AS PRECIO" +
                        ", format (A.AROO,'Currency') AS PPAPREC" +
                        $", format ({Module.precio_final.ToString().Replace(",", ".")}*{CANT_PREPARAC.ToString()},'Currency') as PPACIVA" +
                        $", format ({Module.precio_final.ToString().Replace(",", ".")},'Currency') as PPAPRECUNITARIO" +
                        ", A.ARUBIC1 AS UBICACION, A.AROBSE AS OBSERVACION " +
                        ", A.ARFLIA AS FAMILIA, A.ARPROV AS PROVEEDOR, A.ARORIG AS ORIGEN, A.ARUNID AS UNIDAD, F.FAMIVA AS FAMIVA, A.ARRELA AS RELACION, A.ARBARR, A.ARDEPO, P.PRMARC " +
                        " FROM FAMILIAS AS F INNER JOIN (proveedo AS p INNER JOIN LISTAS AS A ON p.PRCODI = A.ARPROV) ON F.FACODI = A.ARFLIA ";

            try
            {
                if (!string.IsNullOrEmpty(sql))
                {
                    da.SelectCommand = new OleDbCommand(sql + WHERE_PREPARAC, Module.cnn);
                    var ds = new DataSet();
                    da.Fill(ds);
                    if (ds.Tables[0].Rows.Count != 0)
                    {
                        await GrabarArticuloaPreparacAsync(ds, accion);
                    }
                }

                if (Module.cnn.State == ConnectionState.Open)
                {
                    Module.cnn.Close();
                }
            }
            catch (Exception ex)
            {
                if (Module.cnn.State == ConnectionState.Open)
                {
                    Module.cnn.Close();
                }
                throw new Exception($"Error en la Función 'LlenarGrillaArticulos': {ex.Message}");
            }
        }

        /// <summary>
        /// Orquesta el envío del pedido completo al web service WSPalomar.
        /// </summary>
        /// <remarks>
        /// Flujo:
        /// <list type="number">
        /// <item>Llama a <c>wsPreparac.TraerNPedido()</c> para obtener el próximo número de pedido.</item>
        /// <item>Si el número es válido (&gt; 0), actualiza <c>PPANUME</c> en todos los registros PREPARAC.</item>
        /// <item>Llama a <see cref="InsertWSPreparac"/> para enviar cada ítem al WS.</item>
        /// <item>Informa el resultado mediante <c>MessageBox</c>.</item>
        /// </list>
        /// <b>Nota:</b> originalmente era async; las llamadas async están comentadas y reemplazadas
        /// por versiones síncronas.
        /// </remarks>
        public static void ProcesoWS()//async Task ProcesoWSAsync()
        {
            try
            {
                // Obtener y actualizar número de pedido
                var wsPreparac = new WSPalomarSoapClient(WSPalomarSoapClient.EndpointConfiguration.WSPalomarSoap);
                
                int numPed = (int)wsPreparac.TraerNPedido();// await wsPreparac.TraerNPedidoAsync();
                if (numPed != 0)
                {
                    Module.UpdateValorDBLocal($"UPDATE PREPARAC SET PPANUME={numPed}"); //await Module.UpdateValorDBLocalAsync($"UPDATE PREPARAC SET PPANUME={numPed}");
                }
                
                int resultado = InsertWSPreparac();// await InsertWSPreparacAsync();
                if (resultado == 1)
                {
                    MessageBox.Show($"Pedido Cargado por Web Service, Nº Pedido= {numPed}", 
                        "Información de proceso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Pedido NO cargado por Web Service, Nº Pedido= {numPed}", 
                        "Información de proceso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($", dentro de 'ProcesoWS': {ex.Message}");
            }
        }

        /// <summary>
        /// Lee todos los registros de la tabla <c>PREPARAC</c> y los envía uno a uno
        /// al web service WSPalomar mediante <c>wsPreparac.InsertPreparac()</c>.
        /// </summary>
        /// <remarks>
        /// Valores fijos que el WS requiere y no están en PREPARAC:
        /// <list type="bullet">
        /// <item><c>_PPACATA = "6"</c>: código del catálogo (hardcoded).</item>
        /// <item><c>_PPAESTADO = "1"</c>: estado inicial del pedido (hardcoded).</item>
        /// <item><c>_PPAHORA</c>: hora de envío generada en tiempo real.</item>
        /// </list>
        /// <b>Advertencia:</b> el bloque <c>catch</c> tiene un <c>return resultado</c> antes del <c>throw</c>,
        /// lo que hace que la excepción sea inalcanzable. El código retorna 0 ante cualquier error.
        /// </remarks>
        /// <returns>Resultado de la última llamada al WS: 1 = éxito, 0 = error o sin registros.</returns>
        public static int InsertWSPreparac()// async Task<int> InsertWSPreparacAsync()
        {
            int resultado = 0;
            string fechap_ = "";
            var wsPreparac = new WSPalomarSoapClient(WSPalomarSoapClient.EndpointConfiguration.WSPalomarSoap);

            Module.ConectarDBUsuario();
            var da = new OleDbDataAdapter(@"SELECT PPANUME,PPAESTA,PPACLIE,PPABARR,PPACOD1,PPACOD2,PPACOD3,PPACOD4,
                            PPADESC,PPACANT,PPAFECH,PPATPRE,PPAPREC,PPACIVA,PPARAZO,PPASITI,PPAQOT,PPAPPA,PPADESK,
                            PPAORIG,PPACODI,PPATARJ,PPACUOT,PPAVEND,PPAFLIA,PPAIMPO,PPAPROV,PPAAUTO,PPAARFF,PPAFIVA,PPAOBSE,
                            PPAUSER,PPACOMP,PPABULT,PPAKILO,PPASEGU,PPAREMF,PPAREMN,PPAPIEZ,PPAMARC,PPAPOR1,PPAPOR2,PPABASE,
                            PPACONT,PPAFAPL,PPAFORI,PPAUNID,PPADEPO,PPADIRE,PPACPAG
                            from preparac as pr", Module.cnnU);
            try
            {
                var ds = new DataSet();
                da.Fill(ds, "preparac");
                fechap_ = DateTime.Now.ToString();

                // Valores fijos por pedido
                string _PPAHORA = $"{DateTime.Now.Hour}:{DateTime.Now.Minute}:{DateTime.Now.Millisecond}";
                string _PPACATA = "6"; // valor fijo ES EL CODIGO DEL CATALOGO
                string _PPAESTADO = "1"; // valor fijo

                foreach (DataRow fila in ds.Tables["preparac"].Rows)
                {
                    //resultado = await wsPreparac.InsertPreparacAsync(
                    resultado = wsPreparac.InsertPreparac(
                        Convert.ToDouble(fila["PPANUME"]),
                        Convert.ToDouble(fila["PPAESTA"]),
                        Convert.ToDouble(fila["PPACLIE"]),
                        Convert.ToDouble(fila["PPABARR"]),
                        fila["PPACOD1"].ToString(),
                        fila["PPACOD2"].ToString(),
                        fila["PPACOD3"].ToString(),
                        fila["PPACOD4"].ToString(),
                        fila["PPADESC"].ToString(),
                        Convert.ToDouble(fila["PPACANT"]),
                        Convert.ToDateTime(fila["PPAFECH"]),
                        fila["PPATPRE"].ToString(),
                        Convert.ToDouble(fila["PPAPREC"]),
                        Convert.ToDouble(fila["PPACIVA"]),
                        fila["PPARAZO"].ToString(),
                        Convert.ToDouble(fila["PPASITI"]),
                        fila["PPAQOT"].ToString(),
                        Convert.ToDouble(fila["PPAPPA"]),
                        Convert.ToDouble(fila["PPADESK"]),
                        Convert.ToDouble(fila["PPAORIG"]),
                        fila["PPACODI"].ToString(),
                        fila["PPATARJ"].ToString(),
                        Convert.ToDouble(fila["PPACUOT"]),
                        Convert.ToDouble(fila["PPAVEND"]),
                        Convert.ToDouble(fila["PPAFLIA"]),
                        fila["PPAIMPO"].ToString(),
                        Convert.ToDouble(fila["PPAPROV"]),
                        Convert.ToInt32(fila["PPAAUTO"]),
                        fila["PPAARFF"].ToString(),
                        Convert.ToDouble(fila["PPAFIVA"]),
                        fila["PPAOBSE"].ToString(),
                        Convert.ToDouble(fila["PPAUSER"]),
                        fila["PPACOMP"].ToString(),
                        Convert.ToDouble(fila["PPABULT"]),
                        Convert.ToDouble(fila["PPAKILO"]),
                        Convert.ToDouble(fila["PPASEGU"]),
                        Convert.ToDouble(fila["PPAREMF"]),
                        Convert.ToDouble(fila["PPAREMN"]),
                        Convert.ToDouble(fila["PPAPIEZ"]),
                        fila["PPAMARC"].ToString(),
                        Convert.ToDouble(fila["PPAPOR1"]),
                        Convert.ToDouble(fila["PPAPOR2"]),
                        Convert.ToDouble(fila["PPABASE"]),
                        Convert.ToDouble(fila["PPACONT"]),
                        Convert.ToDateTime(fila["PPAFAPL"]),
                        Convert.ToDateTime(fila["PPAFORI"]),
                        Convert.ToDouble(fila["PPAUNID"]),
                        Convert.ToDouble(fila["PPADEPO"]),
                        Convert.ToDouble(fila["PPADIRE"]),
                        Convert.ToDouble(fila["PPACPAG"]),
                        _PPAHORA,
                        Convert.ToDecimal(_PPACATA),
                        Convert.ToDecimal(_PPAESTADO));
                }

                ds.Dispose();
                return resultado;
            }
            catch (Exception ex)
            {
                return resultado;
                throw new Exception($", dentro de 'InsertWSPreparac': {ex}");
            }
            finally
            {
                da.Dispose();
            }
        }

        /// <summary>
        /// Actualiza la cantidad de un ítem del carrito y recalcula automáticamente
        /// el precio total con IVA (<c>PPACIVA</c>) y el subtotal sin IVA (<c>SUBTSIVA</c>).
        /// </summary>
        /// <remarks>
        /// Flujo:
        /// <list type="number">
        /// <item>Llama a <see cref="TraerExtrasPreparac"/> para cargar el precio unitario y costo del ítem en los campos estáticos.</item>
        /// <item>Construye un UPDATE usando la función <c>Format(..., 'Currency')</c> de Access para calcular
        /// los nuevos valores de <c>PPACIVA</c> y <c>SUBTSIVA</c> directamente en SQL.</item>
        /// <item>Refresca el carrito llamando a <c>Module.funcion_inicial_carrito()</c>.</item>
        /// </list>
        /// </remarks>
        /// <param name="_item">Autonúmero del registro a modificar (columna <c>PPAAUTN</c>).</param>
        /// <param name="_cant">Nueva cantidad como cadena (se embebe directamente en el SQL).</param>
        public static void Update_item(long _item, string _cant)
        {
            using (OleDbCommand cmUsu = new OleDbCommand())
            {
                try
                {
                    string SelectPreparacExtras = "SELECT familias.FAIVA,PREPARAC.PPAPRECUNITARIO,PREPARAC.PPAPRECCOSTO " +
                                                "FROM familias INNER JOIN PREPARAC ON familias.FACODI = PREPARAC.PPAFLIA " +
                                                "WHERE PREPARAC.PPAAUTN=" + _item.ToString();

                    // Paso 1: obtener precio unitario y costo actuales del ítem para recalcular con la nueva cantidad
                    TraerExtrasPreparac(SelectPreparacExtras); // Asume versión async

                    string Query = "UPDATE PREPARAC SET PPACANT='" + _cant.ToString() +
                                  "',PPACIVA=format(" + _PPAPRECUNITARIO.ToString().Replace(",", ".") + "*" + _cant.ToString() + ",'Currency') " +
                                  ",SUBTSIVA=format(" + _PPAPRECCOSTO.ToString().Replace(",", ".") + "*" + _cant.ToString() + ",'Currency') " +
                                  " WHERE PPAAUTN=" + _item.ToString();

                    Module.ConectarDBUsuario(); // Asume versión async
                    cmUsu.Connection = Module.cnnU;
                    cmUsu.CommandType = CommandType.Text;
                    cmUsu.CommandText = Query;

                    cmUsu.ExecuteNonQuery();

                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();

                    Module.funcion_inicial_carrito(); // Asume versión async
                }
                catch (Exception ex)
                {
                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();

                    Module.funcion_inicial_carrito(); // Asume versión async
                    throw new Exception(",Error dentro de 'Update_item'" + ex.Message);
                }
            }
        }
        /// <summary>
        /// Helper privado: ejecuta una consulta SELECT y carga en los campos estáticos
        /// <c>_FAMIVA</c>, <c>_PPAPRECUNITARIO</c> y <c>_PPAPRECCOSTO</c> del ítem indicado.
        /// </summary>
        /// <remarks>
        /// Se usa como paso previo en <see cref="Update_item"/> para obtener los precios actuales
        /// antes de recalcular con la nueva cantidad.
        /// En caso de error o sin resultados, los tres campos se ponen en 0.
        /// </remarks>
        /// <param name="querySelect">Sentencia SELECT que retorna FAIVA, PPAPRECUNITARIO y PPAPRECCOSTO.</param>
        private static void TraerExtrasPreparac(string querySelect)
        {
            OleDbCommand cmUsu = new OleDbCommand();
            OleDbDataReader drUsu = null;

            try
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = querySelect;
                drUsu = cmUsu.ExecuteReader();

                if (drUsu.Read())
                {
                    _FAMIVA = Convert.ToDecimal(drUsu["FAIVA"]);
                    _PPAPRECUNITARIO = Convert.ToDecimal(drUsu["PPAPRECUNITARIO"]);
                    _PPAPRECCOSTO = Convert.ToDecimal(drUsu["PPAPRECCOSTO"]);
                }
                else
                {
                    _FAMIVA = 0;
                    _PPAPRECUNITARIO = 0;
                    _PPAPRECCOSTO = 0;
                }

                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();
            }
            catch (Exception ex)
            {
                if (Module.cnnU.State == ConnectionState.Open)
                    Module.cnnU.Close();

                _FAMIVA = 0;
                _PPAPRECUNITARIO = 0;
                _PPAPRECCOSTO = 0;
                throw new Exception(", dentro de 'TraerExtrasPreparac': " + ex.Message);
            }
            finally
            {
                cmUsu.Dispose();
                drUsu?.Dispose();
            }
        }
        /// <summary>
        /// Helper privado async: ejecuta una consulta SELECT y retorna el valor de <c>FAIVA</c>
        /// de la familia del artículo actualmente en procesamiento.
        /// </summary>
        /// <remarks>
        /// Consulta la tabla <c>FAMILIAS</c> unida a <c>PREPARAC</c> para obtener el IVA
        /// correspondiente al artículo por su <c>PPACODI</c>.
        /// Si no hay resultado retorna 0; el llamador (<see cref="GrabarArticuloaPreparacAsync"/>)
        /// asigna 21% como valor por defecto cuando esto ocurre.
        /// </remarks>
        /// <param name="querySelect">Sentencia SELECT que retorna el campo FAIVA.</param>
        /// <returns>Porcentaje de IVA de la familia, o 0 si no se encuentra.</returns>
        private static async Task<decimal> TraerFAMIVAAsync(string querySelect)
        {
            using (OleDbCommand cmUsu = new OleDbCommand())
            {
                decimal _IVA = 0;

                try
                {
                    if (Module.cnnU.State == ConnectionState.Open)
                        await Module.cnnU.CloseAsync();

                    Module.ConectarDBUsuario(); // Asume que tienes una versión async de este método
                    cmUsu.Connection = Module.cnnU;
                    cmUsu.CommandType = CommandType.Text;
                    cmUsu.CommandText = querySelect;

                    using (var drUsu = await cmUsu.ExecuteReaderAsync())
                    {
                        if (await drUsu.ReadAsync())
                        {
                            _IVA = Convert.ToDecimal(drUsu["FAIVA"]);
                        }
                        else
                        {
                            _IVA = 0;
                        }
                    }

                    if (Module.cnnU.State == ConnectionState.Open)
                        await Module.cnnU.CloseAsync();

                    return _IVA;
                }
                catch (Exception ex)
                {
                    if (Module.cnnU.State == ConnectionState.Open)
                        await Module.cnnU.CloseAsync();

                    throw new Exception(", dentro de 'TraerFAMIVAAsync': " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Verifica si un artículo ya fue agregado al carrito (tabla <c>PREPARAC</c>).
        /// Como efecto secundario, carga en <see cref="CantExistente"/> la cantidad actual
        /// del artículo si se lo encuentra.
        /// </summary>
        /// <remarks>
        /// La clave de búsqueda es la combinación <c>PPACOD3 + PPAFLIA + PPAPROV</c>,
        /// que identifica unívocamente un artículo dentro del pedido.
        /// </remarks>
        /// <param name="_PPACOD3">Tercera parte del código del artículo (10 car.).</param>
        /// <param name="_PPAFLIA">Código de familia del artículo.</param>
        /// <param name="_PPAPROV">Código de proveedor del artículo.</param>
        /// <returns><c>true</c> si el artículo ya existe en PREPARAC; <c>false</c> si no o ante error.</returns>
        public static bool ExisteARTICULOPedido(string _PPACOD3, string _PPAFLIA, string _PPAPROV)
        {
            OleDbCommand cmUsu = new OleDbCommand();
            OleDbDataReader dr = null;
            bool existe = false;

            try
            {
                Module.ConectarDBUsuario();
                cmUsu.Connection = Module.cnnU;
                cmUsu.CommandType = CommandType.Text;
                cmUsu.CommandText = "SELECT PPACOD3, PPACANT FROM PREPARAC " +
                                   " WHERE PPACOD3='" + _PPACOD3.ToString() +
                                   "' AND PPAFLIA=" + _PPAFLIA.ToString() +
                                   " AND PPAPROV=" + _PPAPROV.ToString();

                dr = cmUsu.ExecuteReader();
                // Efecto secundario: CantExistente queda con la cantidad actual en PREPARAC
                // para que el llamador pueda sumar o reemplazar según su lógica de negocio
                CantExistente = 0;
                if (dr.Read())
                {
                    CantExistente = Convert.ToDecimal(dr[1].ToString());
                    existe = true;
                }

                dr.Close();
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
                cmUsu.Dispose();
                dr?.Dispose();
            }
        }


        /// <summary>
        /// Elimina un ítem específico del carrito (tabla <c>PREPARAC</c>) por su autonúmero.
        /// </summary>
        /// <param name="item">Autonúmero del registro a eliminar (columna <c>PPAAUTN</c>).</param>
        /// <returns><c>true</c> si la eliminación fue exitosa; <c>false</c> ante cualquier error.</returns>
        public static bool EliminarItemCarro(int item)
        {
            using (OleDbCommand cm = new OleDbCommand())
            {
                try
                {
                    Module.ConectarDBUsuario();
                    cm.Connection = Module.cnnU;
                    cm.CommandType = CommandType.Text;
                    cm.CommandText = "DELETE from PREPARAC where PPAAUTN=" + item.ToString();
                    cm.ExecuteNonQuery();

                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();

                    return true;
                }
                catch (Exception)
                {
                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();

                    return false;
                }
            }
        }

        /// <summary>
        /// Vacía completamente el carrito eliminando <b>todos</b> los registros de la tabla <c>PREPARAC</c>
        /// en la base local del usuario. Equivalente a "limpiar el pedido".
        /// </summary>
        /// <returns><c>true</c> si todos los registros fueron eliminados; <c>false</c> ante cualquier error.</returns>
        public static bool EliminarItemsCarro()
        {
            using (OleDbCommand cm = new OleDbCommand())
            {
                try
                {
                    Module.ConectarDBUsuario();
                    cm.Connection = Module.cnnU;
                    cm.CommandType = CommandType.Text;
                    cm.CommandText = "DELETE from PREPARAC";
                    cm.ExecuteNonQuery();

                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();

                    return true;
                }
                catch (Exception)
                {
                    if (Module.cnnU.State == ConnectionState.Open)
                        Module.cnnU.Close();

                    return false;
                }
            }
        }
    }
}

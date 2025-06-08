using System;
using System.Data;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.Windows.Forms;
using ServiceReferenceWSPalomar;

namespace DistribuidoraPalomarNet
{
    public static class ModuloPreparac
    {
        #region Variables Preparac
        public static string Codi = "";
        public static decimal Precio = 0;
        public static decimal Cant; // As Integer en VB
        public static decimal CantExistente=0;
        public static decimal SUBT, Iva;
        public static string _PPADESC = ""; // 40
        public static string _PPAOBSE = ""; // 255
        public static decimal _PPAORIG = 0;
        public static decimal _PPAPREC = 0;
        public static decimal _PPABASE = 0;
        public static decimal _PPACIVA = 0;
        public static decimal _PPAPRECUNITARIO = 0;
        public static decimal _PPAPRECCOSTO = 0;
        public static decimal _FAMIVA = 0;

        // PARA PRECIO SUBTOTAL Y TOTAL SIN IVA
        public static decimal _SUBTSIVA = 0;
        public static decimal _TOTALSIVA = 0;

        public static decimal _PPABARR = 0;
        public static string _PPACOD1 = ""; // 3
        public static string _PPACOD2 = ""; // 4
        public static string _PPACOD3 = ""; // 10
        public static string _PPACOD4 = ""; // 4
        public static string _PPACODI = ""; // 3
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
        public static string _PPAIMPO = ""; // 1
        public static decimal _PPAPROV = 0;
        public static decimal _PPAAUTO = 1;
        public static string _PPAARFF = ""; // 10
        public static decimal _PPAFIVA = 0;
        public static decimal _PPAUSER = 0;
        public static string _PPACOMP = ""; // 30
        public static decimal _PPABULT = 0;
        public static decimal _PPAKILO = 0;
        public static decimal _PPASEGU = 0;
        public static decimal _PPAREMF = 0;
        public static decimal _PPAREMN = 0;
        public static decimal _PPAPIEZ = 1;
        public static string _PPAMARC = ""; // 15
        public static string _PRMARC = ""; // Campo de fantasía, solo para mostrar en el pdf
        public static decimal _PPAPOR1 = 0;
        public static decimal _PPAPOR2 = 0;
        public static decimal _PPACONT = 0;
        public static string _PPAFAPL = DateTime.Now.ToShortDateString();
        public static string _PPAFORI = DateTime.Now.ToShortDateString();
        public static decimal _PPAUNID = 1;
        public static decimal _PPADEPO = 0;
        public static decimal _PPADIRE = 0;
        public static decimal _PPACPAG = 0;

        // Variable where articulo para preparac
        public static string WHERE_PREPARAC = "";
        public static decimal CANT_PREPARAC = 0;
        #endregion

        public static async Task GrabarArticuloaPreparacAsync(DataSet DTS, string Accion)
        {
            string strQuery = "";
            try
            {
                foreach (DataRow fila in DTS.Tables[0].Rows)
                {
                    string Codi = $"{fila["PRIMERA"].ToString().Trim()} - {fila["TERCERA"].ToString().Trim()}";
                    decimal Precio = Math.Round(Module.precio_costo_sin_iva, 2, MidpointRounding.ToEven);
                    _PPAPRECCOSTO = Math.Round(Module.precio_costo_sin_iva, 2, MidpointRounding.ToEven);

                    if (Precio != 0)
                    {
                        decimal SUBT = 0, Iva = 0;
                        string Cant = ModuloPreparac.CANT_PREPARAC.ToString().Replace(".", ",");

                        if (Cant != "0")
                        {
                            if (decimal.TryParse(Cant, out decimal cantidadDecimal))
                            {
                                SUBT = Math.Round(Precio * cantidadDecimal, 2, MidpointRounding.ToEven);
                            }

                            _PPADESC = !string.IsNullOrEmpty(fila["DESCRIPCION"].ToString()) ? fila["DESCRIPCION"].ToString() : "";
                            _PPAOBSE = !string.IsNullOrEmpty(fila["OBSERVACION"].ToString()) ? fila["OBSERVACION"].ToString() : "";
                            //_PPAORIG = !string.IsNullOrEmpty(fila["PPAORIG"].ToString())
                            //? Convert.ToDecimal(fila["PPAORIG"].ToString().Replace("$", "").Replace(".", "").Replace(",", "."))
                            //: 0;

                            _PPAORIG = !string.IsNullOrEmpty(fila["PPAORIG"].ToString())
                            ? Convert.ToDecimal(fila["PPAORIG"].ToString().Replace("$", "").Replace(".", ""))
                            : 0;
                            //_PPAORIG = !string.IsNullOrEmpty(fila["PPAORIG"].ToString()) ? Convert.ToDecimal(fila["PPAORIG"]) : 0;

                            _PPAPREC = !string.IsNullOrEmpty(fila["PPAPREC"].ToString())
                            ? Convert.ToDecimal(fila["PPAPREC"].ToString().Replace("$", "").Replace(".", ""))
                            : 0;
                            _PPABASE = !string.IsNullOrEmpty(fila["PPAPREC"].ToString())
                            ? Convert.ToDecimal(fila["PPAPREC"].ToString().Replace("$", "").Replace(".", ""))
                            : 0;

                            //_PPAPREC = Convert.ToDecimal(fila["PPAPREC"]);
                            //_PPABASE = Convert.ToDecimal(fila["PPAPREC"]);
                            _PPACIVA = !string.IsNullOrEmpty(fila["PPACIVA"].ToString())
                            ? Convert.ToDecimal(fila["PPACIVA"].ToString().Replace("$", "").Replace(".", ""))
                            : 0;
                            _PPAPRECUNITARIO = !string.IsNullOrEmpty(fila["PPAPRECUNITARIO"].ToString())
                            ? Convert.ToDecimal(fila["PPAPRECUNITARIO"].ToString().Replace("$", "").Replace(".", ""))
                            : 0;

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
                            _PPARAZO = Module.US_CLRAZO;
                            _PPASITI = Module.US_CLCIVA;
                            _PPAVEND = Module.US_CLVEND;

                            string SelectFAMIVA = $"SELECT familias.FAMIVA FROM familias INNER JOIN PREPARAC ON familias.FACODI = PREPARAC.PPAFLIA WHERE PREPARAC.PPACODI='{_PPACODI}'";
                            _PPAFIVA = await TraerFAMIVAAsync(SelectFAMIVA);
                            if (_PPAFIVA == 0) _PPAFIVA = 21;

                            _SUBTSIVA = Math.Round(Module.precio_costo_sin_iva * cantidadDecimal, 2, MidpointRounding.ToEven);
                            _PPACPAG = Module.US_CLCPAG;
                            _PPANUME = 0;

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
        public static async Task BuscarArticuloParaPreparacAsync(string accion)
        {
            using var da = new OleDbDataAdapter();
            string sql = "SELECT TOP 1 A.ARPRIM AS PRIMERA, A.ARSEGU AS SEGUNDA, A.ARTERC AS TERCERA, A.ARCUAR AS CUARTA, A.ARCODI AS CODIGO, A.ARDESC AS DESCRIPCION, A.ARMODE AS MODELO " +
                        ", A.ARMARC AS MARCA, A.ARFF AS STOCK" +
                        ", format (A.AROO+(A.AROO*F.FAMIVA/100),'Currency') as PPAORIG" +
                        $", format ({Module.precio_costo_sin_iva.ToString().Replace(",", ".")}+({Module.precio_costo_sin_iva.ToString().Replace(",", ".")}*F.FAMIVA/100),'Currency') AS PRECIO" +
                        ", format (A.AROO,'Currency') AS PPAPREC" +
                        $", format ({Module.precio_final.ToString().Replace(",", ".")}*{CANT_PREPARAC.ToString()},'Currency') as PPACIVA" +
                        $", format ({Module.precio_final.ToString().Replace(",", ".")},'Currency') as PPAPRECUNITARIO" +
                        ", A.ARUBIC1 AS UBICACION, A.AROBSE AS OBSERVACION " +
                        ", A.ARFLIA AS FAMILIA, A.ARPROV AS PROVEEDOR, A.ARORIG AS ORIGEN, A.ARUNID AS UNIDAD, F.FAMIVA AS FAMIVA, A.ARRELA AS RELACION, A.ARBARR, A.ARDEPO, P.PRMARC " +
                        " FROM FAMILIAS AS F INNER JOIN (proveedo AS p INNER JOIN ARTICULO AS A ON p.PRCODI = A.ARPROV) ON F.FACODI = A.ARFLIA ";

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

        public static async Task ProcesoWSAsync()
        {
            try
            {
                // Obtener y actualizar número de pedido
                var wsPreparac = new WSPalomarSoapClient(WSPalomarSoapClient.EndpointConfiguration.WSPalomarSoap);
                int numPed = (int)await wsPreparac.TraerNPedidoAsync();
                if (numPed != 0)
                {
                    await Module.UpdateValorDBLocalAsync($"UPDATE PREPARAC SET PPANUME={numPed}");
                }
                
                int resultado = await InsertWSPreparacAsync();
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

        public static async Task<int> InsertWSPreparacAsync()
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
                    resultado = await wsPreparac.InsertPreparacAsync(
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

        public static void Update_item(long _item, string _cant)
        {
            using (OleDbCommand cmUsu = new OleDbCommand())
            {
                try
                {
                    string SelectPreparacExtras = "SELECT familias.FAMIVA,PREPARAC.PPAPRECUNITARIO,PREPARAC.PPAPRECCOSTO " +
                                                "FROM familias INNER JOIN PREPARAC ON familias.FACODI = PREPARAC.PPAFLIA " +
                                                "WHERE PREPARAC.PPAAUTN=" + _item.ToString();

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
                    _FAMIVA = Convert.ToDecimal(drUsu["FAMIVA"]);
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
                            _IVA = Convert.ToDecimal(drUsu["FAMIVA"]);
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

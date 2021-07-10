using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;

using System.Timers;
using System.Data;
using System.Threading;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.IO;
using S7.Net;
using System.Text.RegularExpressions;
using System.Reflection;


namespace ServicioPLCSQL
{
    public partial class ServicioPLCSQL : ServiceBase
    {
        Plc oPLC = null;
        cPLC sPLC = new cPLC();
        cDispositivos oDirecciones = new cDispositivos();
        cVarGlobales oVarGlobales = new cVarGlobales();

        System.Timers.Timer Tempo = null;

        public ServicioPLCSQL()
        {
            InitializeComponent();

            Tempo = new System.Timers.Timer();

            LogRegistro = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("ServicioPLCSQL"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "ServicioPLCSQL", "Application");
            }

            LogRegistro.Source = "ServicioPLCSQL";
            LogRegistro.Log = "Application";
        }

        //public void OnDebug()
        //{
        //    OnStart(null);
        //}

        protected override void OnStart(string[] args)
        {
            Direcciones();

            ConectarPLC();

            this.Tempo.Enabled = true;
            this.Tempo.Interval = 5000;
            this.Tempo.Elapsed += new ElapsedEventHandler(Temporizador_Elapsed);
            this.Tempo.Start();

            LogRegistro.WriteEntry("Servicio Iniciado en forma correcta" + DateTime.Now.ToString());
        }

        protected override void OnStop()
        {
            Tempo.Stop();
            Tempo.Dispose();
            LogRegistro.WriteEntry("Servicio Detenido en forma correcta" + DateTime.Now.ToString());
        }

        void Direcciones()
        {
            try
            {
                oDirecciones.IpPLC = ConfigurationManager.AppSettings["IpPCL"].ToString();
                oDirecciones.PortPLC = Convert.ToInt16(ConfigurationManager.AppSettings["PortPLC"].ToString());
                oDirecciones.RackPLC = Convert.ToInt16(ConfigurationManager.AppSettings["RackPLC"].ToString());
                oDirecciones.SlotPLC = Convert.ToInt16(ConfigurationManager.AppSettings["SlotPLC"].ToString());
                oDirecciones.NumDb = Convert.ToInt16(ConfigurationManager.AppSettings["NumDb"].ToString());
            }
            catch (Exception ex)
            {
                LogRegistro.WriteEntry(ex.Message.ToString() + " Funcion Direcciones");
            }
        }

        void ConectarPLC()
        {
            try
            {
                sPLC.IP = oDirecciones.IpPLC;

                if (sPLC.IsAvailable)
                {
                    oPLC = new Plc(CpuType.S71200, oDirecciones.IpPLC, oDirecciones.RackPLC, oDirecciones.SlotPLC);
                    oPLC.Open();

                    cVarGlobales.EstConePLC = oPLC.IsConnected;

                    if (cVarGlobales.EstConePLC)
                    {
                        oVarGlobales.MensajePLC = "PLC Conectado " + oDirecciones.IpPLC;
                        LogRegistro.WriteEntry(oVarGlobales.MensajePLC);
                    }
                    else
                    {
                        oVarGlobales.MensajePLC = "PLC Desconectado" + oDirecciones.IpPLC;
                        LogRegistro.WriteEntry(oVarGlobales.MensajePLC);
                    }
                }
                else
                {
                    cVarGlobales.EstConePLC = false;
                    oVarGlobales.MensajePLC = "PLC Desconectado" + oDirecciones.IpPLC;
                    LogRegistro.WriteEntry(oVarGlobales.MensajePLC);
                }
            }
            catch (Exception ex)
            {
                cVarGlobales.EstConePLC = false;
                oVarGlobales.MensajePLC = ex.Message.ToString();
                LogRegistro.WriteEntry(oVarGlobales.MensajePLC);
            }
        }

        void Temporizador_Elapsed(object sender, ElapsedEventArgs e)
        {
            Tempo.Stop();
            try
            {
                if ((sPLC.IsAvailable) != true)
                {
                    cVarGlobales.EstConePLC = false;
                    ConectarPLC();
                    LeerEtiquetasDerivadas();
                    LeerEtiquetasRecirculadas();
                }
                else
                {
                    if (cVarGlobales.EstConePLC)
                    {
                        LeerEtiquetasDerivadas();
                        LeerEtiquetasRecirculadas();
                    }
                }

            }
            catch (Exception ex)
            {
                LogRegistro.WriteEntry(ex.Message.ToString() + " funcion PLC SQL");
            }

            Tempo.Start();

        }

        private void LeerEtiquetasDerivadas()
        {
            int in_LineaSalida;
            int in_IdSalida;
            int in_Tipo;

            cEntDerivados oSalidas = new cEntDerivados();
            try
            {
                var Derivadasresult1 = (uint)oPLC.Read("DB1010.DBD10004");
                var Derivadasresult2 = (uint)oPLC.Read("DB1010.DBD10008");
                var Derivadasresult3 = (uint)oPLC.Read("DB1010.DBD10012");
                var Derivadasresult4 = (uint)oPLC.Read("DB1010.DBD10016");
                var Derivadasresult5 = (uint)oPLC.Read("DB1010.DBD10020");
                var Derivadasresult6 = (uint)oPLC.Read("DB1010.DBD10024");
                var Derivadasresult7 = (uint)oPLC.Read("DB1010.DBD10028");
                var Derivadasresult8 = (uint)oPLC.Read("DB1010.DBD10032");
                var Derivadasresult9 = (uint)oPLC.Read("DB1010.DBD10036");
                var Derivadasresult10 = (uint)oPLC.Read("DB1010.DBD10040");
                var Derivadasresult11 = (uint)oPLC.Read("DB1010.DBD10044");
                var Derivadasresult12 = (uint)oPLC.Read("DB1010.DBD10048");
                var Derivadasresult13 = (uint)oPLC.Read("DB1010.DBD10052");
                var Derivadasresult14 = (uint)oPLC.Read("DB1010.DBD10056");
                var Derivadasresult15 = (uint)oPLC.Read("DB1010.DBD10060");
                var Derivadasresult16 = (uint)oPLC.Read("DB1010.DBD10064");
                var Derivadasresult17 = (uint)oPLC.Read("DB1010.DBD10068");
                var Derivadasresult18 = (uint)oPLC.Read("DB1010.DBD10072");
                var Derivadasresult19 = (uint)oPLC.Read("DB1010.DBD10076");
                var Derivadasresult20 = (uint)oPLC.Read("DB1010.DBD10080");
                var Derivadasresult21 = (uint)oPLC.Read("DB1010.DBD10084");
                var Derivadasresult22 = (uint)oPLC.Read("DB1010.DBD10088");
                var Derivadasresult23 = (uint)oPLC.Read("DB1010.DBD10092");

                oSalidas.CartonSalida01 = Convert.ToInt32(Derivadasresult1);
                oSalidas.CartonSalida02 = Convert.ToInt32(Derivadasresult2);
                oSalidas.CartonSalida03 = Convert.ToInt32(Derivadasresult3);
                oSalidas.CartonSalida04 = Convert.ToInt32(Derivadasresult4);
                oSalidas.CartonSalida05 = Convert.ToInt32(Derivadasresult5);
                oSalidas.CartonSalida06 = Convert.ToInt32(Derivadasresult6);
                oSalidas.CartonSalida07 = Convert.ToInt32(Derivadasresult7);
                oSalidas.CartonSalida08 = Convert.ToInt32(Derivadasresult8);
                oSalidas.CartonSalida09 = Convert.ToInt32(Derivadasresult9);
                oSalidas.CartonSalida10 = Convert.ToInt32(Derivadasresult10);
                oSalidas.CartonSalida11 = Convert.ToInt32(Derivadasresult11);
                oSalidas.CartonSalida12 = Convert.ToInt32(Derivadasresult12);
                oSalidas.CartonSalida13 = Convert.ToInt32(Derivadasresult13);
                oSalidas.CartonSalida14 = Convert.ToInt32(Derivadasresult14);
                oSalidas.CartonSalida15 = Convert.ToInt32(Derivadasresult15);
                oSalidas.CartonSalida16 = Convert.ToInt32(Derivadasresult16);
                oSalidas.CartonSalida17 = Convert.ToInt32(Derivadasresult17);
                oSalidas.CartonSalida18 = Convert.ToInt32(Derivadasresult18);
                oSalidas.CartonSalida19 = Convert.ToInt32(Derivadasresult19);
                oSalidas.CartonSalida20 = Convert.ToInt32(Derivadasresult20);
                oSalidas.CartonSalida21 = Convert.ToInt32(Derivadasresult21);
                oSalidas.CartonSalida22 = Convert.ToInt32(Derivadasresult22);
                oSalidas.CartonSalida23 = Convert.ToInt32(Derivadasresult23);

                string res = string.Empty;
                PropertyInfo[] properties = typeof(cEntDerivados).GetProperties();
                foreach (PropertyInfo property in properties)
                {
                    //así obtenemos el nombre del atributo
                    string NombreAtributo = property.Name;

                    //así obtenemos el valor del atributo
                    string Valor = property.GetValue(oSalidas).ToString();

                    string loc_conversion = string.Empty;

                    loc_conversion = Regex.Replace(Valor, @"[^\w\s.!@$%^&*()\-\/]+", "");

                    in_IdSalida = Convert.ToInt32(loc_conversion);
                    in_LineaSalida = Convert.ToInt32(NombreAtributo.Substring(12, 2));                    
                    in_Tipo = 1;

                    //if (conver != string.Empty)
                    if (loc_conversion != "0")
                    {
                        res = CartonDerivado(in_IdSalida, in_LineaSalida, in_Tipo);

                        if (res == "1")
                        {
                            LogRegistro.WriteEntry("Se actualizo el movimiento derivado -------------> " + in_IdSalida.ToString());  
                        }
                        else
                        {
                            //LogRegistro.WriteEntry(res + " Lectura etiquetas derivadas");
                        }
                    }
                }

                //limpio las variables de recepcion de datos
                oSalidas = new cEntDerivados();

            }
            catch (Exception ex)
            {
                LogRegistro.WriteEntry(ex.Message.ToString() + " funcion LeerEtiquetasDerivadas");
            }

        }

        private void LeerEtiquetasRecirculadas()
        {
            int in_LineaSalida;
            int in_IdSalida;
            int in_Tipo;

            cEntRecirculado oRecirculado = new cEntRecirculado();
            try
            {
                var Recirculadoresult1 = (uint)oPLC.Read("DB1010.DBD10104");
                var Recirculadoresult2 = (uint)oPLC.Read("DB1010.DBD10108");
                var Recirculadoresult3 = (uint)oPLC.Read("DB1010.DBD10112");
                var Recirculadoresult4 = (uint)oPLC.Read("DB1010.DBD10116");
                var Recirculadoresult5 = (uint)oPLC.Read("DB1010.DBD10120");
                var Recirculadoresult6 = (uint)oPLC.Read("DB1010.DBD10124");
                var Recirculadoresult7 = (uint)oPLC.Read("DB1010.DBD10128");
                var Recirculadoresult8 = (uint)oPLC.Read("DB1010.DBD10132");
                var Recirculadoresult9 = (uint)oPLC.Read("DB1010.DBD10136");
                var Recirculadoresult10 = (uint)oPLC.Read("DB1010.DBD10140");
                var Recirculadoresult11 = (uint)oPLC.Read("DB1010.DBD10144");
                var Recirculadoresult12 = (uint)oPLC.Read("DB1010.DBD10148");
                var Recirculadoresult13 = (uint)oPLC.Read("DB1010.DBD10152");
                var Recirculadoresult14 = (uint)oPLC.Read("DB1010.DBD10156");
                var Recirculadoresult15 = (uint)oPLC.Read("DB1010.DBD10160");
                var Recirculadoresult16 = (uint)oPLC.Read("DB1010.DBD10164");
                var Recirculadoresult17 = (uint)oPLC.Read("DB1010.DBD10168");
                var Recirculadoresult18 = (uint)oPLC.Read("DB1010.DBD10172");
                var Recirculadoresult19 = (uint)oPLC.Read("DB1010.DBD10186");
                var Recirculadoresult20 = (uint)oPLC.Read("DB1010.DBD10180");
                var Recirculadoresult21 = (uint)oPLC.Read("DB1010.DBD10184");
                var Recirculadoresult22 = (uint)oPLC.Read("DB1010.DBD10188");
                var Recirculadoresult23 = (uint)oPLC.Read("DB1010.DBD10192");

                oRecirculado.CartonRecirc01 = Convert.ToInt32(Recirculadoresult1);
                oRecirculado.CartonRecirc02 = Convert.ToInt32(Recirculadoresult2);
                oRecirculado.CartonRecirc03 = Convert.ToInt32(Recirculadoresult3);
                oRecirculado.CartonRecirc04 = Convert.ToInt32(Recirculadoresult4);
                oRecirculado.CartonRecirc05 = Convert.ToInt32(Recirculadoresult5);
                oRecirculado.CartonRecirc06 = Convert.ToInt32(Recirculadoresult6);
                oRecirculado.CartonRecirc07 = Convert.ToInt32(Recirculadoresult7);
                oRecirculado.CartonRecirc08 = Convert.ToInt32(Recirculadoresult8);
                oRecirculado.CartonRecirc09 = Convert.ToInt32(Recirculadoresult9);
                oRecirculado.CartonRecirc10 = Convert.ToInt32(Recirculadoresult10);
                oRecirculado.CartonRecirc11 = Convert.ToInt32(Recirculadoresult11);
                oRecirculado.CartonRecirc12 = Convert.ToInt32(Recirculadoresult12);
                oRecirculado.CartonRecirc13 = Convert.ToInt32(Recirculadoresult13);
                oRecirculado.CartonRecirc14 = Convert.ToInt32(Recirculadoresult14);
                oRecirculado.CartonRecirc15 = Convert.ToInt32(Recirculadoresult15);
                oRecirculado.CartonRecirc16 = Convert.ToInt32(Recirculadoresult16);
                oRecirculado.CartonRecirc17 = Convert.ToInt32(Recirculadoresult17);
                oRecirculado.CartonRecirc18 = Convert.ToInt32(Recirculadoresult18);
                oRecirculado.CartonRecirc19 = Convert.ToInt32(Recirculadoresult19);
                oRecirculado.CartonRecirc20 = Convert.ToInt32(Recirculadoresult20);
                oRecirculado.CartonRecirc21 = Convert.ToInt32(Recirculadoresult21);
                oRecirculado.CartonRecirc22 = Convert.ToInt32(Recirculadoresult22);
                oRecirculado.CartonRecirc23 = Convert.ToInt32(Recirculadoresult23);

                string res = string.Empty;
                PropertyInfo[] properties = typeof(cEntRecirculado).GetProperties();
                foreach (PropertyInfo property in properties)
                {
                    //así obtenemos el nombre del atributo
                    string NombreAtributo = property.Name;

                    //así obtenemos el valor del atributo
                    string Valor = property.GetValue(oRecirculado).ToString();

                    string loc_conversion = string.Empty;

                    loc_conversion = Regex.Replace(Valor, @"[^\w\s.!@$%^&*()\-\/]+", "");
    
                    in_IdSalida = Convert.ToInt32(loc_conversion);
                    in_LineaSalida = Convert.ToInt32(NombreAtributo.Substring(12, 2));
                    in_Tipo = 0;

                    //if (conver != string.Empty)
                    if (loc_conversion != "0")
                    {
                        res = CartonRecirculados(in_IdSalida ,in_LineaSalida, in_Tipo);

                        if (res == "1")
                        {
                            LogRegistro.WriteEntry("Se actualizo el movimiento recirculado -------------> " + in_IdSalida.ToString());  
                        }
                        else
                        {
                            
                        }
                    }
                }

                //limpio las variables de recepcion de datos
                oRecirculado = new cEntRecirculado();

            }
            catch (Exception ex)
            {
                LogRegistro.WriteEntry(ex.Message.ToString() + " funcion LeerEtiquetas Recirculadas");
            }

        }

        #region BASE DATOS

        public string CartonDerivado(int in_mov,int in_linea, int in_tipo)
        {
            int res = 0;
            try
            {
                SqlConnection cnx = new SqlConnection(ConfigurationManager.ConnectionStrings["cnnString"].ToString());
                SqlCommand cmd = new SqlCommand("sp_Actualiza_CartonLPN_SERVICIO", cnx);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.Add("@pid_Movimiento", System.Data.SqlDbType.Int).Value = in_mov ;
                cmd.Parameters.Add("@pLinea", System.Data.SqlDbType.Int).Value = in_linea;
                cmd.Parameters.Add("@pTipoMov", System.Data.SqlDbType.Int).Value = in_tipo;

                //parametros de salida
                cmd.Parameters.Add("@pSalida", SqlDbType.Int).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("@pMensaje", SqlDbType.VarChar, 255).Direction = ParameterDirection.Output;

                cnx.Open();
                res = cmd.ExecuteNonQuery();

                int Salida = Convert.ToInt32(cmd.Parameters["@pSalida"].Value);
                string Mensaje = cmd.Parameters["@pMensaje"].Value.ToString();

                cnx.Close();
                cnx.Dispose();

                if (Salida == 1)
                {
                    return "1";
                }
                else
                {
                    return Mensaje;
                }

            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }


        }

        public string CartonRecirculados(int in_mov, int in_linea, int in_tipo)
        {
            int res = 0;
            try
            {
                SqlConnection cnx = new SqlConnection(ConfigurationManager.ConnectionStrings["cnnString"].ToString());
                SqlCommand cmd = new SqlCommand("sp_Actualiza_CartonLPN_SERVICIO", cnx);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.Add("@pid_Movimiento", System.Data.SqlDbType.Int).Value = in_mov;
                cmd.Parameters.Add("@pLinea", System.Data.SqlDbType.Int).Value = in_linea;
                cmd.Parameters.Add("@pTipoMov", System.Data.SqlDbType.Int).Value = in_tipo;

                //parametros de salida
                cmd.Parameters.Add("@pSalida", SqlDbType.Int).Direction = ParameterDirection.Output;
                cmd.Parameters.Add("@pMensaje", SqlDbType.VarChar, 255).Direction = ParameterDirection.Output;

                cnx.Open();
                res = cmd.ExecuteNonQuery();

                int Salida = Convert.ToInt32(cmd.Parameters["@pSalida"].Value);
                string Mensaje = cmd.Parameters["@pMensaje"].Value.ToString();

                cnx.Close();
                cnx.Dispose();

                if (Salida == 1)
                {
                    return "1";
                }
                else
                {
                    return Mensaje;
                }

            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }


        }

        #endregion


    }
}

using EntidadesCampanasPPG.Modelo;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EntidadesCampanasPPG.BD;

namespace DatosCampanasPPG.Campana
{
    public class CampanaDAT : BD
    {
        public int GuardarCampana(ref int ID_Camp, string Camp_Number, string Nombre_Camp, DateTime FechaAlta,
                int Negocio_Lider, string LiderCampania, string PPG_ID_Lider, int SubCanal, int Moneda,
                bool Express, int TipoCampania, int Alcance, string ClientesOtrosCanales, DateTime Fecha_Inicio_SubCanal,
                DateTime Fecha_Fin_SubCanal, DateTime FechaInicioPublico, DateTime FechaFinPublico,
                string PPG_ID, string NombreUsuario, int Estatus_ID, int ID_TipoCamp, int ID_TipoSell,
                string Objetivo, string JustificacionTyp, int ID_Familia, decimal DF, decimal DF6,
                decimal DF7, decimal DF9, decimal DF11, decimal DF17, string Justificacion)
        {
            const string spName = "InsertInfoCampania";
            int resultado = 0;
            SqlParameter SqlParameterResult;
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                //DATOS CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Camp_Number";
                Parametro.Value = Camp_Number;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Nombre_Camp";
                Parametro.Value = Nombre_Camp;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaAlta";
                Parametro.Value = FechaAlta;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Negocio_Lider";
                Parametro.Value = Negocio_Lider;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@LiderCampania";
                Parametro.Value = LiderCampania;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID_Lider";
                Parametro.Value = PPG_ID_Lider;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@SubCanal";
                Parametro.Value = SubCanal;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Moneda";
                Parametro.Value = Moneda;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Express";
                Parametro.Value = Express;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@TipoCampania";
                Parametro.Value = TipoCampania;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Alcance";
                Parametro.Value = Alcance;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ClientesOtrosCanales";
                Parametro.Value = ClientesOtrosCanales;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Inicio_SubCanal";
                Parametro.Value = Fecha_Inicio_SubCanal;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Fin_SubCanal";
                Parametro.Value = Fecha_Fin_SubCanal;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaInicioPublico";
                Parametro.Value = FechaInicioPublico;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaFinPublico";
                Parametro.Value = FechaFinPublico;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID";
                Parametro.Value = PPG_ID;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@NombreUsuario";
                Parametro.Value = NombreUsuario;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Estatus_ID";
                Parametro.Value = Estatus_ID;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoCamp";
                Parametro.Value = ID_TipoCamp;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoSell";
                Parametro.Value = ID_TipoSell;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Objetivo";
                Parametro.Value = Objetivo;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@JustificacionTyp";
                Parametro.Value = JustificacionTyp;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campc";
                Parametro.SqlDbType = SqlDbType.Int;
                Parametro.Direction = ParameterDirection.Output;
                ListParametros.Add(Parametro);

                //DATOS KROMA

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Familia";
                Parametro.Value = ID_Familia;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF";
                Parametro.Value = DF;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF6";
                Parametro.Value = DF6;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF7";
                Parametro.Value = DF7;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF9";
                Parametro.Value = DF9;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF11";
                Parametro.Value = DF11;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF17";
                Parametro.Value = DF17;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Justificacion";
                Parametro.Value = Justificacion;
                ListParametros.Add(Parametro);

                resultado = base.ExecuteNonQuery(spName, ListParametros);

                if (resultado > 0)
                {
                    SqlParameterResult = ListParametros.Where(n => n.ParameterName == "@ID_Campc").FirstOrDefault();

                    if (SqlParameterResult != null)
                    {
                        ID_Camp = Convert.ToInt32(SqlParameterResult.Value);
                    }
                }

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int GuardarCampanaKroma(int ID_Campana, int ID_Familia, decimal DF, decimal DF6, decimal DF7, decimal DF9, decimal DF11,
                        decimal DF17, string Justificacion)
        {
            const string spName = "InsertParticipacionKroma";
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campaña";
                Parametro.Value = ID_Campana;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Familia";
                Parametro.Value = ID_Familia;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF";
                Parametro.Value = DF;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF6";
                Parametro.Value = DF6;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF7";
                Parametro.Value = DF7;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF9";
                Parametro.Value = DF9;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF11";
                Parametro.Value = DF11;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@DF17";
                Parametro.Value = DF17;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Justificacion";
                Parametro.Value = Justificacion;
                ListParametros.Add(Parametro);

                return base.ExecuteNonQuery(spName, ListParametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int GuardarCampanaListaPublicidad(int ID_Campana, int ID_Publicidad, decimal Monto)
        {

            const string spName = "InsertPublicidad";
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campana";
                Parametro.Value = ID_Campana;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Publicidad";
                Parametro.Value = ID_Publicidad;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Monto";
                Parametro.Value = Monto;
                ListParametros.Add(Parametro);

                return base.ExecuteNonQuery(spName, ListParametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int GuardarCampanaListaProducto(int ID_Campana, int ID_Mecanica, int ID_LineaProducto, int ID_FamiliaProducto,
                        decimal Descuento, string Unidad_Medida, int ID_AlcanceEspecifico, int ID_Roles, string Comentario,
                        string Sistema_Aplicacion)
        {
            const string spName = "InsertEscenario";
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campana";
                Parametro.Value = ID_Campana;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Mecanica";
                Parametro.Value = ID_Mecanica;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_LineaProducto";
                Parametro.Value = ID_LineaProducto;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_FamiliaProducto";
                Parametro.Value = ID_Campana;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Descuento";
                Parametro.Value = Descuento;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Unidad_Medida";
                Parametro.Value = Unidad_Medida;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_AlcanceEspecifico";
                Parametro.Value = ID_AlcanceEspecifico;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Roles";
                Parametro.Value = ID_Roles;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Comentario";
                Parametro.Value = Comentario;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Sistema_Aplicacion";
                Parametro.Value = Sistema_Aplicacion;
                ListParametros.Add(Parametro);

                return base.ExecuteNonQuery(spName, ListParametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int GuardarCampanaCompleto(EntidadesCampanasPPG.Modelo.Campana Campana)
        {            
            const string spNameCampana = "MostrarIDCampania";
            const string spNameUpdateCampana = "UpdateInfoCampania";
            const string spNamePublicidad = "InsertPublicidad";
            const string spNameEliminarPublicidad = "DeletePublicidad";
            const string spNameKroma = "InsertParticipacionKroma";
            const string spNameEliminarKroma = "DeleteParticipacionKroma";
            //const string spNameProducto = "InsertEscenario";
            const string spNameSignoVital = "InsertCampaniaSignoVital";
            const string spNameEliminarSignoVital = "DeleteCampaniaSignoVital";

            int resultado = 0;
            List<SqlParameter> ListParametrosCampana = new List<SqlParameter>();
            List<SqlParameter> ListParametrosUpdateCampana = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosPublicidadGrupo = new List<SqlParameterGroup>();
            List<SqlParameter> ListParametrosPublicidad = new List<SqlParameter>();
            List<SqlParameter> ListParametrosEliminarPublicidad = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosKromaGrupo = new List<SqlParameterGroup>();
            List<SqlParameter> ListParametrosKroma = new List<SqlParameter>();
            List<SqlParameter> ListParametrosEliminarKroma = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosSignoVitalGrupo = new List<SqlParameterGroup>();
            List<SqlParameter> ListParametrosSignoVital = new List<SqlParameter>();
            List<SqlParameter> ListParametrosEliminarSignoVital = new List<SqlParameter>();

            SqlParameterGroup ParameterGroup;
            SqlParameter Parametro;

            try
            {
                IFormatProvider culture = new CultureInfo("es-MX", true);

                //DATOS MOSTRAR CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Clv_Campania";
                Parametro.Value = Campana.Title;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campc";
                Parametro.SqlDbType = SqlDbType.Int;
                Parametro.Direction = ParameterDirection.Output;
                ListParametrosCampana.Add(Parametro);


                //DateTime fechaDocumento = new DateTime();
                //fechaDocumento = DateTime.ParseExact(Campana.FechaDocumento, "dd/MM/yyyy", culture);

                //DateTime fechaInicioSubCanal = new DateTime();
                //fechaInicioSubCanal = DateTime.ParseExact(Campana.FechaInicioSubCanal, "dd/MM/yyyy", culture);

                //DateTime fechaFinSubCanal = new DateTime();
                //fechaFinSubCanal = DateTime.ParseExact(Campana.FechaFinSubCanal, "dd/MM/yyyy", culture);

                //DateTime fechaInicioPublico = new DateTime();
                //fechaInicioPublico = DateTime.ParseExact(Campana.FechaInicioPublico, "dd/MM/yyyy", culture);

                //DateTime fechaFinPublico = new DateTime();
                //fechaFinPublico = DateTime.ParseExact(Campana.FechaFinPublico, "dd/MM/yyyy", culture);

                //DATOS ACTUALIZAR CAMPANA
                //Parametro = new SqlParameter();
                //Parametro.ParameterName = "@ID_Campania";
                //Parametro.Value = Campana.IdCampana;
                //ListParametrosUpdateCampana.Add(Parametro);

                //Parametro = new SqlParameter();
                //Parametro.ParameterName = "@Camp_Number";
                //Parametro.Value = Campana.Title;
                //ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Nombre_Camp";
                Parametro.Value = Campana.NombreCampa;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaAlta";
                if (Campana.FechaDocumento != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaDocumento, "dd/MM/yyyy", culture);
                }
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Negocio_Lider";
                Parametro.Value = Campana.IdNegocioLider;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@LiderCampania";
                Parametro.Value = Campana.LiderCampa;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID_Lider";
                Parametro.Value = Campana.PPGIDLiderCampa;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@SubCanal";
                Parametro.Value = Campana.IdSubcanal;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Moneda";
                Parametro.Value = Campana.IdMoneda;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Express";
                Parametro.Value = Campana.CampaExpress;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@TipoCampania";
                Parametro.Value = Campana.IdTipoCampa;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Alcance";
                Parametro.Value = Campana.IdAlcanceTerritorial;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ClientesOtrosCanales";
                Parametro.Value = null;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Inicio_SubCanal";
                if (Campana.FechaInicioSubCanal != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaInicioSubCanal, "dd/MM/yyyy", culture);
                }
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Fin_SubCanal";
                if (Campana.FechaFinSubCanal != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaFinSubCanal, "dd/MM/yyyy", culture);
                }
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaInicioPublico";
                if (Campana.FechaInicioPublico != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaInicioPublico, "dd/MM/yyyy", culture);
                }
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaFinPublico";
                if (Campana.FechaFinPublico != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaFinPublico, "dd/MM/yyyy", culture);
                }
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID";
                Parametro.Value = Campana.PPGIDRegistraCampa;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@NombreUsuario";
                Parametro.Value = Campana.RegistraCampa;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Estatus";
                Parametro.Value = Campana.Status;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoCamp";
                Parametro.Value = Campana.IdTipoCampa;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoSell";
                Parametro.Value = Campana.IdTipoSell;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Objetivo";
                Parametro.Value = Campana.ObjetivoNC;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@JustificacionTyp";
                Parametro.Value = Campana.JustificacionNC;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ComentarioVenta";
                Parametro.Value = Campana.ComentarioVenta;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@CampaniaAnterior";
                Parametro.Value = Campana.CampaniaAnterior;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Camp_Number_Ant";
                Parametro.Value = Campana.TitleAnterior;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Nombre_Camp_Ant";
                Parametro.Value = Campana.NombreCampaAnterior;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@JustificacionKroma";
                Parametro.Value = Campana.JustificacionKroma;
                ListParametrosUpdateCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@JustificacionObjetivo";
                Parametro.Value = Campana.ObjetivoKroma;
                ListParametrosUpdateCampana.Add(Parametro);

                //Parametro = new SqlParameter();
                //Parametro.ParameterName = "@TipoSubCanal";
                //Parametro.Value = Campana.TipoSubCanal;
                //ListParametrosUpdateCampana.Add(Parametro);

                //DATOS SIGNO BITAL
                foreach (SignoVital signoVital in Campana.ListaSignoVital)
                {
                    ListParametrosSignoVital = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SignoVital";
                    Parametro.Value = signoVital.Descripcion;
                    ListParametrosSignoVital.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosSignoVital;

                    ListParametrosSignoVitalGrupo.Add(ParameterGroup);
                }

                //DATOS KROMA
                foreach (Kroma kroma in Campana.ListaKroma)
                {
                    ListParametrosKroma = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Base";
                    Parametro.Value = kroma.Base;
                    ListParametrosKroma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Linea";
                    Parametro.Value = kroma.Linea;
                    ListParametrosKroma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SobrePrecioKroma";
                    Parametro.Value = kroma.SobrePrecioKroma;
                    ListParametrosKroma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Porcentaje";
                    Parametro.Value = kroma.Porcentaje;
                    ListParametrosKroma.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosKroma;

                    ListParametrosKromaGrupo.Add(ParameterGroup);
                }

                //DATOS PUBLICIDAD
                foreach (Publicidad publicidad in Campana.ListaPublicidad)
                {
                    ListParametrosPublicidad = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Publicidad";
                    Parametro.Value = publicidad.IdPublicidad;
                    ListParametrosPublicidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Monto";
                    Parametro.Value = publicidad.Monto;
                    ListParametrosPublicidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@MontoAnterior";
                    Parametro.Value = publicidad.MontoAnterior;
                    ListParametrosPublicidad.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosPublicidad;

                    ListParametrosPublicidadGrupo.Add(ParameterGroup);
                }

                //foreach (Producto producto in Campana.ListaProducto)
                //{
                //    ListParametrosProducto = new List<SqlParameter>();

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@ID_Mecanica";
                //    Parametro.Value = producto.IdMecanica;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@ID_LineaProducto";
                //    Parametro.Value = producto.IdLineaProducto;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@ID_FamiliaProducto";
                //    Parametro.Value = producto.IdFamiliaEstelar;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@Descuento";
                //    Parametro.Value = producto.CantidadOdescuento;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@Unidad_Medida";
                //    Parametro.Value = producto.CapacidadProducto;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@ID_AlcanceEspecifico";
                //    Parametro.Value = producto.Alcance;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@ID_Roles";
                //    Parametro.Value = producto.IdRol;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@Comentario";
                //    Parametro.Value = producto.Observaciones;
                //    ListParametrosProducto.Add(Parametro);

                //    Parametro = new SqlParameter();
                //    Parametro.ParameterName = "@Sistema_Aplicacion";
                //    Parametro.Value = producto.SistemaAplicacion;
                //    ListParametrosProducto.Add(Parametro);

                //    ParameterGroup = new SqlParameterGroup();
                //    ParameterGroup.ListSqlParameter = ListParametrosProducto;

                //    ListParametrosProductoGrupo.Add(ParameterGroup);
                //}

                resultado = base.ExecuteNonQueryCampana(spNameCampana, ListParametrosCampana, 
                                            spNameUpdateCampana, ListParametrosUpdateCampana,
                                            spNamePublicidad, ListParametrosPublicidadGrupo,
                                            spNameEliminarPublicidad, ListParametrosEliminarPublicidad, 
                                            spNameKroma, ListParametrosKromaGrupo, 
                                            spNameEliminarKroma, ListParametrosEliminarKroma,
                                            spNameSignoVital, ListParametrosSignoVitalGrupo,
                                            spNameEliminarSignoVital, ListParametrosEliminarSignoVital);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int GuardarCampanaCronogramaCompleto(EntidadesCampanasPPG.Modelo.Campana Campana)
        {
            const string spNameCampana = "InsertInfoCampania";
            const string spNameCronograma = "InsertCronograma";
            const string spNameRegion = "InsertRegion";
            const string spNameRegionEliminar = "DeleteRegion";
            const string spNameSubCanal = "InsertSubCanal";
            const string spNameSubCanalEliminar = "DeleteSubCanal";

            int resultado = 0;
            List<SqlParameter> ListParametrosCampana = new List<SqlParameter>();

            List<SqlParameter> ListParametrosCronograma = new List<SqlParameter>();
            List<SqlParameterGroup> ListParametrosCronogramaGrupo = new List<SqlParameterGroup>();

            List<SqlParameter> ListParametrosRegion = new List<SqlParameter>();
            List<SqlParameterGroup> ListParametrosRegionGrupo = new List<SqlParameterGroup>();

            List<SqlParameter> ListParametrosRegionEliminar = new List<SqlParameter>();

            List<SqlParameter> ListParametrosSubCanal = new List<SqlParameter>();
            List<SqlParameterGroup> ListParametrosSubCanalGrupo = new List<SqlParameterGroup>();

            List<SqlParameter> ListParametrosSubCanalEliminar = new List<SqlParameter>();

            SqlParameterGroup ParameterGroup;

            SqlParameter Parametro;

            try
            {
                IFormatProvider culture = new CultureInfo("es-MX", true);

                //DATOS CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Camp_Number";
                Parametro.Value = Campana.Title;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Nombre_Camp";
                Parametro.Value = Campana.NombreCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaAlta";
                if (Campana.FechaDocumento != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaDocumento, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Negocio_Lider";
                Parametro.Value = Campana.IdNegocioLider;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@LiderCampania";
                Parametro.Value = Campana.LiderCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID_Lider";
                Parametro.Value = Campana.PPGIDLiderCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@SubCanal";
                Parametro.Value = Campana.IdSubcanal;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Moneda";
                Parametro.Value = Campana.IdMoneda;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Express";
                Parametro.Value = Campana.CampaExpress;
                ListParametrosCampana.Add(Parametro);

                //Parametro = new SqlParameter();
                //Parametro.ParameterName = "@TipoCampania";
                //Parametro.Value = Campana.IdTipoCampa;
                //ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Alcance";
                Parametro.Value = Campana.IdAlcanceTerritorial;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Inicio_SubCanal";
                if (Campana.FechaInicioSubCanal != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaInicioSubCanal, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Fin_SubCanal";
                if (Campana.FechaFinSubCanal != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaFinSubCanal, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaInicioPublico";
                if (Campana.FechaInicioPublico != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaInicioPublico, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaFinPublico";
                if (Campana.FechaFinPublico != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaFinPublico, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID";
                Parametro.Value = Campana.PPGIDRegistraCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@NombreUsuario";
                Parametro.Value = Campana.RegistraCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Estatus";
                Parametro.Value = Campana.Status;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoCamp";
                Parametro.Value = Campana.IdTipoCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoSell";
                Parametro.Value = Campana.IdTipoSell;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Objetivo";
                Parametro.Value = Campana.ObjetivoNC;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@JustificacionTyp";
                Parametro.Value = Campana.JustificacionNC;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ImporteNotaCredito";
                Parametro.Value = Campana.ImporteNotaCredito;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@TipoSubCanal";
                Parametro.Value = Campana.TipoSubCanal;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campc";
                Parametro.SqlDbType = SqlDbType.Int;
                Parametro.Direction = ParameterDirection.Output;
                ListParametrosCampana.Add(Parametro);

                if(Campana.ListaCronograma == null)
                {
                    Campana.ListaCronograma = new List<Cronograma>();
                }
                
                //DATOS CRONOGRAMA
                foreach (Cronograma cronograma in Campana.ListaCronograma)
                {
                    //if(cronograma.IDTarea == 3)
                    //{
                    //    cronograma.PorcentajeUsuario = 100;
                    //    cronograma.PorcentajeSistema = 100;
                    //    cronograma.FechaInicio = DateTime.Now;
                    //    cronograma.FechaFin = DateTime.Now;
                    //}

                    ListParametrosCronograma = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Tarea";
                    Parametro.Value = cronograma.IDTarea;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Padre";
                    Parametro.Value = cronograma.IDPadre;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Padre";
                    Parametro.Value = cronograma.Padre;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Actividad";
                    Parametro.Value = cronograma.Actividad;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Duracion";
                    Parametro.Value = cronograma.Duracion;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaInicio";
                    Parametro.Value = cronograma.FechaInicio;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaFin";
                    Parametro.Value = cronograma.FechaFin;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TiempoOptimista";
                    Parametro.Value = cronograma.TiempoOptimista;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TiempoPesimista";
                    Parametro.Value = cronograma.TiempoPesimista;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@NombreResponsable";
                    Parametro.Value = cronograma.NombreResponsable;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Correo";
                    Parametro.Value = cronograma.Correo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID";
                    Parametro.Value = cronograma.PPGID;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@NombreResponsable_2";
                    Parametro.Value = cronograma.NombreResponsable_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Correo_2";
                    Parametro.Value = cronograma.Correo_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID_2";
                    Parametro.Value = cronograma.PPGID_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Comentario";
                    Parametro.Value = cronograma.Comentario;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PorcentajeSis";
                    Parametro.Value = cronograma.PorcentajeSistema;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PorcentajeUsu";
                    Parametro.Value = cronograma.PorcentajeUsuario;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Predecesor";
                    Parametro.Value = cronograma.Predecesor;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@UsuarioCreacion";
                    Parametro.Value = cronograma.UsuarioCreacion;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@UsuarioModificacion";
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TipoFlujo";
                    Parametro.Value = cronograma.TipoFlujo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IdTipoFlujo";
                    Parametro.Value = cronograma.IdTipoFlujo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Incluir";
                    Parametro.Value = cronograma.Incluir;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IdIncluir";
                    Parametro.Value = cronograma.IdIncluir;
                    ListParametrosCronograma.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosCronograma;

                    ListParametrosCronogramaGrupo.Add(ParameterGroup);
                }

                if(Campana.ListaRegion == null)
                {
                    Campana.ListaRegion = new List<Region>();
                }

                //DATOS REGION
                foreach (Region region in Campana.ListaRegion)
                {
                    ListParametrosRegion = new List<SqlParameter>();

                    //Parametro = new SqlParameter();
                    //Parametro.ParameterName = "@ID_Campania";
                    //Parametro.Value = region.IDCampania;
                    //ListParametrosRegion.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Region";
                    Parametro.Value = region.NombreRegion;
                    ListParametrosRegion.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID";
                    Parametro.Value = region.PPGID;
                    ListParametrosRegion.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Nombre_Usuario";
                    Parametro.Value = region.Usuario;
                    ListParametrosRegion.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosRegion;

                    ListParametrosRegionGrupo.Add(ParameterGroup);
                }

                if(Campana.ListaSubCanal == null)
                {
                    Campana.ListaSubCanal = new List<SubCanal>();
                }

                //DATOS SUBCANAL
                foreach (SubCanal subCanal in Campana.ListaSubCanal)
                {
                    ListParametrosSubCanal = new List<SqlParameter>();

                    //Parametro = new SqlParameter();
                    //Parametro.ParameterName = "@ID_SubCanal";
                    //Parametro.Value = subCanal.IdSubCanal;
                    //ListParametrosSubCanal.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SubCanal";
                    Parametro.Value = subCanal.Descripcion;
                    ListParametrosSubCanal.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID";
                    Parametro.Value = subCanal.PPGID;
                    ListParametrosSubCanal.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Nombre_Usuario";
                    Parametro.Value = subCanal.NombreUsuario;
                    ListParametrosSubCanal.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosSubCanal;

                    ListParametrosSubCanalGrupo.Add(ParameterGroup);
                }


                resultado = base.ExecuteNonQueryCampana(spNameCampana, ListParametrosCampana, 
                                                        spNameCronograma, ListParametrosCronogramaGrupo,
                                                        spNameRegionEliminar, ListParametrosRegionEliminar,
                                                        spNameRegion, ListParametrosRegionGrupo,
                                                        spNameSubCanalEliminar, ListParametrosRegionEliminar,
                                                        spNameSubCanal, ListParametrosSubCanalGrupo);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int GuardarCronogramaCompleto(EntidadesCampanasPPG.Modelo.Campana Campana)
        {
            const string spNameCronograma = "InsertCronogramaClaveCamp";

            int resultado = 0;

            List<SqlParameter> ListParametrosCronograma = new List<SqlParameter>();
            List<SqlParameterGroup> ListParametrosCronogramaGrupo = new List<SqlParameterGroup>();

            SqlParameterGroup ParameterGroup;
            SqlParameter Parametro;

            try
            {
                //DATOS CRONOGRAMA
                foreach (Cronograma cronograma in Campana.ListaCronograma)
                {

                    ListParametrosCronograma = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ClaveCampania";
                    Parametro.Value = Campana.Title;
                    ListParametrosCronograma.Add(Parametro);              

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Tarea";
                    Parametro.Value = cronograma.IDTarea;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Padre";
                    Parametro.Value = cronograma.IDPadre;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Padre";
                    Parametro.Value = cronograma.Padre;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Actividad";
                    Parametro.Value = cronograma.Actividad;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Duracion";
                    Parametro.Value = cronograma.Duracion;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaInicio";
                    Parametro.Value = cronograma.FechaInicio;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaFin";
                    Parametro.Value = cronograma.FechaFin;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TiempoOptimista";
                    Parametro.Value = cronograma.TiempoOptimista;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TiempoPesimista";
                    Parametro.Value = cronograma.TiempoPesimista;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@NombreResponsable";
                    Parametro.Value = cronograma.NombreResponsable;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Correo";
                    Parametro.Value = cronograma.Correo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID";
                    Parametro.Value = cronograma.PPGID;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@NombreResponsable_2";
                    Parametro.Value = cronograma.NombreResponsable_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Correo_2";
                    Parametro.Value = cronograma.Correo_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID_2";
                    Parametro.Value = cronograma.PPGID_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Comentario";
                    Parametro.Value = cronograma.Comentario;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PorcentajeSis";
                    Parametro.Value = cronograma.PorcentajeSistema;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PorcentajeUsu";
                    Parametro.Value = cronograma.PorcentajeUsuario;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Predecesor";
                    Parametro.Value = cronograma.Predecesor;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@UsuarioCreacion";
                    Parametro.Value = cronograma.UsuarioCreacion;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@UsuarioModificacion";
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TipoFlujo";
                    Parametro.Value = cronograma.TipoFlujo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IdTipoFlujo";
                    Parametro.Value = cronograma.IdTipoFlujo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Incluir";
                    Parametro.Value = cronograma.Incluir;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IdIncluir";
                    Parametro.Value = cronograma.IdIncluir;
                    ListParametrosCronograma.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosCronograma;

                    ListParametrosCronogramaGrupo.Add(ParameterGroup);
                }

                resultado = base.ExecuteNonQueryTrans(spNameCronograma, ListParametrosCronogramaGrupo);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int EditCampanaCronogramaCompleto(EntidadesCampanasPPG.Modelo.Campana Campana)
        {
            const string spNameCampana = "UpdateInfoCampania";
            const string spNameCronograma = "UpdateCronogramaNuevaVersion";
            const string spNameRegion = "InsertRegion";
            const string spNameRegionEliminar = "DeleteRegion";
            const string spNameSubCanal = "InsertSubCanal";
            const string spNameSubCanalEliminar = "DeleteSubCanal";

            int resultado = 0;
            List<SqlParameter> ListParametrosCampana = new List<SqlParameter>();

            List<SqlParameter> ListParametrosCronograma = new List<SqlParameter>();
            List<SqlParameterGroup> ListParametrosCronogramaGrupo = new List<SqlParameterGroup>();

            List<SqlParameter> ListParametrosRegion = new List<SqlParameter>();
            List<SqlParameterGroup> ListParametrosRegionGrupo = new List<SqlParameterGroup>();

            List<SqlParameter> ListParametrosRegionEliminar = new List<SqlParameter>();

            List<SqlParameter> ListParametrosSubCanal = new List<SqlParameter>();
            List<SqlParameterGroup> ListParametrosSubCanalGrupo = new List<SqlParameterGroup>();

            List<SqlParameter> ListParametrosSubCanalEliminar = new List<SqlParameter>();

            SqlParameterGroup ParameterGroup;

            SqlParameter Parametro;

            try
            {
                IFormatProvider culture = new CultureInfo("es-MX", true);

                //DATOS CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Camp_Number";
                Parametro.Value = Campana.Title;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Nombre_Camp";
                Parametro.Value = Campana.NombreCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaAlta";
                if (Campana.FechaDocumento != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaDocumento, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Negocio_Lider";
                Parametro.Value = Campana.IdNegocioLider;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@LiderCampania";
                Parametro.Value = Campana.LiderCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID_Lider";
                Parametro.Value = Campana.PPGIDLiderCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@SubCanal";
                Parametro.Value = Campana.IdSubcanal;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Moneda";
                Parametro.Value = Campana.IdMoneda;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Express";
                Parametro.Value = Campana.CampaExpress;
                ListParametrosCampana.Add(Parametro);

                //Parametro = new SqlParameter();
                //Parametro.ParameterName = "@TipoCampania";
                //Parametro.Value = Campana.IdTipoCampa;
                //ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Alcance";
                Parametro.Value = Campana.IdAlcanceTerritorial;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Inicio_SubCanal";
                if (Campana.FechaInicioSubCanal != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaInicioSubCanal, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Fecha_Fin_SubCanal";
                if (Campana.FechaFinSubCanal != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaFinSubCanal, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaInicioPublico";
                if (Campana.FechaInicioPublico != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaInicioPublico, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaFinPublico";
                if (Campana.FechaFinPublico != null)
                {
                    Parametro.Value = DateTime.ParseExact(Campana.FechaFinPublico, "dd/MM/yyyy", culture);
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPG_ID";
                Parametro.Value = Campana.PPGIDRegistraCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@NombreUsuario";
                Parametro.Value = Campana.RegistraCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Estatus";
                Parametro.Value = Campana.Status;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoCamp";
                Parametro.Value = Campana.IdTipoCampa;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_TipoSell";
                Parametro.Value = Campana.IdTipoSell;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Objetivo";
                Parametro.Value = Campana.ObjetivoNC;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@JustificacionTyp";
                Parametro.Value = Campana.JustificacionNC;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campc";
                Parametro.SqlDbType = SqlDbType.Int;
                Parametro.Direction = ParameterDirection.Output;
                ListParametrosCampana.Add(Parametro);

                //DATOS CRONOGRAMA
                foreach (Cronograma cronograma in Campana.ListaCronograma)
                {
                    //if(cronograma.IDTarea == 3)
                    //{
                    //    cronograma.PorcentajeUsuario = 100;
                    //    cronograma.PorcentajeSistema = 100;
                    //    cronograma.FechaInicio = DateTime.Now;
                    //    cronograma.FechaFin = DateTime.Now;
                    //}

                    ListParametrosCronograma = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IDTarea";
                    Parametro.Value = cronograma.IDTarea;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IDPadre";
                    Parametro.Value = cronograma.IDPadre;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Padre";
                    Parametro.Value = cronograma.Padre;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Actividad";
                    Parametro.Value = cronograma.Actividad;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Duracion";
                    Parametro.Value = cronograma.Duracion;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaInicio";
                    Parametro.Value = cronograma.FechaInicio;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaFin";
                    Parametro.Value = cronograma.FechaFin;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TiempoOptimista";
                    Parametro.Value = cronograma.TiempoOptimista;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TiempoPesimista";
                    Parametro.Value = cronograma.TiempoPesimista;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@NombreResponsable";
                    Parametro.Value = cronograma.NombreResponsable;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Correo";
                    Parametro.Value = cronograma.Correo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID";
                    Parametro.Value = cronograma.PPGID;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@NombreResponsable_2";
                    Parametro.Value = cronograma.NombreResponsable_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Correo_2";
                    Parametro.Value = cronograma.Correo_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID_2";
                    Parametro.Value = cronograma.PPGID_2;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Comentario";
                    Parametro.Value = cronograma.Comentario;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PorcentajeSis";
                    Parametro.Value = cronograma.PorcentajeSistema;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PorcentajeUsu";
                    Parametro.Value = cronograma.PorcentajeUsuario;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Predecesor";
                    Parametro.Value = cronograma.Predecesor;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@UsuarioCreacion";
                    Parametro.Value = cronograma.UsuarioCreacion;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@UsuarioModificacion";
                    Parametro.Value = cronograma.UsuarioCreacion;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TipoFlujo";
                    Parametro.Value = cronograma.TipoFlujo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IdTipoFlujo";
                    Parametro.Value = cronograma.IdTipoFlujo;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Incluir";
                    Parametro.Value = cronograma.Incluir;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@IdIncluir";
                    Parametro.Value = cronograma.IdIncluir;
                    ListParametrosCronograma.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Version";
                    Parametro.Value = cronograma.VersionCronograma;
                    ListParametrosCronograma.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosCronograma;

                    ListParametrosCronogramaGrupo.Add(ParameterGroup);
                }

                //DATOS REGION
                foreach (Region region in Campana.ListaRegion)
                {
                    ListParametrosRegion = new List<SqlParameter>();

                    //Parametro = new SqlParameter();
                    //Parametro.ParameterName = "@ID_Campania";
                    //Parametro.Value = region.IDCampania;
                    //ListParametrosRegion.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Region";
                    Parametro.Value = region.NombreRegion;
                    ListParametrosRegion.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID";
                    Parametro.Value = region.PPGID;
                    ListParametrosRegion.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Nombre_Usuario";
                    Parametro.Value = region.Usuario;
                    ListParametrosRegion.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosRegion;

                    ListParametrosRegionGrupo.Add(ParameterGroup);
                }

                //DATOS SUBCANAL
                foreach (SubCanal subCanal in Campana.ListaSubCanal)
                {
                    ListParametrosSubCanal = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_SubCanal";
                    Parametro.Value = subCanal.IdSubCanal;
                    ListParametrosSubCanal.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SubCanal";
                    Parametro.Value = subCanal.Descripcion;
                    ListParametrosSubCanal.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PPGID";
                    Parametro.Value = subCanal.PPGID;
                    ListParametrosSubCanal.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Nombre_Usuario";
                    Parametro.Value = subCanal.NombreUsuario;
                    ListParametrosSubCanal.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametrosSubCanal;

                    ListParametrosSubCanalGrupo.Add(ParameterGroup);
                }

                resultado = base.ExecuteNonQueryCampana(spNameCampana, ListParametrosCampana, 
                                                        spNameCronograma, ListParametrosCronogramaGrupo,
                                                        spNameRegionEliminar, ListParametrosRegionEliminar,
                                                        spNameRegion, ListParametrosRegionGrupo,
                                                        spNameSubCanalEliminar, ListParametrosSubCanalEliminar,
                                                        spNameSubCanal, ListParametrosSubCanalGrupo);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<EntidadesCampanasPPG.Modelo.Campana> MostrarCampana(EntidadesCampanasPPG.Modelo.Campana Campana)
        {
            List<EntidadesCampanasPPG.Modelo.Campana> ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

            const string spNameCampana = "MostrarBitacoraCampania";

            SqlParameter Parametro = new SqlParameter();
            List<SqlParameter> ListParametrosCampana = new List<SqlParameter>();

            DataTable dtCampana = new DataTable();

            try
            {
                //DATOS CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@IDCampania";
                if(Campana.IdCampana > 0)
                {
                    Parametro.Value = Campana.IdCampana;
                }                
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ClvCampania";
                if(!string.IsNullOrEmpty(Campana.Title))
                {
                    Parametro.Value = Campana.Title;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@NombreCampania";
                if (!string.IsNullOrEmpty(Campana.NombreCampa))
                {
                    Parametro.Value = Campana.NombreCampa;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Lider_PPG";
                if(!string.IsNullOrEmpty(Campana.PPGIDLiderCampa))
                {
                    Parametro.Value = Campana.PPGIDLiderCampa;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Urgente";
                if (Campana.CampaExpress.ToString() != "")
                {
                    Parametro.Value = Campana.CampaExpress;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@IdTipoCampania";
                if(Campana.IdTipoCampa > 0)
                {
                    Parametro.Value = Campana.IdTipoCampa;
                }              
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@IdTipoVenta";
                if(Campana.IdTipoSell > 0)
                {
                    Parametro.Value = Campana.IdTipoSell;
                }               
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@IdAlcance";
                if(Campana.IdAlcanceTerritorial > 0)
                {
                    Parametro.Value = Campana.IdAlcanceTerritorial;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaInicioSubCanal";
                if(Campana.FechaInicioSubCanal != null && Campana.FechaInicioSubCanal != "01/01/1900")
                {
                    Parametro.Value = Campana.FechaInicioSubCanal;
                } 
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaFinSubCanal";
                if (Campana.FechaFinSubCanal != null && Campana.FechaFinSubCanal != "01/01/1900")
                {
                    Parametro.Value = Campana.FechaFinSubCanal;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaInicioPublico";
                if (Campana.FechaInicioPublico != null && Campana.FechaInicioPublico != "01/01/1900")
                {
                    Parametro.Value = Campana.FechaInicioPublico;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@FechaFinPublico";
                if (Campana.FechaFinPublico != null && Campana.FechaFinPublico != "01/01/1900")
                {
                    Parametro.Value = Campana.FechaFinPublico;
                }
                ListParametrosCampana.Add(Parametro);

                dtCampana = base.ExecuteDataTable(spNameCampana, ListParametrosCampana);

                ListCampana = dtCampana.AsEnumerable()
                                .Select(row => new EntidadesCampanasPPG.Modelo.Campana
                                {
                                    IdCampana = row.Field<int?>("ID").GetValueOrDefault(),
                                    Title = row.Field<string>("Camp_Number"),
                                    NombreCampa = row.Field<string>("Nombre_Camp"),
                                    RegistraCampa = row.Field<string>("Nombre_Usuario"),
                                    PPGIDRegistraCampa = row.Field<string>("PPGID"),
                                    LiderCampa = row.Field<string>("Lider_Campania"),
                                    PPGIDLiderCampa = row.Field<string>("PPGID_Lider"),
                                    FechaInicioPublico = row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaFinPublico = row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaDocumento = row.Field<DateTime?>("Fecha_Creacion").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Creacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    IdMoneda = row.Field<int?>("ID_Moneda").GetValueOrDefault(),
                                    Moneda = row.Field<string>("Moneda"),
                                    IdTipoCampa = row.Field<int?>("ID_TipoCamp").GetValueOrDefault(),
                                    TipoCampa = row.Field<string>("TipoCamp"),
                                    IdAlcanceTerritorial = row.Field<int?>("ID_Alcance").GetValueOrDefault(),
                                    AlcanceTerritorial = row.Field<string>("Alcance"),
                                    IdTipoSell = row.Field<int?>("ID_TipoSell").GetValueOrDefault(),
                                    TipoSell = row.Field<string>("TipoSell"),
                                    CampaExpress = row.Field<bool?>("Express").GetValueOrDefault(),
                                    IdNegocioLider = row.Field<int?>("ID_Negocio_Lider").GetValueOrDefault(),
                                    NegocioLider = row.Field<string>("LiderNegocio"),
                                    IdSubcanal = row.Field<int?>("ID_Subcanal").GetValueOrDefault(),
                                    Subcanal = row.Field<string>("Subcanal"),
                                    IdEstatus = row.Field<int?>("ID_Estatus").GetValueOrDefault(),
                                    Estatus = row.Field<string>("Estatus"),
                                    Status = row.Field<string>("EstatusCat"),
                                    TipoSubCanal = row.Field<string>("TipoSubCanal"),
                                }).ToList();
            }
            catch(Exception ex)
            {
                throw ex;
            }

            return ListCampana;
        }
        public List<EntidadesCampanasPPG.Modelo.Campana> MostrarCampanaDetAnioAnterior(EntidadesCampanasPPG.Modelo.Campana Campana)
        {
            const string spNameCampana = "MostrarDetCampaniaAnioAnterior";

            List<EntidadesCampanasPPG.Modelo.Campana> ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();
            List<Kroma> ListKroma = new List<Kroma>();
            List<Publicidad> ListPublicidad = new List<Publicidad>();
            List<Region> ListRegion = new List<Region>();
            List<SubCanal> ListSubCanal = new List<SubCanal>();
            List<SignoVital> ListSignoVital = new List<SignoVital>();

            SqlParameter Parametro = new SqlParameter();
            List<SqlParameter> ListParametrosCampana = new List<SqlParameter>();

            DataSet dsCampana = new DataSet();

            try
            {
                //DATOS CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@IDCampania";
                if (Campana.IdCampana > 0)
                {
                    Parametro.Value = Campana.IdCampana;
                }
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ClvCampania";
                Parametro.Value = Campana.Title;
                ListParametrosCampana.Add(Parametro);

                dsCampana = base.ExecuteDataSet(spNameCampana, ListParametrosCampana);

                if (dsCampana.Tables.Count > 0)
                {
                    //INFORMACION CAMPANA
                    if (dsCampana.Tables[0].Rows.Count > 0)
                    {
                        ListCampana = dsCampana.Tables[0].AsEnumerable()
                                .Select(row => new EntidadesCampanasPPG.Modelo.Campana
                                {
                                    IdCampana = row.Field<int?>("ID").GetValueOrDefault(),
                                    Title = row.Field<string>("Camp_Number"),
                                    NombreCampa = row.Field<string>("Nombre_Camp"),
                                    RegistraCampa = row.Field<string>("Nombre_Usuario"),
                                    PPGIDRegistraCampa = row.Field<string>("PPGID"),
                                    LiderCampa = row.Field<string>("Lider_Campania"),
                                    PPGIDLiderCampa = row.Field<string>("PPGID_Lider"),
                                    FechaInicioSubCanal = row.Field<DateTime?>("Fecha_Inicio_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Inicio_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaFinSubCanal = row.Field<DateTime?>("Fecha_Fin_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Fin_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaInicioPublico = row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaFinPublico = row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaDocumento = row.Field<DateTime?>("Fecha_Creacion").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Creacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    IdMoneda = row.Field<int?>("ID_Moneda").GetValueOrDefault(),
                                    Moneda = row.Field<string>("Moneda"),
                                    IdTipoCampa = row.Field<int?>("ID_TipoCamp").GetValueOrDefault(),
                                    TipoCampa = row.Field<string>("TipoCamp"),
                                    IdAlcanceTerritorial = row.Field<int?>("ID_Alcance").GetValueOrDefault(),
                                    AlcanceTerritorial = row.Field<string>("Alcance"),
                                    IdTipoSell = row.Field<int?>("ID_TipoSell").GetValueOrDefault(),
                                    TipoSell = row.Field<string>("TipoSell"),
                                    CampaExpress = row.Field<bool?>("Express").GetValueOrDefault(),
                                    IdNegocioLider = row.Field<int?>("ID_Negocio_Lider").GetValueOrDefault(),
                                    NegocioLider = row.Field<string>("LiderNegocio"),
                                    IdSubcanal = row.Field<int?>("ID_Subcanal").GetValueOrDefault(),
                                    Subcanal = row.Field<string>("Subcanal"),
                                    IdEstatus = row.Field<int?>("ID_Estatus").GetValueOrDefault(),
                                    Estatus = row.Field<string>("Estatus"),
                                    Status = row.Field<string>("EstatusCat"),
                                    JustificacionNC = row.Field<string>("Justificacion"),
                                    ObjetivoNC = row.Field<string>("Objetivo"),
                                    JustificacionKroma = row.Field<string>("JustificacionKroma"),
                                    ObjetivoKroma = row.Field<string>("JustificacionObjetivo"),
                                    TipoSubCanal = row.Field<string>("TipoSubCanal")
                                }).ToList();
                    }

                    //INFORMACION PARTICIPACION KROMA
                    if (dsCampana.Tables[1].Rows.Count > 0)
                    {
                        ListKroma = dsCampana.Tables[1].AsEnumerable()
                            .Select(row => new Kroma
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Base = row.Field<string>("Base"),
                                Linea = row.Field<string>("Linea"),
                                SobrePrecioKroma = row.Field<string>("SobrePrecioKroma"),
                                Porcentaje = row.Field<decimal?>("Porcentaje").GetValueOrDefault()
                            }).ToList();
                    }

                    //INFORMACION PUBLICIDAD
                    if (dsCampana.Tables[2].Rows.Count > 0)
                    {
                        ListPublicidad = dsCampana.Tables[2].AsEnumerable()
                            .Select(row => new Publicidad
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                IdPublicidad = row.Field<int?>("ID_Publicidad").GetValueOrDefault(),
                                PublicidadDescripcion = row.Field<string>("Publicidad"),
                                Comentario = row.Field<string>("Comentario"),
                                Monto = row.Field<decimal?>("Cantidad").GetValueOrDefault(),
                                MontoAnterior = row.Field<decimal?>("CantidadAnterior").GetValueOrDefault(),
                            }).ToList();
                    }

                    //INFORMACION REGION
                    if (dsCampana.Tables[3].Rows.Count > 0)
                    {
                        ListRegion = dsCampana.Tables[3].AsEnumerable()
                            .Select(row => new Region
                            {
                                IdCampania = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                NombreRegion = row.Field<string>("Region")
                            }).ToList();
                    }

                    //INFORMACION SUBCANAL
                    if (dsCampana.Tables[4].Rows.Count > 0)
                    {
                        ListSubCanal = dsCampana.Tables[4].AsEnumerable()
                            .Select(row => new SubCanal
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                IdSubCanal = row.Field<int?>("ID_SubCanal").GetValueOrDefault(),
                                Descripcion = row.Field<string>("SubCanal")
                            }).ToList();
                    }

                    //INFORMACION SIGNO VITAL
                    if (dsCampana.Tables[5].Rows.Count > 0)
                    {
                        ListSignoVital = dsCampana.Tables[5].AsEnumerable()
                            .Select(row => new SignoVital
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Descripcion = row.Field<string>("SignoVital")
                            }).ToList();
                    }


                    foreach (EntidadesCampanasPPG.Modelo.Campana campana in ListCampana)
                    {
                        campana.ListaKroma = ListKroma.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListaPublicidad = ListPublicidad.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListaRegion = ListRegion.Where(n => n.IdCampania == campana.IdCampana).ToList();

                        campana.ListaSubCanal = ListSubCanal.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListaSignoVital = ListSignoVital.Where(n => n.IdCampana == campana.IdCampana).ToList();
                    }


                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ListCampana;
        }
        public List<LineaFamilia> GuardarProdutoCampanaCompleto(string ClaveCampana, List<MecanicaRegalo> ListMecanicaRegalo, List<MecanicaMultiplo> ListMecanicaMultiplo,
                                                    List<MecanicaDescuento> ListMecanicaDescuento, List<MecanicaVolumen> ListMecanicaVolumen,
                                                    List<MecanicaKit> ListMecanicaKit, List<MecanicaCombo> ListMecanicaCombo, List<Tienda> ListTienda,
                                                    List<Tienda> ListTiendaExclusion, List<Alcance> ListAlcance)//, DataTable dtTienda)
        {
            List<LineaFamilia> ListLineaFamilia = new List<LineaFamilia>();
            DataTable dtMostrarLinea = new DataTable();

            const string spNameCampana = "MostrarIDCampania";

            const string spNameRegalo = "InsertRegalo";
            const string spNameEliminarRegalo = "DeleteEscenarioRegalo";

            const string spNameMultiplo = "InsertMultiplo";
            const string spNameEliminarMultiplo = "DeleteEscenarioMultiplo";

            const string spNameDescuento = "InsertDescuento";
            const string spNameEliminarDescuento = "DeleteEscenarioDescuento";

            const string spNameVolumen = "InsertVolumen";
            const string spNameEliminarVolumen = "DeleteEscenarioVolumen";

            const string spNameKit = "InsertKit";
            const string spNameEliminarKit = "DeleteEscenarioKit";

            const string spNameCombo = "InsertCombo";
            const string spNameEliminarCombo = "DeleteEscenarioCombo";

            const string spNameTienda = "InsertCampaniaTienda";
            const string spNameEliminarTienda = "DeleteCampaniaTienda";

            const string spNameAlcance = "InsertAlcanceProducto";
            const string spNameEliminarAlcance = "DeleteAlcanceProducto";

            const string spNameMostrarLinea = "MostrarLineas";

            List<SqlParameter> ListParametrosCampana = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosRegaloGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterRegaloGroup;
            List<SqlParameter> ListParametrosEliminarRegalo = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosMultiploGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterMultiploGroup;
            List<SqlParameter> ListParametrosEliminarMultiplo = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosDescuentoGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterDescuentoGroup;
            List<SqlParameter> ListParametrosEliminarDescuento = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosVolumenGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterVolumenGroup;
            List<SqlParameter> ListParametrosEliminarVolumen = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosKitGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterKitGroup;
            List<SqlParameter> ListParametrosEliminarKit = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosComboGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterComboGroup;
            List<SqlParameter> ListParametrosEliminarCombo = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosTiendaGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterTiendaGroup;
            List<SqlParameter> ListParametrosEliminarTienda = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosTiendaExclusionGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterTiendaExclusionGroup;
            List<SqlParameter> ListParametrosEliminarTiendaExclusion = new List<SqlParameter>();

            List<SqlParameterGroup> ListParametrosAlcanceGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterAlcanceGroup;
            List<SqlParameter> ListParametrosEliminarAlcance = new List<SqlParameter>();

            List<SqlParameter> ListParametrosMostrarLinea = new List<SqlParameter>();

            List<SqlParameter> ListParametros;
            SqlParameter Parametro;

            try
            {
                //DATOS CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Clv_Campania";
                Parametro.Value = ClaveCampana;
                ListParametrosCampana.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campc";
                Parametro.SqlDbType = SqlDbType.Int;
                Parametro.Direction = ParameterDirection.Output;
                ListParametrosCampana.Add(Parametro);

                //MECANICA REGALO
                foreach (MecanicaRegalo mecanicaRegalo in ListMecanicaRegalo)
                {
                    //DATOS MECANICA DESCUENTO
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Familia";
                    Parametro.Value = mecanicaRegalo.Familia;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SKU";
                    Parametro.Value = mecanicaRegalo.SKU;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion";
                    Parametro.Value = mecanicaRegalo.Descripcion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Tipo";
                    Parametro.Value = mecanicaRegalo.Tipo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Grupo";
                    Parametro.Value = mecanicaRegalo.Grupo_Regalo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@NumeroHijo";
                    Parametro.Value = mecanicaRegalo.NumeroHijo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = mecanicaRegalo.Alcance;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Capacidad";
                    Parametro.Value = mecanicaRegalo.Capacidad;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Dinamica";
                    Parametro.Value = mecanicaRegalo.Dinamica;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Ventas_Litros_Anio_Anterior";
                    Parametro.Value = mecanicaRegalo.VLitrosAnioAnt;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Presupuesto_Litros_Sin_Campania";
                    Parametro.Value = mecanicaRegalo.PLitrosSinCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Presupuesto_Litros_Con_Campania";
                    Parametro.Value = mecanicaRegalo.PLitrosConCamp;
                    ListParametros.Add(Parametro);

                    //Parametro = new SqlParameter();
                    //Parametro.ParameterName = "@SKU_Hijo";
                    //Parametro.Value = mecanicaRegalo.SKUHijo;
                    //ListParametros.Add(Parametro);

                    //Parametro = new SqlParameter();
                    //Parametro.ParameterName = "@Descripcion_Hijo";
                    //Parametro.Value = mecanicaRegalo.DescripcionHijo;
                    //ListParametros.Add(Parametro);

                    //Parametro = new SqlParameter();
                    //Parametro.ParameterName = "@Capacidad_Hijo";
                    //Parametro.Value = mecanicaRegalo.CapacidadHijo;
                    //ListParametros.Add(Parametro);

                    ParameterRegaloGroup = new SqlParameterGroup();
                    ParameterRegaloGroup.ListSqlParameter = ListParametros;

                    ListParametrosRegaloGrupo.Add(ParameterRegaloGroup);
                }

                //MECANICA MULTIPLO
                foreach (MecanicaMultiplo mecanicaMultiplo in ListMecanicaMultiplo)
                {
                    //DATOS MECANICA DESCUENTO
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Familia";
                    Parametro.Value = mecanicaMultiplo.Familia;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SKU";
                    Parametro.Value = mecanicaMultiplo.SKU;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion";
                    Parametro.Value = mecanicaMultiplo.Descripcion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = mecanicaMultiplo.Alcance;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Capacidad";
                    Parametro.Value = mecanicaMultiplo.Capacidad;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Dinamica";
                    Parametro.Value = mecanicaMultiplo.Dinamica;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@MultiploPadre";
                    Parametro.Value = mecanicaMultiplo.Multiplo_Padre;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@MultiploHijo";
                    Parametro.Value = mecanicaMultiplo.Multiplo_Hijo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Punto_Venta";
                    Parametro.Value = mecanicaMultiplo.Punto_Venta;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Venta_Anio_Anterior";
                    Parametro.Value = mecanicaMultiplo.VLitrosAnioAnt;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Presupuesto_Sin_campaña";
                    Parametro.Value = mecanicaMultiplo.PLitrosSinCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Presupuesto_Con_campaña";
                    Parametro.Value = mecanicaMultiplo.PLitrosConCamp;
                    ListParametros.Add(Parametro);

                    ParameterMultiploGroup = new SqlParameterGroup();
                    ParameterMultiploGroup.ListSqlParameter = ListParametros;

                    ListParametrosMultiploGrupo.Add(ParameterMultiploGroup);
                }

                //MECANICA DESCUENTO
                foreach (MecanicaDescuento mecanicaDescuento in ListMecanicaDescuento)
                {
                    //DATOS MECANICA DESCUENTO
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Familia";
                    Parametro.Value = mecanicaDescuento.Familia;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SKU";
                    Parametro.Value = mecanicaDescuento.SKU;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion";
                    Parametro.Value = mecanicaDescuento.Descripcion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = mecanicaDescuento.Alcance;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Capacidad";
                    Parametro.Value = mecanicaDescuento.Capacidad;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Dinamica";
                    Parametro.Value = mecanicaDescuento.Dinamica;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Porcentaje";
                    Parametro.Value = mecanicaDescuento.Porcentaje;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Importe";
                    Parametro.Value = mecanicaDescuento.Importe;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VentaAnioAnterior";
                    Parametro.Value = mecanicaDescuento.VLitrosAnioAnt;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoSinCampania";
                    Parametro.Value = mecanicaDescuento.PLitrosSinCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoConCampania";
                    Parametro.Value = mecanicaDescuento.PLitrosConCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Bloque";
                    Parametro.Value = mecanicaDescuento.Grupo_Descuento;
                    ListParametros.Add(Parametro);

                    ParameterDescuentoGroup = new SqlParameterGroup();
                    ParameterDescuentoGroup.ListSqlParameter = ListParametros;

                    ListParametrosDescuentoGrupo.Add(ParameterDescuentoGroup);
                }

                //MECANICA VOLUMEN
                foreach (MecanicaVolumen mecanicaVolumen in ListMecanicaVolumen)
                {
                    //DATOS MECANICA VOLUMEN
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Familia";
                    Parametro.Value = mecanicaVolumen.Familia;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SKU";
                    Parametro.Value = mecanicaVolumen.SKU;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion";
                    Parametro.Value = mecanicaVolumen.Descripcion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Bloque";
                    Parametro.Value = mecanicaVolumen.Bloque;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = mecanicaVolumen.Alcance;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Capacidad";
                    Parametro.Value = mecanicaVolumen.Capacidad;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Dinamica";
                    Parametro.Value = mecanicaVolumen.Dinamica;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VentaDesde";
                    Parametro.Value = mecanicaVolumen.De;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VentaHasta";
                    Parametro.Value = mecanicaVolumen.Hasta;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PorcentajeDescuento";
                    Parametro.Value = mecanicaVolumen.Descuento;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VentaAnioAnterior";
                    Parametro.Value = mecanicaVolumen.VLitrosAnioAnt;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoSinCampania";
                    Parametro.Value = mecanicaVolumen.PLitrosSinCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoConCampania";
                    Parametro.Value = mecanicaVolumen.PLitrosConCamp;
                    ListParametros.Add(Parametro);

                    ParameterVolumenGroup = new SqlParameterGroup();
                    ParameterVolumenGroup.ListSqlParameter = ListParametros;

                    ListParametrosVolumenGrupo.Add(ParameterVolumenGroup);
                }

                //MECANICA KIT
                foreach (MecanicaKit mecanicaKit in ListMecanicaKit)
                {
                    //DATOS MECANICA DESCUENTO
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Familia";
                    Parametro.Value = mecanicaKit.Familia;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SKU";
                    Parametro.Value = mecanicaKit.SKU;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Grupo";
                    Parametro.Value = mecanicaKit.Grupo_Kit;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Tipo";
                    Parametro.Value = mecanicaKit.Tipo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Numero_Hijo";
                    Parametro.Value = mecanicaKit.NumeroHijo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion";
                    Parametro.Value = mecanicaKit.Descripcion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = mecanicaKit.Alcance;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Capacidad";
                    Parametro.Value = mecanicaKit.Capacidad;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Dinamica";
                    Parametro.Value = mecanicaKit.Dinamica;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Porcentaje";
                    Parametro.Value = mecanicaKit.Porcentaje;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Importe";
                    Parametro.Value = mecanicaKit.Importe;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VentaAnioAnterior";
                    Parametro.Value = mecanicaKit.VLitrosAnioAnt;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoSinCampania";
                    Parametro.Value = mecanicaKit.PLitrosSinCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoConCampania";
                    Parametro.Value = mecanicaKit.PLitrosConCamp;
                    ListParametros.Add(Parametro);

                    ParameterKitGroup = new SqlParameterGroup();
                    ParameterKitGroup.ListSqlParameter = ListParametros;

                    ListParametrosKitGrupo.Add(ParameterKitGroup);
                }

                //MECANICA COMBO
                foreach (MecanicaCombo mecanicaCombo in ListMecanicaCombo)
                {
                    //DATOS MECANICA DESCUENTO
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Familia";
                    Parametro.Value = mecanicaCombo.Familia;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SKU";
                    Parametro.Value = mecanicaCombo.SKU;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion";
                    Parametro.Value = mecanicaCombo.Descripcion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TipoArticulo";
                    Parametro.Value = mecanicaCombo.Tipo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Grupo_Combo";
                    Parametro.Value = mecanicaCombo.Grupo_Combo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Numero_Padre";
                    Parametro.Value = mecanicaCombo.NumeroPadre;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Numero_Hijo";
                    Parametro.Value = mecanicaCombo.NumeroHijo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = mecanicaCombo.Alcance;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Capacidad";
                    Parametro.Value = mecanicaCombo.Capacidad;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Dinamica";
                    Parametro.Value = mecanicaCombo.Dinamica;
                    ListParametros.Add(Parametro);


                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VentaAnioAnterior";
                    Parametro.Value = mecanicaCombo.VLitrosAnioAnt;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoSinCampania";
                    Parametro.Value = mecanicaCombo.PLitrosSinCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoConCampania";
                    Parametro.Value = mecanicaCombo.PLitrosConCamp;
                    ListParametros.Add(Parametro);

                    ParameterComboGroup = new SqlParameterGroup();
                    ParameterComboGroup.ListSqlParameter = ListParametros;

                    ListParametrosComboGrupo.Add(ParameterComboGroup);
                }

                //CLIENTES
                foreach(Tienda tienda in ListTienda)
                {
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Tienda";
                    Parametro.Value = tienda.IdTienda;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Bill_To";
                    Parametro.Value = tienda.BillTo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Customer_Name";
                    Parametro.Value = tienda.CustomerName;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Region";
                    Parametro.Value = tienda.Region;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion_Region";
                    Parametro.Value = tienda.DescripcionRegion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion_Zona";
                    Parametro.Value = tienda.DescripcionZona;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Segmento";
                    Parametro.Value = tienda.Segmento;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Clave_Sobreprecio";
                    Parametro.Value = tienda.ClaveSobreprecio;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SubCanal";
                    Parametro.Value = tienda.SubCanal;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Exclusion";
                    Parametro.Value = tienda.Exclusion;
                    ListParametros.Add(Parametro);

                    ParameterTiendaGroup = new SqlParameterGroup();
                    ParameterTiendaGroup.ListSqlParameter = ListParametros;

                    ListParametrosTiendaGrupo.Add(ParameterTiendaGroup);
                }

                //EXCLUSIONES 
                foreach (Tienda tienda in ListTiendaExclusion)
                {
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Tienda";
                    Parametro.Value = tienda.IdTienda;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Bill_To";
                    Parametro.Value = tienda.BillTo;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Customer_Name";
                    Parametro.Value = tienda.CustomerName;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Region";
                    Parametro.Value = tienda.Region;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion_Region";
                    Parametro.Value = tienda.DescripcionRegion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion_Zona";
                    Parametro.Value = tienda.DescripcionZona;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Segmento";
                    Parametro.Value = tienda.Segmento;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Clave_Sobreprecio";
                    Parametro.Value = tienda.ClaveSobreprecio;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SubCanal";
                    Parametro.Value = tienda.SubCanal;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Exclusion";
                    Parametro.Value = tienda.Exclusion;
                    ListParametros.Add(Parametro);

                    ParameterTiendaExclusionGroup = new SqlParameterGroup();
                    ParameterTiendaExclusionGroup.ListSqlParameter = ListParametros;

                    ListParametrosTiendaExclusionGrupo.Add(ParameterTiendaExclusionGroup);
                }

                //ALCANCE
                foreach (Alcance alcance in ListAlcance)
                {
                    //DATOS MECANICA DESCUENTO
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@UnidadNegocio";
                    Parametro.Value = alcance.UnidadNegocio;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FamiliaEstelar";
                    Parametro.Value = alcance.ProductoEstelar;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Mecanica";
                    Parametro.Value = alcance.MecanicaPromocional;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = alcance.Observaciones;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Factor";
                    Parametro.Value = alcance.Factor;
                    ListParametros.Add(Parametro);

                    ParameterAlcanceGroup = new SqlParameterGroup();
                    ParameterAlcanceGroup.ListSqlParameter = ListParametros;

                    ListParametrosAlcanceGrupo.Add(ParameterAlcanceGroup);
                }

                dtMostrarLinea = base.ExecuteNonQueryProductoCampana(spNameCampana, ListParametrosCampana,
                                                            spNameRegalo, ListParametrosRegaloGrupo,
                                                            spNameEliminarRegalo, ListParametrosEliminarRegalo,
                                                            spNameMultiplo, ListParametrosMultiploGrupo,
                                                            spNameEliminarMultiplo, ListParametrosEliminarMultiplo,
                                                            spNameDescuento, ListParametrosDescuentoGrupo,
                                                            spNameEliminarDescuento, ListParametrosEliminarDescuento,
                                                            spNameVolumen, ListParametrosVolumenGrupo,
                                                            spNameEliminarVolumen, ListParametrosEliminarVolumen,
                                                            spNameKit, ListParametrosKitGrupo,
                                                            spNameEliminarKit, ListParametrosEliminarKit,
                                                            spNameCombo, ListParametrosComboGrupo,
                                                            spNameEliminarCombo, ListParametrosEliminarCombo,
                                                            spNameTienda, ListParametrosTiendaGrupo,
                                                            spNameEliminarTienda, ListParametrosEliminarTienda,
                                                            spNameTienda, ListParametrosTiendaExclusionGrupo,
                                                            spNameEliminarTienda, ListParametrosEliminarTiendaExclusion,
                                                            spNameAlcance, ListParametrosAlcanceGrupo,
                                                            spNameEliminarAlcance, ListParametrosEliminarAlcance,
                                                            spNameMostrarLinea, ListParametrosMostrarLinea);//, dtTienda);

                if(dtMostrarLinea != null)
                {
                    ListLineaFamilia = dtMostrarLinea.AsEnumerable()
                                        .Select(row => new LineaFamilia
                                        {
                                            Linea = string.IsNullOrEmpty(row.Field<string>("Linea")) ? "" : row.Field<string>("Linea"),
                                            Producto = string.IsNullOrEmpty(row.Field<string>("Familia")) ? "" : row.Field<string>("Familia"),
                                            //Region = row.Field<string>("Region"),
                                            DF = string.IsNullOrEmpty(row.Field<string>("Clave_Sobreprecio")) ? "" : row.Field<string>("Clave_Sobreprecio"),
                                            Validacion = string.IsNullOrEmpty(row.Field<string>("Validacion")) ? "" : row.Field<string>("Validacion")
                                        }).ToList();

                    if(ListLineaFamilia.Count <= 0)
                    {
                        ListLineaFamilia.Add(new LineaFamilia());
                    }

                    //LineaFamilia lineaFamilia = new LineaFamilia();

                    //lineaFamilia.Linea = "B1";
                    //lineaFamilia.DF = "DF";
                    //lineaFamilia.Producto = "Comex 100";
                    ////lineaFamilia.Region = "Norte";

                    //ListLineaFamilia.Add(lineaFamilia);

                    //lineaFamilia = new LineaFamilia();

                    //lineaFamilia.Linea = "B1";
                    //lineaFamilia.DF = "DF7";
                    //lineaFamilia.Producto = "Comex 100";
                    ////lineaFamilia.Region = "Norte";

                    //ListLineaFamilia.Add(lineaFamilia);

                    //lineaFamilia = new LineaFamilia();

                    //lineaFamilia.Linea = "B1";
                    //lineaFamilia.DF = "DF11";
                    //lineaFamilia.Producto = "Comex 100";
                    ////lineaFamilia.Region = "Norte";

                    //ListLineaFamilia.Add(lineaFamilia);
                }
                else
                {
                    throw new Exception("Ocurrio un error al querer guardar la informacion, intente de nuevo o consulte al administrador de sistemas.");
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ListLineaFamilia;
        }
        public int GuardarProductoCompleto(List<MecanicaDescuento> ListMecanicaDescuento)
        {
            const string spName = "InsertDescuento";
            int resultado = 0;

            List<SqlParameterGroup> ListParametrosGrupo = new List<SqlParameterGroup>();
            SqlParameterGroup ParameterGroup;
            List<SqlParameter> ListParametros;
            SqlParameter Parametro;

            try
            {
                foreach(MecanicaDescuento mecanicaDescuento in ListMecanicaDescuento)
                {
                    ListParametros = new List<SqlParameter>();

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Clv_Campania";
                    Parametro.Value = mecanicaDescuento.ClaveCampana;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@familia";
                    Parametro.Value = mecanicaDescuento.Familia;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@SKU";
                    Parametro.Value = mecanicaDescuento.SKU;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Descripcion";
                    Parametro.Value = mecanicaDescuento.Descripcion;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Alcance";
                    Parametro.Value = mecanicaDescuento.Alcance;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Capacidad";
                    Parametro.Value = mecanicaDescuento.Capacidad;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Dinamica";
                    Parametro.Value = mecanicaDescuento.Dinamica;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Porcentaje";
                    Parametro.Value = mecanicaDescuento.Porcentaje;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Importe";
                    Parametro.Value = mecanicaDescuento.Importe;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VentaAñoAnterior";
                    Parametro.Value = mecanicaDescuento.VLitrosAnioAnt;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoSinCampaña";
                    Parametro.Value = mecanicaDescuento.PLitrosSinCamp;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@PresupuestoConCampaña";
                    Parametro.Value = mecanicaDescuento.PLitrosConCamp;
                    ListParametros.Add(Parametro);

                    ParameterGroup = new SqlParameterGroup();
                    ParameterGroup.ListSqlParameter = ListParametros;

                    ListParametrosGrupo.Add(ParameterGroup);
                }

                resultado = base.ExecuteNonQueryTrans(spName, ListParametrosGrupo);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public Rentabilidad MostrarRentabilidad(Rentabilidad Rentabilidad)
        {
            const string spNameRentabilidad = "CalculoRentabilidad";

            const string tepexpan = "TEPEXPAN";
            const string aga = "AGA";
            const string fpu = "FPU";
            const string comprado = "COMPRADO";
            const string kroma = "KROMA";

            const string semaforoVerde = "VERDE";
            const string semaforoAmarillo = "AMARILLO";
            const string semafofoRojo = "ROJO";

            DataSet dsRentabilidad = new DataSet();
            DataTable dtRentabilidad = new DataTable();
            DataTable dtReporteCEO = new DataTable();
            DataTable dtReporteCEOMecanica = new DataTable();
            DataTable dtReporteCEOPublicidad = new DataTable();
            DataTable dtReporteMKT = new DataTable();
            DataTable dtReporteSKU = new DataTable();
            DataTable dtCampania = new DataTable();
            DataTable dtSKU = new DataTable();
            DataTable dtSKUPrecioCosto = new DataTable();

            List<SqlParameter> ListParametrosRentabilidad = new List<SqlParameter>();

            SqlParameter Parametro;

            List<Rentabilidad> ListRentabilidad = new List<Rentabilidad>();
            List<ReporteCEO> ListReporteCEO = new List<ReporteCEO>();
            List<ReporteCEOMecanica> ListReporteCEOMecanica = new List<ReporteCEOMecanica>();
            List<ReporteCEOPublicidad> ListReporteCEOPublicidad = new List<ReporteCEOPublicidad>();
            List<ReporteMKT> ListReporteMKT = new List<ReporteMKT>();
            List<ReporteSKU> ListReporteSKU = new List<ReporteSKU>();
            List<EntidadesCampanasPPG.Modelo.Campana> ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();
            List<SKUPrecioCosto> ListSKUPrecioCosto = new List<SKUPrecioCosto>();

            GastoPlanta gastoPlanta;

            SemaforoRentabilidad semaforoRentabilidad;

            try
            {
                //IFormatProvider culture = new CultureInfo("es-MX", true);

                //DATOS RENTABILIDAD
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Clv_Campania";
                Parametro.Value = Rentabilidad.ClaveCampania;
                ListParametrosRentabilidad.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PeriodoAnt";
                Parametro.Value = Rentabilidad.PeriodoAnt;
                ListParametrosRentabilidad.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PeriodoAct";
                Parametro.Value = Rentabilidad.PeriodoAct;
                ListParametrosRentabilidad.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Validacion";
                Parametro.Value = Rentabilidad.EsGuardar;
                ListParametrosRentabilidad.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPGID";
                Parametro.Value = Rentabilidad.PPGID;
                ListParametrosRentabilidad.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Usuario";
                Parametro.Value = Rentabilidad.Usuario;
                ListParametrosRentabilidad.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Estatus";
                Parametro.Value = Rentabilidad.Estatus;
                ListParametrosRentabilidad.Add(Parametro);

                //DATOS TEPEXPAN
                gastoPlanta = Rentabilidad.ListGastoPlanta.Where(n => n.Planta.ToUpper().Contains(tepexpan)).FirstOrDefault();

                if(gastoPlanta != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TepexpanAnterior";
                    Parametro.Value = gastoPlanta.AnioAnterior;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@TepexpanActual";
                    Parametro.Value = gastoPlanta.AnioActual;
                    ListParametrosRentabilidad.Add(Parametro);
                }

                //DATOS AGA
                gastoPlanta = Rentabilidad.ListGastoPlanta.Where(n => n.Planta.ToUpper().Contains(aga)).FirstOrDefault();

                if (gastoPlanta != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@AgaAnterior";
                    Parametro.Value = gastoPlanta.AnioAnterior;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@AgaActual";
                    Parametro.Value = gastoPlanta.AnioActual;
                    ListParametrosRentabilidad.Add(Parametro);
                }

                //DATOS FPU
                gastoPlanta = Rentabilidad.ListGastoPlanta.Where(n => n.Planta.ToUpper().Contains(fpu)).FirstOrDefault();

                if (gastoPlanta != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FPUAnterior";
                    Parametro.Value = gastoPlanta.AnioAnterior;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FPUActual";
                    Parametro.Value = gastoPlanta.AnioActual;
                    ListParametrosRentabilidad.Add(Parametro);
                }

                //DATOS COMPRADO
                gastoPlanta = Rentabilidad.ListGastoPlanta.Where(n => n.Planta.ToUpper().Contains(comprado)).FirstOrDefault();

                if (gastoPlanta != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@CompradoAnterior";
                    Parametro.Value = gastoPlanta.AnioAnterior;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@CompradoActual";
                    Parametro.Value = gastoPlanta.AnioActual;
                    ListParametrosRentabilidad.Add(Parametro);
                }

                //DATOS KROMA
                gastoPlanta = Rentabilidad.ListGastoPlanta.Where(n => n.Planta.ToUpper().Contains(kroma)).FirstOrDefault();

                if (gastoPlanta != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@KromaAnterior";
                    Parametro.Value = gastoPlanta.AnioAnterior;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@KromaActual";
                    Parametro.Value = gastoPlanta.AnioActual;
                    ListParametrosRentabilidad.Add(Parametro);
                }


                //DATOS SEMAFORO VERDE
                semaforoRentabilidad = Rentabilidad.ListSemaforoRentabilidad.Where(n => n.Descripcion.ToUpper().Contains(semaforoVerde)).FirstOrDefault();

                if (semaforoRentabilidad != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VerdeDE";
                    Parametro.Value = semaforoRentabilidad.De;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@VerdeHasta";
                    Parametro.Value = semaforoRentabilidad.Hasta;
                    ListParametrosRentabilidad.Add(Parametro);
                }

                //DATOS SEMAFORO AMARILLO
                semaforoRentabilidad = Rentabilidad.ListSemaforoRentabilidad.Where(n => n.Descripcion.ToUpper().Contains(semaforoAmarillo)).FirstOrDefault();

                if (semaforoRentabilidad != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@AmarilloDE";
                    Parametro.Value = semaforoRentabilidad.De;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@AmarilloHasta";
                    Parametro.Value = semaforoRentabilidad.Hasta;
                    ListParametrosRentabilidad.Add(Parametro);
                }

                //DATOS SEMAFORO ROJO
                semaforoRentabilidad = Rentabilidad.ListSemaforoRentabilidad.Where(n => n.Descripcion.ToUpper().Contains(semafofoRojo)).FirstOrDefault();

                if (semaforoRentabilidad != null)
                {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@RojoDE";
                    Parametro.Value = semaforoRentabilidad.De;
                    ListParametrosRentabilidad.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@RojoHasta";
                    Parametro.Value = semaforoRentabilidad.Hasta;
                    ListParametrosRentabilidad.Add(Parametro);
                }


                dsRentabilidad = base.ExecuteDataSet(spNameRentabilidad, ListParametrosRentabilidad);

                if(dsRentabilidad.Tables.Count > 0)
                {
                    //INFORMACION DE LAS TABLAS
                    dtSKU = dsRentabilidad.Tables[0];

                    if (dtSKU.Rows.Count > 0)
                    {
                        Rentabilidad.SKU = dtSKU.Rows[0][0].ToString();
                    }


                    if (dsRentabilidad.Tables.Count > 1)
                    {
                        dtRentabilidad = dsRentabilidad.Tables[1];

                        //DATOS RENTABILIDAD
                        ListRentabilidad = dtRentabilidad.AsEnumerable()
                                                .Select(row => new Rentabilidad
                                                {
                                                    PorcentajeCompania = row.Field<decimal?>("PorcenCampania").GetValueOrDefault(),
                                                    PorcentajeNecesario = row.Field<decimal?>("PorcenNecesario").GetValueOrDefault(),
                                                    Comentario = string.IsNullOrEmpty(row.Field<string>("Comentario")) ? "": row.Field<string>("Comentario")
                                                }).ToList();

                        if (ListRentabilidad.Count > 0)
                        {
                            Rentabilidad.PorcentajeCompania = ListRentabilidad.FirstOrDefault().PorcentajeCompania;
                            Rentabilidad.PorcentajeNecesario = ListRentabilidad.FirstOrDefault().PorcentajeNecesario;
                            Rentabilidad.Comentario = ListRentabilidad.FirstOrDefault().Comentario;
                        }
                    }

                    if (dsRentabilidad.Tables.Count > 2)
                    {
                        dtCampania = dsRentabilidad.Tables[2];

                        //DATOS CAMPAÑA
                        ListCampana = dtCampania.AsEnumerable()
                                        .Select(row => new EntidadesCampanasPPG.Modelo.Campana
                                        {
                                            Title = string.IsNullOrEmpty(row.Field<string>("Camp_Number")) ? "" : row.Field<string>("Camp_Number"),
                                            NombreCampa = string.IsNullOrEmpty(row.Field<string>("Nombre_Camp")) ? "" : row.Field<string>("Nombre_Camp"),
                                            FechaInicioSubCanal = row.Field<DateTime?>("Fecha_Inicio_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                            FechaFinSubCanal = row.Field<DateTime?>("Fecha_Fin_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                            FechaInicioPublico = row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                            FechaFinPublico = row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy")
                                        }).ToList();

                        Rentabilidad.ListCampania = ListCampana;
                    }

                    if (dsRentabilidad.Tables.Count > 3)
                    {
                        dtReporteCEO = dsRentabilidad.Tables[3];

                        //DATOS REPORTE CEO
                        ListReporteCEO = dtReporteCEO.AsEnumerable()
                                            .Select(row => new ReporteCEO
                                            {
                                                ClaveCampania = string.IsNullOrEmpty(row.Field<string>("NOMBRE CAMPANIA")) ? "" : row.Field<string>("NOMBRE CAMPANIA"),
                                                Periodo = string.IsNullOrEmpty(row.Field<string>("PERIODO")) ? "" : row.Field<string>("PERIODO"),
                                                SubtotalLitros = row.Field<decimal?>("SUBTOTAL LITROS").GetValueOrDefault(),
                                                SubtotalPiezas = row.Field<decimal?>("SUBTOTAL PIEZAS").GetValueOrDefault(),
                                                TotalLitrosPiezas = row.Field<decimal?>("TOTAL LITROS / PIEZAS").GetValueOrDefault(),
                                                Importe = row.Field<decimal?>("IMPORTE").GetValueOrDefault(),
                                                ImporteLitros = row.Field<decimal?>("IMPORTE LITROS").GetValueOrDefault(),
                                                ImportePiezas = row.Field<decimal?>("IMPORTE PZAS").GetValueOrDefault(),
                                                PrecioPromedioLitro = row.Field<decimal?>("PRECIO PROMEDIO X LT").GetValueOrDefault(),
                                                CostoMP = row.Field<decimal?>("COSTO MATERIA PRIMA + ENVASE + FABRICACION").GetValueOrDefault(),
                                                UtilidadMP = row.Field<decimal?>("UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION").GetValueOrDefault(),
                                                FactorUtilidad = row.Field<decimal?>("FACTOR UTILIDAD").GetValueOrDefault(),
                                                //MP = row.Field<decimal?>("MATERIA PRIMA + ENVASE + FABRICACION").GetValueOrDefault(),
                                                MargenUtilidad = row.Field<decimal?>("% MARGEN DE UTILIDAD CONSIDERA MP + ENV + FABR").GetValueOrDefault(),
                                                InversionPublicidad = row.Field<decimal?>("INVERSION PUBLICITARIA").GetValueOrDefault(),
                                                OtrosGastos = row.Field<decimal?>("OTROS GASTOS DE OPERACIÓN PLANTA").GetValueOrDefault(),
                                                GastosOperacion = row.Field<decimal?>("GASTOS DE OPERACIÓN GENERALES KROMA").GetValueOrDefault(),
                                                NotasCredito = row.Field<decimal?>("NOTAS DE CRÉDITO").GetValueOrDefault(),
                                                UtilidadConsideraMP = row.Field<decimal?>("UTILIDAD CONSIDERA MP + ENV + FABR + KROMA + PLANTA + NOTAS").GetValueOrDefault(),
                                                UtilidadLitroPieza = row.Field<decimal?>("UTILIDAD POR LT/PZ").GetValueOrDefault(),
                                                FactorUtilidadMP = row.Field<decimal?>("FACTOR DE UTILIDAD CONSIDERA MP + ENV + FABR + KROMA + PLANTA").GetValueOrDefault(),
                                                PorcenUtilidadConsideraMP = row.Field<decimal?>("% MARGEN DE UTILIDAD CONSIDERA MP + ENVASE + FABR + KROMA + PLANTA").GetValueOrDefault(),
                                                PorcenIncrementoLitros = row.Field<decimal?>("% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR").GetValueOrDefault(),
                                                PorcenIncrementoLitrosPresu = row.Field<decimal?>("% INCREMENTO EN LITROS Vs PRESUPUESTO").GetValueOrDefault(),
                                                Roi = row.Field<decimal?>("ROI").GetValueOrDefault(),
                                                ImportePrecioPublico = row.Field<decimal?>("IMPORTE PRECIO PUBLICO").GetValueOrDefault(),
                                                ImportePrecioPublicoSinIva = row.Field<decimal?>("IMPORTE PRECIO PUBLICO SIN IVA").GetValueOrDefault(),
                                                UtilidadEnConcesionario = row.Field<decimal?>("UTILIDAD EN CONCESIONARIO").GetValueOrDefault(),
                                                MargenConcesionario = row.Field<decimal?>("MARGEN CONCESIONARIO").GetValueOrDefault()
                                            }).ToList();

                        Rentabilidad.ListReporteCEO = ListReporteCEO;
                    }

                    if (dsRentabilidad.Tables.Count > 4)
                    {
                        dtReporteCEOMecanica = dsRentabilidad.Tables[4];

                        //DATOS REPORTE CEO MECANICA
                        ListReporteCEOMecanica = dtReporteCEOMecanica.AsEnumerable()
                                            .Select(row => new ReporteCEOMecanica
                                            {
                                                //ClaveCampania = row.Field<string>("NOMBRE CAMPANIA"),
                                                IdCampania = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                                Cantidad = string.IsNullOrEmpty(row.Field<string>("Cantidad")) ? "" : row.Field<string>("Cantidad"),
                                                Mecanica = string.IsNullOrEmpty(row.Field<string>("Mecanica")) ? "": row.Field<string>("Mecanica"),
                                                Articulo = string.IsNullOrEmpty(row.Field<string>("FamiliaEstelar")) ? "": row.Field<string>("FamiliaEstelar"),
                                                Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "": row.Field<string>("Alcance"),
                                                Periodo = row.Field<int?>("Periodo").GetValueOrDefault(),
                                            }).ToList();

                        Rentabilidad.ListReporteCEOMecanica = ListReporteCEOMecanica;
                    }

                    if (dsRentabilidad.Tables.Count > 5)
                    {
                        dtReporteCEOPublicidad = dsRentabilidad.Tables[5];

                        //DATOS REPORTE CEO PUBLICIDAD
                        ListReporteCEOPublicidad = dtReporteCEOPublicidad.AsEnumerable()
                                            .Select(row => new ReporteCEOPublicidad
                                            {
                                                //ClaveCampania = row.Field<string>("ClaveCampania"),
                                                IdCampania = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                                PeriodoAnterior = row.Field<decimal?>("PeriodoAnterior").GetValueOrDefault(),
                                                PeriodoActual = row.Field<decimal?>("PeriodoActual").GetValueOrDefault(),
                                                Publicidad = string.IsNullOrEmpty(row.Field<string>("Publicidad")) ? "" : row.Field<string>("Publicidad")
                                            }).ToList();

                        Rentabilidad.ListReporteCEOPublicidad = ListReporteCEOPublicidad;
                    }

                    if (dsRentabilidad.Tables.Count > 6)
                    {
                        dtReporteMKT = dsRentabilidad.Tables[6];

                        //DATOS REPORTE MERCADOTECNIA
                        ListReporteMKT = dtReporteMKT.AsEnumerable()
                                            .Select(row => new ReporteMKT
                                            {
                                                Articulo = string.IsNullOrEmpty(row.Field<string>("Articulo")) ? "" : row.Field<string>("Articulo"),
                                                Descripcion = string.IsNullOrEmpty(row.Field<string>("Descripcion")) ? "" : row.Field<string>("Descripcion"),
                                                LitrosActualCC = row.Field<decimal?>("LITROS ACTUAL CC").GetValueOrDefault(),
                                                PiezasActualCC = row.Field<decimal?>("PIEZAS ACTUAL CC").GetValueOrDefault(),
                                                ImporteCPRecioLitros = row.Field<decimal?>("IMPORTE C/PRECIO CONC. LITROS").GetValueOrDefault(),
                                                ImporteCPrecioPiezas = row.Field<decimal?>("Importe C/ Precio Piezas").GetValueOrDefault(),
                                                Rentabilidad = string.IsNullOrEmpty(row.Field<string>("Rentabilidad")) ? "" : row.Field<string>("Rentabilidad"),
                                                Litros = row.Field<decimal?>("Litros").GetValueOrDefault(),
                                                Piezas = row.Field<decimal?>("Piezas").GetValueOrDefault(),
                                                Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance")
                                            }).ToList();

                        Rentabilidad.ListReporteMKT = ListReporteMKT;
                    }

                    if (dsRentabilidad.Tables.Count > 7)
                    {
                        dtReporteSKU = dsRentabilidad.Tables[7];

                        //DATOS REPORTE SKU
                        ListReporteSKU = dtReporteSKU.AsEnumerable()
                                            .Select(row => new ReporteSKU
                                            {
                                                //ClaveCampania = row.Field<string>("Articulo"),
                                                Articulo = string.IsNullOrEmpty(row.Field<string>(0)) ? "" : row.Field<string>(0),
                                                Descripcion = string.IsNullOrEmpty(row.Field<string>(1)) ? "" : row.Field<string>(1),
                                                Capacidad = string.IsNullOrEmpty(row.Field<string>(2)) ? "" : row.Field<string>(2),
                                                CostoMPAnterior = row.Field<decimal?>(3).GetValueOrDefault(),
                                                GastoPlantaAnterior = row.Field<decimal?>(4).GetValueOrDefault(),
                                                GastoKromaAñoAnterior = row.Field<decimal?>(5).GetValueOrDefault(),
                                                CostoMPEnvConvAñoActual = row.Field<decimal?>(6).GetValueOrDefault(),
                                                CostoPlantaAñoActual = row.Field<decimal?>(7).GetValueOrDefault(),
                                                GastoKromaActual = row.Field<decimal?>(8).GetValueOrDefault(),
                                                PrecioConcAnterior = row.Field<decimal?>(9).GetValueOrDefault(),
                                                PrecioAntConPromocion = row.Field<decimal?>(10).GetValueOrDefault(),
                                                InversionTintaAnt = row.Field<decimal?>(11).GetValueOrDefault(),
                                                PrecioPublicoAntCP = row.Field<decimal?>(12).GetValueOrDefault(),
                                                PrecioConc = row.Field<decimal?>(13).GetValueOrDefault(),
                                                FactorUtilidadBruto = row.Field<decimal?>(14).GetValueOrDefault(),
                                                PrecioPublicoSC = row.Field<decimal?>(15).GetValueOrDefault(),
                                                InversionTintaActSC = row.Field<decimal?>(16).GetValueOrDefault(),
                                                PorcMargenConcSC = row.Field<decimal?>(17).GetValueOrDefault(),
                                                PrecioConce = row.Field<decimal?>(18).GetValueOrDefault(),
                                                //18
                                                FactorUtilidadCC = row.Field<decimal?>(19).GetValueOrDefault(),
                                                PrecioPublicoCC = row.Field<decimal?>(20).GetValueOrDefault(),
                                                InversionTintaActCC = row.Field<decimal?>(21).GetValueOrDefault(),
                                                PorcMargenConcCC = row.Field<decimal?>(22).GetValueOrDefault(),
                                                LitrosAnterior = row.Field<decimal?>(23).GetValueOrDefault(),
                                                PiezasVacioAnt = row.Field<decimal?>(24).GetValueOrDefault(),
                                                PiezasAnt = row.Field<decimal?>(25).GetValueOrDefault(),
                                                CostoMPEnvFab = row.Field<decimal?>(26).GetValueOrDefault(),
                                                GastoPlantaAnt = row.Field<decimal?>(27).GetValueOrDefault(),
                                                GastoKromaAnter = row.Field<decimal?>(28).GetValueOrDefault(),
                                                UtilidadBrutaComexAnt = row.Field<decimal?>(29).GetValueOrDefault(),
                                                ImporteCPrecioConcLitro = row.Field<decimal?>(30).GetValueOrDefault(),
                                                ImporteCPrecioConcPiezas = row.Field<decimal?>(31).GetValueOrDefault(),
                                                UtilidadConcAnterior = row.Field<decimal?>(32).GetValueOrDefault(),
                                                InversionTintaAnterior = row.Field<decimal?>(33).GetValueOrDefault(),
                                                ImporteCPrecioPublicoAnt = row.Field<decimal?>(34).GetValueOrDefault(),
                                                //34
                                                LitrosActualSC = row.Field<decimal?>(35).GetValueOrDefault(),
                                                LitrosVacioSC = row.Field<decimal?>(36).GetValueOrDefault(),
                                                PiezasActualSC = row.Field<decimal?>(37).GetValueOrDefault(),
                                                CostoMPEnvFabrica = row.Field<decimal?>(38).GetValueOrDefault(),
                                                GastoPlantaActualSC = row.Field<decimal?>(39).GetValueOrDefault(),
                                                GastoKromaActualSC = row.Field<decimal?>(40).GetValueOrDefault(),
                                                UtilidadBrutaComexSC = row.Field<decimal?>(41).GetValueOrDefault(),
                                                Importe = row.Field<decimal?>(42).GetValueOrDefault(),
                                                ImporteCPrecioConcPiezasSC = row.Field<decimal?>(43).GetValueOrDefault(),
                                                UtilidadConcSC = row.Field<decimal?>(44).GetValueOrDefault(),
                                                InversionTintaSC = row.Field<decimal?>(45).GetValueOrDefault(),
                                                ImporteCPrecioPublicoSC = row.Field<decimal?>(46).GetValueOrDefault(),
                                                LitrosActualCC = row.Field<decimal?>(47).GetValueOrDefault(),
                                                LitrosVacioCC = row.Field<decimal?>(48).GetValueOrDefault(),
                                                PiezasActualCC = row.Field<decimal?>(49).GetValueOrDefault(),
                                                //FALTA
                                                CostoMPEnvFabCC = row.Field<decimal?>(50).GetValueOrDefault(),
                                                GastoPlantaActualCC = row.Field<decimal?>(51).GetValueOrDefault(),
                                                GastoKromaActualCC = row.Field<decimal?>(52).GetValueOrDefault(),
                                                UtilidadBrutaComexCC = row.Field<decimal?>(53).GetValueOrDefault(),
                                                ImporteCPrecioConcLitros = row.Field<decimal?>(54).GetValueOrDefault(),
                                                ImporteCPrecioConcCCPiezas = row.Field<decimal?>(55).GetValueOrDefault(),
                                                UtilidadConcCC = row.Field<decimal?>(56).GetValueOrDefault(),
                                                InversionTintaCC = row.Field<decimal?>(57).GetValueOrDefault(),
                                                ImporteCPrecioPublicoCC = row.Field<decimal?>(58).GetValueOrDefault(),
                                                Anterior = row.Field<decimal?>(59).GetValueOrDefault(),
                                                SinCampania = row.Field<decimal?>(60).GetValueOrDefault(),
                                                ConPromocion = row.Field<decimal?>(61).GetValueOrDefault(),
                                                Rentabilidad = row.Field<decimal?>(62).GetValueOrDefault(),
                                                Comentario = string.IsNullOrEmpty(row.Field<string>(63)) ? "" : row.Field<string>(63),
                                                //NUEVAS COLUMNAS
                                                LitrosNecesarios = row.Field<decimal?>(64).GetValueOrDefault(),
                                                PiezasNecesario = row.Field<decimal?>(65).GetValueOrDefault(),
                                                Piezas = row.Field<decimal?>(66).GetValueOrDefault(),
                                                CostoMPEnvNecesario = row.Field<decimal?>(67).GetValueOrDefault(),
                                                CostoPlantaNecesario = row.Field<decimal?>(68).GetValueOrDefault(),
                                                GastoKromaNecesario = row.Field<decimal?>(69).GetValueOrDefault(),
                                                UtilidadBruta = row.Field<decimal?>(70).GetValueOrDefault(),
                                                ImporteCPrecioConcLitrosNecesario = row.Field<decimal?>(71).GetValueOrDefault(),
                                                ImporteCPrecioPublicoNecesario = row.Field<decimal?>(72).GetValueOrDefault(),
                                                UtilidadConc = row.Field<decimal?>(73).GetValueOrDefault(),
                                                TintaNecesario = row.Field<decimal?>(74).GetValueOrDefault(),
                                                ImporteCPrecioConcPiezasNecesario = row.Field<decimal?>(75).GetValueOrDefault(),
                                                Alcance = string.IsNullOrEmpty(row.Field<string>(76)) ? "" : row.Field<string>(76),
                                            }).ToList();

                        Rentabilidad.ListReporteSKU = ListReporteSKU;
                    }

                    if (dsRentabilidad.Tables.Count > 8)
                    {
                        dtSKUPrecioCosto = dsRentabilidad.Tables[8];

                        //DATOS SKU CON PRECIO CONCESIONARIO Y COSTO MP
                        ListSKUPrecioCosto = dtSKUPrecioCosto.AsEnumerable()
                                            .Select(row => new SKUPrecioCosto
                                            {
                                                Articulo = string.IsNullOrEmpty(row.Field<string>("Articulo")) ? "" : row.Field<string>("Articulo"),
                                                PrecioConcesionario = string.IsNullOrEmpty(row.Field<string>("PrecioConcesionario")) ? "" : row.Field<string>("PrecioConcesionario"),
                                                CostoMP = string.IsNullOrEmpty(row.Field<string>("CostoMP")) ? "" : row.Field<string>("CostoMP")
                                            }).ToList();

                        Rentabilidad.ListSKUPrecioCosto = ListSKUPrecioCosto;
                    }

                    Rentabilidad.Mensaje = "OK";
                }
                else
                {
                    Rentabilidad.Mensaje = "No se encontraron resultados con los parametros ingresados.";
                }

                return Rentabilidad;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<ReporteCEO> MostrarReporteCEO(string ClaveCampania)
        {
            const string spName = "ReporteCEO";
            List<ReporteCEO> ListReporteCEO = new List<ReporteCEO>();
            DataTable dtReporteCEO = new DataTable();

            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Clv_Campania";
                Parametro.Value = ClaveCampania;
                ListParametros.Add(Parametro);

                dtReporteCEO = base.ExecuteDataTable(spName, ListParametros);

                if (dtReporteCEO.Rows.Count > 0)
                {
                    ListReporteCEO = dtReporteCEO.AsEnumerable()
                                            .Select(row => new ReporteCEO
                                            {
                                                ClaveCampania = row.Field<string>("NOMBRE CAMPANIA"),
                                                Periodo = row.Field<string>("PERIODO"),
                                                SubtotalLitros = row.Field<decimal?>("SUBTOTAL LITROS").GetValueOrDefault(),
                                                SubtotalPiezas = row.Field<decimal?>("SUBTOTAL PIEZAS").GetValueOrDefault(),
                                                TotalLitrosPiezas = row.Field<decimal?>("TOTAL LITROS / PIEZAS").GetValueOrDefault(),
                                                Importe = row.Field<decimal?>("IMPORTE").GetValueOrDefault(),
                                                ImporteLitros = row.Field<decimal?>("IMPORTE LITROS").GetValueOrDefault(),
                                                ImportePiezas = row.Field<decimal?>("IMPORTE PZAS").GetValueOrDefault(),
                                                PrecioPromedioLitro = row.Field<decimal?>("PRECIO PROMEDIO X LT").GetValueOrDefault(),
                                                CostoMP = row.Field<decimal?>("COSTO MATERIA PRIMA + ENVASE + FABRICACION").GetValueOrDefault(),
                                                UtilidadMP = row.Field<decimal?>("UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION").GetValueOrDefault(),
                                                FactorUtilidad = row.Field<decimal?>("FACTOR UTILIDAD").GetValueOrDefault(),
                                                //MP = row.Field<decimal?>("MATERIA PRIMA + ENVASE + FABRICACION").GetValueOrDefault(),
                                                MargenUtilidad = row.Field<decimal?>("% MARGEN DE UTILIDAD CONSIDERA MP + ENV + FABR").GetValueOrDefault(),
                                                InversionPublicidad = row.Field<decimal?>("INVERSION PUBLICITARIA").GetValueOrDefault(),
                                                OtrosGastos = row.Field<decimal?>("OTROS GASTOS DE OPERACIÓN PLANTA").GetValueOrDefault(),
                                                GastosOperacion = row.Field<decimal?>("GASTOS DE OPERACIÓN GENERALES KROMA").GetValueOrDefault(),
                                                NotasCredito = row.Field<decimal?>("NOTAS DE CRÉDITO").GetValueOrDefault(),
                                                UtilidadConsideraMP = row.Field<decimal?>("UTILIDAD CONSIDERA MP + ENV + FABR + KROMA + PLANTA + NOTAS").GetValueOrDefault(),
                                                UtilidadLitroPieza = row.Field<decimal?>("UTILIDAD POR LT/PZ").GetValueOrDefault(),
                                                FactorUtilidadMP = row.Field<decimal?>("FACTOR DE UTILIDAD CONSIDERA MP + ENV + FABR + KROMA + PLANTA").GetValueOrDefault(),
                                                PorcenUtilidadConsideraMP = row.Field<decimal?>("% MARGEN DE UTILIDAD CONSIDERA MP + ENVASE + FABR + KROMA + PLANTA").GetValueOrDefault(),
                                                PorcenIncrementoLitros = row.Field<decimal?>("% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR").GetValueOrDefault(),
                                                PorcenIncrementoLitrosPresu = row.Field<decimal?>("% INCREMENTO EN LITROS Vs PRESUPUESTO").GetValueOrDefault(),
                                                Roi = row.Field<decimal?>("ROI").GetValueOrDefault()
                                            }).ToList();

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ListReporteCEO;
        }
        public List<FlujoActividad> MostrarAprobadorPredecesor(FlujoActividad flujoActividad)
        {
            const string spName = "MostrarAprobadorPredecesor";

            SqlParameter Parametro = new SqlParameter();
            List<SqlParameter> ListParametros = new List<SqlParameter>();

            List<FlujoActividad> ListFlujoActividad = new List<FlujoActividad>();
            DataTable dtFlujoActividad = new DataTable();

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@IdCampania";
                Parametro.Value = flujoActividad.IdCampania;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@IdTarea";
                Parametro.Value = flujoActividad.IDTarea;
                ListParametros.Add(Parametro);

                dtFlujoActividad = base.ExecuteDataTable(spName, ListParametros);

                ListFlujoActividad = dtFlujoActividad.AsEnumerable()
                                        .Select(n => new FlujoActividad
                                        {
                                            IdCampania = n.Field<int?>("ID_Campaña").GetValueOrDefault(),
                                            ClaveCampania = string.IsNullOrEmpty(n.Field<string>("Camp_Number")) ? "": n.Field<string>("Camp_Number"),
                                            NombreCampania = string.IsNullOrEmpty(n.Field<string>("Nombre_Camp")) ? "" : n.Field<string>("Nombre_Camp"),
                                            TipoFlujo = n.Field<int?>("IdTipoFlujo").GetValueOrDefault(),
                                            IDTarea = n.Field<int?>("ID_Tarea").GetValueOrDefault(),
                                            TxtTarea = string.IsNullOrEmpty(n.Field<string>("Actividad")) ? "" : n.Field<string>("Actividad"),
                                            FechaInicio = n.Field<DateTime?>("FechaInicio").GetValueOrDefault(),
                                            FechaFin = n.Field<DateTime?>("FechaFin").GetValueOrDefault(),
                                            IdDependiente = string.IsNullOrEmpty(n.Field<string>("IdDependiente")) ? "" : n.Field<string>("IdDependiente"),
                                            MailResponsable = string.IsNullOrEmpty(n.Field<string>("Correo")) ? "" : n.Field<string>("Correo"),
                                            MailResponsable2 = string.IsNullOrEmpty(n.Field<string>("Correo_2")) ? "" : n.Field<string>("Correo_2")
                                        }).ToList();
            }
            catch(Exception ex)
            {
                throw ex;
            }

            return ListFlujoActividad;
        }
        public GrupoReporteGR MostrarReporteGR(int IdCampania)
        {
            GrupoReporteGR grupoReporteGR = new GrupoReporteGR();
            grupoReporteGR.ListReporteGR1 = new List<ReporteGR1>();
            grupoReporteGR.ListReporteGR2 = new List<ReporteGR2>();
            grupoReporteGR.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

            string spName = "SalidaPlexus";
            DataSet dsReporteGR = new DataSet();

            SqlParameter Parametro = new SqlParameter();
            List<SqlParameter> ListParametros = new List<SqlParameter>();

            try
            {
                //DATOS RENTABILIDAD
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campania";
                Parametro.Value = IdCampania;
                ListParametros.Add(Parametro);

                dsReporteGR = base.ExecuteDataSet(spName, ListParametros);

                if(dsReporteGR.Tables.Count > 0)
                {
                    //DATOS CAMPAÑA
                    grupoReporteGR.ListCampana = dsReporteGR.Tables[0].AsEnumerable()
                                    .Select(row => new EntidadesCampanasPPG.Modelo.Campana
                                    {
                                        Title = string.IsNullOrEmpty(row.Field<string>("Camp_Number")) ? "" : row.Field<string>("Camp_Number"),
                                        NombreCampa = string.IsNullOrEmpty(row.Field<string>("Nombre_Camp")) ? "" : row.Field<string>("Nombre_Camp"),
                                        //FechaInicioSubCanal = row.Field<DateTime?>("Fecha_Inicio_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                        //FechaFinSubCanal = row.Field<DateTime?>("Fecha_Fin_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                        FechaInicioPublico = row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                        FechaFinPublico = row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy")
                                    }).ToList();
                }

                if (dsReporteGR.Tables.Count > 1)
                {
                    grupoReporteGR.ListReporteGR1 = dsReporteGR.Tables[1].AsEnumerable()
                                                    .Select(n => new ReporteGR1
                                                    {
                                                        BillTo = n.Field<string>("Bill_To"),
                                                        CustomerName = string.IsNullOrEmpty(n.Field<string>("Customer_Name")) ? "" : n.Field<string>("Customer_Name"),
                                                        DescripcionRegion = string.IsNullOrEmpty(n.Field<string>("Descripcion_Region")) ? "" : n.Field<string>("Descripcion_Region"),
                                                        IdTienda = n.Field<decimal?>("ID_Tienda").GetValueOrDefault(),
                                                        ZonaTerritorial = string.IsNullOrEmpty(n.Field<string>("Zona_Territorial")) ? "" : n.Field<string>("Zona_Territorial"),
                                                        Segmento = n.Field<string>("Segmento")
                                                    }).ToList();
                }

                if(dsReporteGR.Tables.Count > 2)
                {
                    grupoReporteGR.ListReporteGR2 = dsReporteGR.Tables[2].AsEnumerable()
                                                    .Select(n => new ReporteGR2
                                                    {
                                                        Articulo = string.IsNullOrEmpty(n.Field<string>("Articulo")) ? "" : n.Field<string>("Articulo"),
                                                        CodigoEPR = string.IsNullOrEmpty(n.Field<string>("Codigo_EPR")) ? "" : n.Field<string>("Codigo_EPR"),
                                                        Descripcion = string.IsNullOrEmpty(n.Field<string>("Descripcion")) ? "" : n.Field<string>("Descripcion"),
                                                        EnvaseDescripcion = string.IsNullOrEmpty(n.Field<string>("Envase_Descripcion")) ? "" : n.Field<string>("Envase_Descripcion"),
                                                        EstatusProducto = string.IsNullOrEmpty(n.Field<string>("Estatus_Producto")) ? "" : n.Field<string>("Estatus_Producto"),
                                                        Mecanica = string.IsNullOrEmpty(n.Field<string>("Mecanica")) ? "" : n.Field<string>("Mecanica"),
                                                        Segmento = n.Field<string>("Segmento")
                                                    }).ToList();
                }
                
            }
            catch(Exception ex)
            {
                throw ex;
            }


            return grupoReporteGR;
        }
        public DataSet MostrarReporteCircular(int IdCampania)
        {
            string spName = "MostrarCircular";
            DataSet dsReporteCircular = new DataSet();

            SqlParameter Parametro = new SqlParameter();
            List<SqlParameter> ListParametros = new List<SqlParameter>();

            try
            {
                //DATOS RENTABILIDAD
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campania";
                Parametro.Value = IdCampania;
                ListParametros.Add(Parametro);

                dsReporteCircular = base.ExecuteDataSet(spName, ListParametros);

            }
            catch (Exception ex)
            {
                throw ex;
            }


            return dsReporteCircular;
        }
        public List<EntidadesCampanasPPG.Modelo.Campana> MostrarCampanaBitacoraDetalle(EntidadesCampanasPPG.Modelo.Campana Campana)
        {
            List<EntidadesCampanasPPG.Modelo.Campana> ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

            const string spNameCampana = "MostrarBitacoraDetalleCampania";

            SqlParameter Parametro = new SqlParameter();
            List<SqlParameter> ListParametrosCampana = new List<SqlParameter>();

            DataSet dsCampana = new DataSet();

            DataTable dtCampana = new DataTable();
            DataTable dtBitacora = new DataTable();
            DataTable dtCronograma = new DataTable();
            DataTable dtParticipacionKroma = new DataTable();
            DataTable dtPublicidad = new DataTable();
            DataTable dtEscenarioRegalo = new DataTable();
            DataTable dtEscenarioMultiplo = new DataTable();
            DataTable dtEscenarioDescuento = new DataTable();
            DataTable dtEscenarioVolumen = new DataTable();
            DataTable dtEscenarioKit = new DataTable();
            DataTable dtEscenarioCombo = new DataTable();
            DataTable dtTienda = new DataTable();
            DataTable dtRentabilidad = new DataTable();

            List<Bitacora> ListBitacora = new List<Bitacora>();
            List<Cronograma> ListCronograma = new List<Cronograma>();
            List<Kroma> ListKroma = new List<Kroma>();
            List<Publicidad> ListPublicidad = new List<Publicidad>();
            List<MecanicaRegalo> ListMecanicaRegalo = new List<MecanicaRegalo>();
            List<MecanicaMultiplo> ListMecanicaMultiplo = new List<MecanicaMultiplo>();
            List<MecanicaDescuento> ListMecanicaDescuento = new List<MecanicaDescuento>();
            List<MecanicaVolumen> ListMecanicaVolumen = new List<MecanicaVolumen>();
            List<MecanicaKit> ListMecanicaKit = new List<MecanicaKit>();
            List<MecanicaCombo> ListMecanicaCombo = new List<MecanicaCombo>();
            List<Tienda> ListTienda = new List<Tienda>();
            List<Rentabilidad> ListRentabilidad = new List<Rentabilidad>();

            try
            {
                //DATOS CAMPANA
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ClvCampania";
                Parametro.Value = Campana.Title;
                ListParametrosCampana.Add(Parametro);

                dsCampana = base.ExecuteDataSet(spNameCampana, ListParametrosCampana);

                if(dsCampana.Tables.Count > 0)
                {
                    //INFORMACION CAMPANA
                    if(dsCampana.Tables[0].Rows.Count > 0)
                    {
                        ListCampana = dsCampana.Tables[0].AsEnumerable()
                                .Select(row => new EntidadesCampanasPPG.Modelo.Campana
                                {
                                    IdCampana = row.Field<int?>("ID").GetValueOrDefault(),
                                    Title = string.IsNullOrEmpty(row.Field<string>("Camp_Number")) ? "" : row.Field<string>("Camp_Number"),
                                    NombreCampa = string.IsNullOrEmpty(row.Field<string>("Nombre_Camp")) ? "" : row.Field<string>("Nombre_Camp"),
                                    RegistraCampa = string.IsNullOrEmpty(row.Field<string>("Nombre_Usuario")) ? "" : row.Field<string>("Nombre_Usuario"),
                                    PPGIDRegistraCampa = string.IsNullOrEmpty(row.Field<string>("PPGID")) ? "" : row.Field<string>("PPGID"),
                                    LiderCampa = string.IsNullOrEmpty(row.Field<string>("Lider_Campania")) ? "" : row.Field<string>("Lider_Campania"),
                                    PPGIDLiderCampa = string.IsNullOrEmpty(row.Field<string>("PPGID_Lider")) ? "" : row.Field<string>("PPGID_Lider"),
                                    FechaInicioSubCanal = row.Field<DateTime?>("Fecha_Inicio_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Inicio_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaFinSubCanal = row.Field<DateTime?>("Fecha_Fin_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Fin_SubCanal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaInicioPublico = row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Inicio_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaFinPublico = row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Fin_Publico").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaDocumento = row.Field<DateTime?>("Fecha_Creacion").GetValueOrDefault().ToString("dd/MM/yyyy") == "01/01/0001" ? "--/--/----" : row.Field<DateTime?>("Fecha_Creacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    IdMoneda = row.Field<int?>("ID_Moneda").GetValueOrDefault(),
                                    Moneda = string.IsNullOrEmpty(row.Field<string>("Moneda")) ? "" : row.Field<string>("Moneda"),
                                    IdTipoCampa = row.Field<int?>("ID_TipoCamp").GetValueOrDefault(),
                                    TipoCampa = string.IsNullOrEmpty(row.Field<string>("TipoCamp")) ? "" : row.Field<string>("TipoCamp"),
                                    IdAlcanceTerritorial = row.Field<int?>("ID_Alcance").GetValueOrDefault(),
                                    AlcanceTerritorial = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance"),
                                    IdTipoSell = row.Field<int?>("ID_TipoSell").GetValueOrDefault(),
                                    TipoSell = string.IsNullOrEmpty(row.Field<string>("TipoSell")) ? "" : row.Field<string>("TipoSell"),
                                    CampaExpress = row.Field<bool?>("Express").GetValueOrDefault(),
                                    IdNegocioLider = row.Field<int?>("ID_Negocio_Lider").GetValueOrDefault(),
                                    NegocioLider = string.IsNullOrEmpty(row.Field<string>("LiderNegocio")) ? "" : row.Field<string>("LiderNegocio"),
                                    IdSubcanal = row.Field<int?>("ID_Subcanal").GetValueOrDefault(),
                                    Subcanal = string.IsNullOrEmpty(row.Field<string>("Subcanal")) ? "" : row.Field<string>("Subcanal"),
                                    IdEstatus = row.Field<int?>("ID_Estatus").GetValueOrDefault(),
                                    Estatus = string.IsNullOrEmpty(row.Field<string>("Estatus")) ? "" : row.Field<string>("Estatus"),
                                    Status = string.IsNullOrEmpty(row.Field<string>("EstatusCat")) ? "" : row.Field<string>("EstatusCat"),
                                    TipoSubCanal = string.IsNullOrEmpty(row.Field<string>("TipoSubCanal")) ? "" : row.Field<string>("TipoSubCanal")
                                }).ToList();
                    }

                    //INFORMACION BITACORA
                    if (dsCampana.Tables[1].Rows.Count > 0)
                    {
                        ListBitacora = dsCampana.Tables[1].AsEnumerable()
                                .Select(row => new EntidadesCampanasPPG.Modelo.Bitacora
                                {
                                    IDCampania = row.Field<int?>("ID").GetValueOrDefault(),
                                    ClaveCampania = string.IsNullOrEmpty(row.Field<string>("Camp_Number")) ? "" : row.Field<string>("Camp_Number"),
                                    NombreCampania = string.IsNullOrEmpty(row.Field<string>("Nombre_Camp")) ? "" : row.Field<string>("Nombre_Camp"),
                                    LiderCampania = string.IsNullOrEmpty(row.Field<string>("Lider_Campania")) ? "" : row.Field<string>("Lider_Campania"),
                                    Estatus = string.IsNullOrEmpty(row.Field<string>("Estatus")) ? "" : row.Field<string>("Estatus"),
                                    Comentario = string.IsNullOrEmpty(row.Field<string>("Comentario")) ? "" : row.Field<string>("Comentario"),
                                    FechaCreacion = row.Field<DateTime?>("FechaCreacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaModificacion = row.Field<DateTime?>("FechaModificacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    PPGID = string.IsNullOrEmpty(row.Field<string>("PPGID")) ? "" : row.Field<string>("PPGID"),
                                    Usuario = string.IsNullOrEmpty(row.Field<string>("NombreUsuario")) ? "" : row.Field<string>("NombreUsuario"),
                                    IdTarea = row.Field<int?>("ID_Tarea").GetValueOrDefault(),
                                    TipoFlujo = row.Field<int?>("TipoFlujo").GetValueOrDefault(),
                                    CorreoResponsable = string.IsNullOrEmpty(row.Field<string>("CorreoUsuario")) ? "" : row.Field<string>("CorreoUsuario"),
                                    Completado = row.Field<int?>("Completado").GetValueOrDefault(),
                                    Actividad = string.IsNullOrEmpty(row.Field<string>("Actividad")) ? "" : row.Field<string>("Actividad"),
                                    FechaInicioCrono = row.Field<DateTime?>("FechaInicioCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaInicioRealCrono = row.Field<DateTime?>("FechaInicioRealCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaFinCrono = row.Field<DateTime?>("FechaFinCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    FechaFinRealCrono = row.Field<DateTime?>("FechaFinRealCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    PPGIDCrono = string.IsNullOrEmpty(row.Field<string>("PPGIDCrono")) ? "" : row.Field<string>("PPGIDCrono"),
                                    PPGID2Crono = string.IsNullOrEmpty(row.Field<string>("PPGID2Crono")) ? "" : row.Field<string>("PPGID2Crono"),
                                    CorreoCrono = string.IsNullOrEmpty(row.Field<string>("CorreoCrono")) ? "" : row.Field<string>("CorreoCrono"),
                                    CorreoCrono2 = string.IsNullOrEmpty(row.Field<string>("CorreoCrono2")) ? "" : row.Field<string>("CorreoCrono2"),
                                    NombreResponsable = string.IsNullOrEmpty(row.Field<string>("NombreResponsable")) ? "" : row.Field<string>("NombreResponsable"),
                                    NombreResponsable2 = string.IsNullOrEmpty(row.Field<string>("NombreResponsable_2")) ? "" : row.Field<string>("NombreResponsable_2")
                                }).ToList();
                    }

                    //INFORMACION CRONOGRAMA
                    if (dsCampana.Tables[2].Rows.Count > 0)
                    {
                        ListCronograma = dsCampana.Tables[2].AsEnumerable()
                                .Select(row => new Cronograma
                                {
                                    ID = row.Field<int?>("ID").GetValueOrDefault(),
                                    IDCampania = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                    IDPadre = row.Field<int?>("ID_Padre").GetValueOrDefault(),
                                    Padre = row.Field<bool?>("Padre").GetValueOrDefault(),
                                    IDTarea = row.Field<int?>("ID_Tarea").GetValueOrDefault(),
                                    Actividad = string.IsNullOrEmpty(row.Field<string>("Actividad")) ? "" : row.Field<string>("Actividad"),
                                    PorcentajeUsuario = row.Field<decimal?>("PorcentajeUsuario").GetValueOrDefault(),
                                    PorcentajeSistema = row.Field<decimal?>("PorcentajeSistema").GetValueOrDefault(),
                                    PorcentajeUsuarioEsfuerzo = row.Field<decimal?>("PorcentajeUsuEsfuerzo").GetValueOrDefault(),
                                    PorcentajeSistemaReal = row.Field<decimal?>("PorcentajeSistemaReal").GetValueOrDefault(),
                                    PorcentajeDiferencia = row.Field<decimal?>("PorcentajeDiferencia").GetValueOrDefault(),
                                    Predecesor = string.IsNullOrEmpty(row.Field<string>("Predecesor")) ? "" : row.Field<string>("Predecesor"),
                                    Duracion = row.Field<decimal?>("Duracion").GetValueOrDefault(),
                                    StrFechaInicio = row.Field<DateTime?>("FechaInicio").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    StrFechaFin = row.Field<DateTime?>("FechaFin").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    StrFechaInicioReal = row.Field<DateTime?>("FechaInicioReal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    StrFechaFinReal = row.Field<DateTime?>("FechaFinReal").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    TiempoOptimista = row.Field<int?>("TiempoOptimista").GetValueOrDefault(),
                                    TiempoPesimista = row.Field<int?>("TiempoPesimista").GetValueOrDefault(),
                                    NombreResponsable = string.IsNullOrEmpty(row.Field<string>("NombreResponsable")) ? "" : row.Field<string>("NombreResponsable"),
                                    Correo = string.IsNullOrEmpty(row.Field<string>("Correo")) ? "" : row.Field<string>("Correo"),
                                    PPGID = string.IsNullOrEmpty(row.Field<string>("PPGID")) ? "" : row.Field<string>("PPGID"),
                                    NombreResponsable_2 = string.IsNullOrEmpty(row.Field<string>("NombreResponsable_2")) ? "" : row.Field<string>("NombreResponsable_2"),
                                    Correo_2 = string.IsNullOrEmpty(row.Field<string>("Correo_2")) ? "" : row.Field<string>("Correo_2"),
                                    PPGID_2 = string.IsNullOrEmpty(row.Field<string>("PPGID_2")) ? "" : row.Field<string>("PPGID_2"),
                                    Comentario = string.IsNullOrEmpty(row.Field<string>("Comentario")) ? "" : row.Field<string>("Comentario"),
                                    VersionCronograma = row.Field<int?>("VersionCronograma").GetValueOrDefault(),
                                    StrFechaHoy = row.Field<DateTime?>("FechaHoy").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    StrFechaCreacion = row.Field<DateTime?>("FechaCreacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    StrFechaModificacion = row.Field<DateTime?>("FechaModificacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                    UsuarioCreacion = string.IsNullOrEmpty(row.Field<string>("UsuarioCreacion")) ? "" : row.Field<string>("UsuarioCreacion"),
                                    UsuarioModificacion = string.IsNullOrEmpty(row.Field<string>("UsuarioModificacion")) ? "" : row.Field<string>("UsuarioModificacion"),
                                    TipoFlujo = string.IsNullOrEmpty(row.Field<string>("TipoFlujo")) ? "" : row.Field<string>("TipoFlujo"),
                                    Incluir = string.IsNullOrEmpty(row.Field<string>("Incluir")) ? "" : row.Field<string>("Incluir"),
                                    EstatusEnvio = row.Field<int?>("EstatusEnvio").GetValueOrDefault(),
                                    Update = false
                                }).ToList();
                    }

                    //INFORMACION PARTICIPACION KROMA
                    if (dsCampana.Tables[3].Rows.Count > 0)
                    {
                        ListKroma = dsCampana.Tables[3].AsEnumerable()
                            .Select(row => new Kroma
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Base = string.IsNullOrEmpty(row.Field<string>("Base")) ? "" : row.Field<string>("Base"),
                                Linea = string.IsNullOrEmpty(row.Field<string>("Linea")) ? "" : row.Field<string>("Linea"),
                                SobrePrecioKroma = string.IsNullOrEmpty(row.Field<string>("SobrePrecioKroma")) ? "" : row.Field<string>("SobrePrecioKroma"),
                                Porcentaje = row.Field<decimal?>("Porcentaje").GetValueOrDefault()                               
                            }).ToList();
                    }

                    //INFORMACION PUBLICIDAD
                    if (dsCampana.Tables[4].Rows.Count > 0)
                    {
                        ListPublicidad = dsCampana.Tables[4].AsEnumerable()
                            .Select(row => new Publicidad
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                IdPublicidad = row.Field<int?>("ID_Publicidad").GetValueOrDefault(),
                                PublicidadDescripcion = string.IsNullOrEmpty(row.Field<string>("Publicidad")) ? "" : row.Field<string>("Publicidad"),
                                Comentario = string.IsNullOrEmpty(row.Field<string>("Comentario")) ? "" : row.Field<string>("Comentario"),
                                Monto = row.Field<decimal?>("Cantidad").GetValueOrDefault(),
                                MontoAnterior = row.Field<decimal?>("CantidadAnterior").GetValueOrDefault(),
                            }).ToList();
                    }

                    //INFORMACION ESCENARIO REGALO
                    if (dsCampana.Tables[5].Rows.Count > 0)
                    {
                        ListMecanicaRegalo = dsCampana.Tables[5].AsEnumerable()
                           .Select(row => new MecanicaRegalo
                           {
                               IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                               Familia = string.IsNullOrEmpty(row.Field<string>("Familia")) ? "" : row.Field<string>("Familia"),
                               SKU = string.IsNullOrEmpty(row.Field<string>("Articulo")) ? "" : row.Field<string>("Articulo"),
                               Descripcion = string.IsNullOrEmpty(row.Field<string>("Descripcion")) ? "" : row.Field<string>("Descripcion"),
                               Tipo = string.IsNullOrEmpty(row.Field<string>("Tipo")) ? "" : row.Field<string>("Tipo"),
                               Grupo_Regalo = row.Field<int?>("Grupo_Regalo").GetValueOrDefault(),
                               NumeroHijo = row.Field<int?>("Numero_Hijo").GetValueOrDefault(),
                               Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance"),
                               Capacidad = string.IsNullOrEmpty(row.Field<string>("Capacidad")) ? "" : row.Field<string>("Capacidad"),
                               Dinamica = string.IsNullOrEmpty(row.Field<string>("Dinamica")) ? "" : row.Field<string>("Dinamica"),
                               VLitrosAnioAnt = row.Field<decimal?>("Ventas_Litros_Anio_Anterior").GetValueOrDefault(),
                               PLitrosSinCamp = row.Field<decimal?>("Presupuesto_Litros_Sin_Campania").GetValueOrDefault(),
                               PLitrosConCamp = row.Field<decimal?>("Presupuesto_Litros_Con_Campania").GetValueOrDefault()
                           }).ToList();
                    }

                    //INFORMACION ESCENARIO MULTIPLO
                    if (dsCampana.Tables[6].Rows.Count > 0)
                    {
                        ListMecanicaMultiplo = dsCampana.Tables[6].AsEnumerable()
                            .Select(row => new MecanicaMultiplo
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Familia = string.IsNullOrEmpty(row.Field<string>("Familia")) ? "" : row.Field<string>("Familia"),
                                SKU = string.IsNullOrEmpty(row.Field<string>("SKU")) ? "" : row.Field<string>("SKU"),
                                Descripcion = string.IsNullOrEmpty(row.Field<string>("Descripcion")) ? "" : row.Field<string>("Descripcion"),
                                Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance"),
                                Capacidad = string.IsNullOrEmpty(row.Field<string>("Capacidad")) ? "" : row.Field<string>("Capacidad"),
                                Dinamica = string.IsNullOrEmpty(row.Field<string>("Dinamica")) ? "" : row.Field<string>("Dinamica"),
                                Multiplo_Padre = row.Field<int?>("MultiploPadre").GetValueOrDefault(),
                                Multiplo_Hijo = row.Field<int?>("MultiploHijo").GetValueOrDefault(),
                                Punto_Venta = string.IsNullOrEmpty(row.Field<string>("Punto_Venta")) ? "" : row.Field<string>("Punto_Venta"),
                                VLitrosAnioAnt = row.Field<decimal?>("Venta_Anio_Anterior").GetValueOrDefault(),
                                PLitrosSinCamp = row.Field<decimal?>("Presupuesto_Sin_Campania").GetValueOrDefault(),
                                PLitrosConCamp = row.Field<decimal?>("Presupuesto_Con_Campania").GetValueOrDefault()
                            }).ToList();
                    }

                    //INFORMACION ESCENARIO DESCUENTO
                    if (dsCampana.Tables[7].Rows.Count > 0)
                    {
                        ListMecanicaDescuento = dsCampana.Tables[7].AsEnumerable()
                            .Select(row => new MecanicaDescuento
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Familia = string.IsNullOrEmpty(row.Field<string>("Familia")) ? "" : row.Field<string>("Familia"),
                                SKU = string.IsNullOrEmpty(row.Field<string>("Articulo")) ? "" : row.Field<string>("Articulo"),
                                Descripcion = string.IsNullOrEmpty(row.Field<string>("Descripcion")) ? "" : row.Field<string>("Descripcion"),
                                Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance"),
                                Capacidad = string.IsNullOrEmpty(row.Field<string>("Capacidad")) ? "" : row.Field<string>("Capacidad"),
                                Dinamica = string.IsNullOrEmpty(row.Field<string>("Dinamica")) ? "" : row.Field<string>("Dinamica"),
                                Porcentaje = row.Field<decimal?>("Porcentaje").GetValueOrDefault(),
                                Importe = row.Field<decimal?>("Importe").GetValueOrDefault(),
                                VLitrosAnioAnt = row.Field<decimal?>("VentaAñoAnterior").GetValueOrDefault(),
                                PLitrosSinCamp = row.Field<decimal?>("PresupuestoSinCampaña").GetValueOrDefault(),
                                PLitrosConCamp = row.Field<decimal?>("PresupuestoConCampaña").GetValueOrDefault()
                            }).ToList();
                    }

                    //INFORMACION ESCENRIO VOLUMEN
                    if (dsCampana.Tables[8].Rows.Count > 0)
                    {
                        ListMecanicaVolumen = dsCampana.Tables[8].AsEnumerable()
                            .Select(row => new MecanicaVolumen
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Familia = string.IsNullOrEmpty(row.Field<string>("Familia")) ? "" : row.Field<string>("Familia"),
                                SKU = string.IsNullOrEmpty(row.Field<string>("Articulo")) ? "" : row.Field<string>("Articulo"),
                                Descripcion = string.IsNullOrEmpty(row.Field<string>("Descripcion")) ? "" : row.Field<string>("Descripcion"),
                                Bloque = row.Field<int?>("Bloque").GetValueOrDefault(),
                                Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance"),
                                Capacidad = string.IsNullOrEmpty(row.Field<string>("Capacidad")) ? "" : row.Field<string>("Capacidad"),
                                Dinamica = string.IsNullOrEmpty(row.Field<string>("Dinamica")) ? "" : row.Field<string>("Dinamica"),
                                De = row.Field<decimal?>("VentaDesde").GetValueOrDefault(),
                                Hasta = row.Field<decimal?>("VentaHasta").GetValueOrDefault(),
                                Descuento = row.Field<decimal?>("PorcentajeDescuento").GetValueOrDefault(),
                                VLitrosAnioAnt = row.Field<decimal?>("VentaAnioAnterior").GetValueOrDefault(),
                                PLitrosSinCamp = row.Field<decimal?>("PresupuestoSinCampania").GetValueOrDefault(),
                                PLitrosConCamp = row.Field<decimal?>("PresupuestoConCampania").GetValueOrDefault()
                            }).ToList();
                    }

                    //INFORMACION ESCENARIO KIT
                    if (dsCampana.Tables[9].Rows.Count > 0)
                    {
                        ListMecanicaKit = dsCampana.Tables[9].AsEnumerable()
                            .Select(row => new MecanicaKit
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Familia = string.IsNullOrEmpty(row.Field<string>("Familia")) ? "" : row.Field<string>("Familia"),
                                SKU = string.IsNullOrEmpty(row.Field<string>("Articulo")) ? "" : row.Field<string>("Articulo"),
                                Descripcion = string.IsNullOrEmpty(row.Field<string>("Descripcion")) ? "" : row.Field<string>("Descripcion"),
                                Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance"),
                                Capacidad = string.IsNullOrEmpty(row.Field<string>("Capacidad")) ? "" : row.Field<string>("Capacidad"),
                                Dinamica = string.IsNullOrEmpty(row.Field<string>("Dinamica")) ? "" : row.Field<string>("Dinamica"),
                                Porcentaje = row.Field<decimal?>("Porcentaje").GetValueOrDefault(),
                                Importe = row.Field<decimal?>("Importe").GetValueOrDefault(),
                                VLitrosAnioAnt = row.Field<decimal?>("VentaAnioAnterior").GetValueOrDefault(),
                                PLitrosSinCamp = row.Field<decimal?>("PresupuestoSinCampania").GetValueOrDefault(),
                                PLitrosConCamp = row.Field<decimal?>("PresupuestoConCampania").GetValueOrDefault()
                            }).ToList();
                    }

                    //INFORMACION ESCENARIO COMBO
                    if (dsCampana.Tables[10].Rows.Count > 0)
                    {
                        ListMecanicaCombo = dsCampana.Tables[10].AsEnumerable()
                            .Select(row => new MecanicaCombo
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                Familia = string.IsNullOrEmpty(row.Field<string>("Familia")) ? "" : row.Field<string>("Familia"),
                                SKU = string.IsNullOrEmpty(row.Field<string>("Articulo")) ? "" : row.Field<string>("Articulo"),
                                Descripcion = string.IsNullOrEmpty(row.Field<string>("Descripcion")) ? "" : row.Field<string>("Descripcion"),
                                Tipo = string.IsNullOrEmpty(row.Field<string>("TipoArticulo")) ? "" : row.Field<string>("TipoArticulo"),
                                Grupo_Combo = row.Field<int?>("Grupo_Combo").GetValueOrDefault(),
                                NumeroPadre = row.Field<int?>("Numero_Padre").GetValueOrDefault(),
                                NumeroHijo = row.Field<int?>("Numero_Hijo").GetValueOrDefault(),
                                Alcance = string.IsNullOrEmpty(row.Field<string>("Alcance")) ? "" : row.Field<string>("Alcance"),
                                Capacidad = string.IsNullOrEmpty(row.Field<string>("Capacidad")) ? "" : row.Field<string>("Capacidad"),
                                Dinamica = string.IsNullOrEmpty(row.Field<string>("Dinamica")) ? "" : row.Field<string>("Dinamica"),
                                VLitrosAnioAnt = row.Field<decimal?>("VentaAnioAnterior").GetValueOrDefault(),
                                PLitrosSinCamp = row.Field<decimal?>("PresupuestoSinCampania").GetValueOrDefault(),
                                PLitrosConCamp = row.Field<decimal?>("PresupuestoConCampania").GetValueOrDefault()
                            }).ToList();
                    }

                    //INFORMACION TIENDA
                    if (dsCampana.Tables[11].Rows.Count > 0)
                    {
                        ListTienda = dsCampana.Tables[11].AsEnumerable()
                            .Select(row => new Tienda
                            {
                                IdCampana = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                IdTienda = row.Field<int?>("ID_Tienda").GetValueOrDefault(),
                                BillTo = string.IsNullOrEmpty(row.Field<string>("Bill_To")) ? "" : row.Field<string>("Bill_To"),
                                CustomerName = string.IsNullOrEmpty(row.Field<string>("Customer_Name")) ? "" : row.Field<string>("Customer_Name"),
                                Region = string.IsNullOrEmpty(row.Field<string>("Region")) ? "" : row.Field<string>("Region"),
                                DescripcionRegion = string.IsNullOrEmpty(row.Field<string>("Descripcion_Region")) ? "" : row.Field<string>("Descripcion_Region"),
                                DescripcionZona = string.IsNullOrEmpty(row.Field<string>("Descripcion_Zona")) ? "" : row.Field<string>("Descripcion_Zona"),
                                Segmento = string.IsNullOrEmpty(row.Field<string>("Segmento")) ? "" : row.Field<string>("Segmento"),
                                ClaveSobreprecio = string.IsNullOrEmpty(row.Field<string>("Clave_Sobreprecio")) ? "" : row.Field<string>("Clave_Sobreprecio"),
                            }).ToList();
                    }

                    //INFORMACION RENTABILIDAD
                    if (dsCampana.Tables[12].Rows.Count > 0)
                    {
                        ListRentabilidad = dsCampana.Tables[12].AsEnumerable()
                                .Select(row => new Rentabilidad
                                {
                                    IdCampania = row.Field<int?>("ID_Campania").GetValueOrDefault(),
                                    PorcentajeCompania = row.Field<decimal?>("PorcRentabilidad").GetValueOrDefault(),
                                    Comentario = string.IsNullOrEmpty(row.Field<string>("Comentario")) ? "" : row.Field<string>("Comentario"),
                                    PorcentajeNecesario = row.Field<decimal?>("PorcNecesario").GetValueOrDefault(),
                                }).ToList();
                    }


                    //ORDENAR LA INFORMACION DE CAMPAÑAS
                    foreach(EntidadesCampanasPPG.Modelo.Campana campana in ListCampana)
                    {
                        campana.ListBitacora = ListBitacora.Where(n => n.IDCampania == campana.IdCampana).ToList();

                        campana.ListaCronograma = ListCronograma.Where(n => n.IDCampania == campana.IdCampana).ToList();

                        campana.ListaKroma = ListKroma.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListaPublicidad = ListPublicidad.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListMecanicaRegalo = ListMecanicaRegalo.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListMecanicaMultiplo = ListMecanicaMultiplo.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListMecanicaDescuento = ListMecanicaDescuento.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListMecanicaVolumen = ListMecanicaVolumen.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListMecanicaKit = ListMecanicaKit.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListMecanicaCombo = ListMecanicaCombo.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListTienda = ListTienda.Where(n => n.IdCampana == campana.IdCampana).ToList();

                        campana.ListRentabilidad = ListRentabilidad.Where(n => n.IdCampania == campana.IdCampana).ToList();
                    }

                }

            }
            catch(Exception ex)
            {
                throw ex;
            }

            return ListCampana;
        }
        public List<GastoPlantaHistorico> MostrarGastoPlantaHistorico(GastoPlantaHistorico gastoPlantaHistorico)
        {
            List<GastoPlantaHistorico> ListGastoPlantaHistorico = new List<GastoPlantaHistorico>();

            string spName = "MostrarGastoPLanta";

            SqlParameter Parametro = new SqlParameter();
            List<SqlParameter> ListParametros = new List<SqlParameter>();

            DataTable dtGastoPlantaHistorico = new DataTable();

            try
            {
                //DATOS GASTO PLANTA HISTORICO
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Periodo";
                Parametro.Value = gastoPlantaHistorico.Periodo;
                ListParametros.Add(Parametro);

                dtGastoPlantaHistorico = ExecuteDataTable(spName, ListParametros);

                ListGastoPlantaHistorico = dtGastoPlantaHistorico.AsEnumerable()
                                            .Select(n => new GastoPlantaHistorico
                                            {
                                                Id = n.Field<int?>("ID").GetValueOrDefault(),
                                                IdPlanta = n.Field<int?>("ID_Planta").GetValueOrDefault(),
                                                Planta = string.IsNullOrEmpty(n.Field<string>("Planta")) ? "" : n.Field<string>("Planta"),
                                                Porcentaje = n.Field<decimal?>("Porcentaje").GetValueOrDefault(),
                                                Periodo = string.IsNullOrEmpty(n.Field<string>("Periodo")) ? "" : n.Field<string>("Periodo"),
                                                UsuarioCreacion = string.IsNullOrEmpty(n.Field<string>("NombreUsuario")) ? "" : n.Field<string>("NombreUsuario"),
                                                UsuarioModificacion = string.IsNullOrEmpty(n.Field<string>("UsuarioModificacion")) ? "" : n.Field<string>("UsuarioModificacion")
                                            }).ToList();
            }
            catch(Exception ex)
            {
                throw ex;
            }

            return ListGastoPlantaHistorico;
        }
    }
}

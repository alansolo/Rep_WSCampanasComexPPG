using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatosCampanasPPG.BitacoraDAT
{
    public class BitacoraDAT:BD
    {
        public int GuardarBitacora(EntidadesCampanasPPG.Modelo.Bitacora Bitacora)
        {
            const string spName = "InsertBitacora";
            int resultado = 0;
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Campania";
                Parametro.DbType = DbType.Int32;
                Parametro.Direction = ParameterDirection.InputOutput;
                Parametro.Value = Bitacora.IDCampania;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Clv_Campania";
                Parametro.Value = Bitacora.ClaveCampania;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Estatus";
                Parametro.Value = Bitacora.Estatus;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@NombreUsuario";
                Parametro.Value = Bitacora.Usuario;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Comentario";
                Parametro.Value = Bitacora.Comentario;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@PPGID";
                Parametro.Value = Bitacora.PPGID;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Tarea";
                Parametro.Value = Bitacora.IdTarea;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@TipoFlujo";
                Parametro.Value = Bitacora.TipoFlujo;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Correo";
                Parametro.Value = Bitacora.CorreoResponsable;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Completado";
                Parametro.Value = Bitacora.Completado;
                ListParametros.Add(Parametro);


                resultado = base.ExecuteNonQueryBitacora(spName, ListParametros);

                Parametro = ListParametros.Where(n => n.ParameterName == "@ID_Campania").FirstOrDefault();

                Bitacora.IDCampania = Convert.ToInt32(Parametro.Value);

                return resultado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable MostrarBitacora(EntidadesCampanasPPG.Modelo.Bitacora Bitacora)
        {
            const string spName = "MostrarBitacora";
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Campania";
                    Parametro.Value = Bitacora.IDCampania;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Clv_Campania";
                    Parametro.Value = Bitacora.ClaveCampania;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Nombre_Camp";
                    Parametro.Value = Bitacora.NombreCampania;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Lider_Campania";
                    Parametro.Value = Bitacora.LiderCampania;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Estatus";
                    Parametro.Value = Bitacora.Estatus;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaInicio";
                    if(Bitacora.FechaInicio.ToString("dd/MM/yyyy") != "01/01/0001")
                    {
                        Parametro.Value = Bitacora.FechaInicio;
                    }
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@FechaFin";
                    if (Bitacora.FechaFin.ToString("dd/MM/yyyy") != "01/01/0001")
                    {
                        Parametro.Value = Bitacora.FechaFin;
                    }
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@ID_Tarea";
                    if (Bitacora.IdTarea > 0)
                    {
                        Parametro.Value = Bitacora.IdTarea;
                    }
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Tipo_Flujo";
                    if (Bitacora.TipoFlujo > 0)
                    {
                        Parametro.Value = Bitacora.TipoFlujo;
                    }
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@CorreoUsuario";
                    Parametro.Value = Bitacora.CorreoResponsable;
                    ListParametros.Add(Parametro);

                    Parametro = new SqlParameter();
                    Parametro.ParameterName = "@Completado";
                    if (Bitacora.Completado > 0)
                {
                    Parametro.Value = Bitacora.Completado;
                }
                ListParametros.Add(Parametro);

                return base.ExecuteDataTable(spName, ListParametros);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatosCampanasPPG.Catalogo
{
    public class LdapDAT:BD
    {
        public DataTable GetDirectorioLdap(int id, string claveDirectorio)
        {
            const string spName = "MostrarDirectorioActivo";
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID";
                if (id > 0)
                {
                    Parametro.Value = id;
                }
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Clave";
                if (Parametro != null)
                {
                    Parametro.Value = claveDirectorio;
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

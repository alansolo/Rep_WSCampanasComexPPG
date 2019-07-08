using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatosCampanasPPG.Catalogo
{
    public class RolCronogramaDAT:BD
    {
        public DataTable GetRolCronograma(Nullable<int> id_rol, string rol_nombre)
        {
            const string spName = "Mostrar_Roles";
            List<SqlParameter> ListParametros = new List<SqlParameter>();
            SqlParameter Parametro;

            try
            {
                Parametro = new SqlParameter();
                Parametro.ParameterName = "@ID_Rol";
                Parametro.Value = id_rol;
                ListParametros.Add(Parametro);

                Parametro = new SqlParameter();
                Parametro.ParameterName = "@Rol_Nombre";
                Parametro.Value = rol_nombre;
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

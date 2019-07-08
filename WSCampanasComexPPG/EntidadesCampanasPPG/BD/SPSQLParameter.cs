using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.BD
{
    public class SpSqlParameter
    {
        string SpName { get; set; }

        List<SqlParameterGroup> ListSqlParameterGroup { get; set; }
    }
}

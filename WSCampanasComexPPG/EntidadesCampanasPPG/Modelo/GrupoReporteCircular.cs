using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class GrupoReporteCircular
    {
        public DataSet DsReporteCircular { get; set; }
        public string Mensaje { get; set; }
        public int OK { get; set; }
    }
}

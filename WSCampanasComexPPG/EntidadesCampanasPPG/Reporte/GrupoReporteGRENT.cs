using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class GrupoReporteGRENT
    {
        public List<ReporteGR1> ListReporteGR1 { get; set; }
        public List<ReporteGR2> ListReporteGR2 { get; set; }
        public string Mensaje { get; set; }
        public int OK { get; set; }
    }
}

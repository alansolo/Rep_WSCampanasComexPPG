using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class ReporteCEOMecanica
    {
        public int IdCampania { get; set; }
        public string ClaveCampania { get; set; }
        public string Articulo { get; set; }
        public string Cantidad { get; set; }
        public int Periodo { get; set; }
        public string Alcance { get; set; }
        public string Mecanica { get; set; }

    }
}

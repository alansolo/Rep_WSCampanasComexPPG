using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class ReporteCEOPublicidad
    {
        public int IdCampania { get; set; }
        public string ClaveCampania { get; set; }
        public string Publicidad { get; set; }
        public decimal PeriodoAnterior { get; set; }
        public decimal PeriodoActual { get; set; }

    }
}

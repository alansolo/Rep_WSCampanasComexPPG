using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class ReporteMKT
    {
        public string Articulo { get; set; }
        public string Descripcion { get; set; }
        public decimal LitrosActualCC { get; set; }
        public decimal PiezasActualCC { get; set; }
        public decimal ImporteCPRecioLitros { get; set; }
        public decimal ImporteCPrecioPiezas { get; set; }
        public string Rentabilidad { get; set; }
        public decimal Litros { get; set; }
        public decimal Piezas { get; set; }
        public string Alcance { get; set; }

    }
}

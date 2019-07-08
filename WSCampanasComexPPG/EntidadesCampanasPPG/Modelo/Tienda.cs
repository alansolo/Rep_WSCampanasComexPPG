using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class Tienda
    {
        public int IdCampana { get; set; }
        public int IdTienda { get; set; }
        public string BillTo { get; set; }
        public string CustomerName { get; set; }
        public string Region { get; set; }
        public string DescripcionRegion { get; set; }
        public string DescripcionZona { get; set; }
        public string Segmento { get; set; }
        public string ClaveSobreprecio { get; set; }
        public string SubCanal { get; set; }
        public string Exclusion { get; set; }
    }
}

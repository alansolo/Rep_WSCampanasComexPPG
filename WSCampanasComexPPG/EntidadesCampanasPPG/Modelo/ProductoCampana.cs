using EntidadesCampanasPPG.Campana;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class ProductoCampana
    {
        public string Mensaje { get; set; }
        public List<string> ListMensaje { get; set; }
        //public List<ProductoLinea> ListProductoLinea { get; set; }
        public List<LineaFamilia> ListLineaFamilia { get; set; }
        public List<SKUValidacion> ListSKUValidacion { get; set; }
    }
}

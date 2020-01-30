using EntidadesCampanasPPG.Campana;
using EntidadesCampanasPPG.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Producto
{
    public class ProductoLineaENT
    {
        public string ClaveCampana { get; set; }
        public string UrlArchivo { get; set; }
        //public List<ProductoLinea> ListProductoLinea { get; set; }
        public List<LineaFamilia> ListLineaFamilia { get; set; }
        public List<string> ListMensaje { get; set; }
        public string Mensaje { get; set; }
        public List<SKUValidacion> ListSkuValidacion { get; set; }
    }
}

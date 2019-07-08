using EntidadesCampanasPPG.Modelo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Producto
{
    public class CronogramaENT
    {
        public string ClaveCampana { get; set; }
        public string UrlArchivo { get; set; }
        public List<Cronograma> ListActividad { get; set; }
        public string Mensaje { get; set; }
    }
}

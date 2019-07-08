using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EntidadesCampanasPPG.Modelo;

namespace EntidadesCampanasPPG.FlujoActividad
{
    public class FlujoActividadENT
    {
        public List<EntidadesCampanasPPG.Modelo.FlujoActividad> ListFlujoActividad = new List<Modelo.FlujoActividad>();
        public string Mensaje { get; set; }
        public int OK { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EntidadesCampanasPPG.Modelo;

namespace EntidadesCampanasPPG.Modelo
{
    public class CampanaENT
    {
        public Campana Campana { get; set; }
        public MostrarCronograma MostrarCronograma { get; set; }
        public string Mensaje { get; set; }
        public int OK { get; set; }

    }
}

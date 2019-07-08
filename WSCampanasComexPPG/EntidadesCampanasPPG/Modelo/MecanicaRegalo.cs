using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class MecanicaRegalo
    {
        public int IdCampana { get; set; }
        public string ClaveCampana { get; set; }
        public string Familia { get; set; }
        public string Submarca { get; set; }
        public string Tipo_Marca { get; set; }
        public string SKU { get; set; }
        public string Descripcion { get; set; }
        public string Tipo { get; set; }
        public int Grupo_Regalo { get; set; }
        public int NumeroHijo { get; set; }
        public string Alcance { get; set; }
        public string Capacidad { get; set; }
        public string Dinamica { get; set; }
        public decimal VLitrosAnioAnt { get; set; }
        public decimal PLitrosSinCamp { get; set; }
        public decimal PLitrosConCamp { get; set; }
        
    }
}

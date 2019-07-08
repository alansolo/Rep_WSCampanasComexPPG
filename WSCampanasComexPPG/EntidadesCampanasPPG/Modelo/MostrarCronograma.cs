using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class MostrarCronograma
    {
        //DATOS CRONOGRAMA
        [DataMember]
        public List<Cronograma> ListaCronograma { get; set; }
        [DataMember]
        public bool Correcto { get; set; }
        [DataMember]
        public bool CorrectoWarning { get; set; }
        [DataMember]
        public int OK { get; set; }
        [DataMember]
        public string Mensaje { get; set; }

        //DATOS CAMPAÑIA
        [DataMember]
        public string ClaveCampania { get; set; }
        [DataMember]
        public int IdCampania { get; set; }
        [DataMember]
        public string FechaInicio { get; set; }
        [DataMember]
        public string FechaFinal { get; set; }
    }
}

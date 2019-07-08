using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class FlujoActividadAprobador
    {
        [DataMember]
        public List<FlujoActividad> ListFlujoActividad { get; set; }
        //ERROR
        [DataMember]
        public string Mensaje { get; set; }
    }
}

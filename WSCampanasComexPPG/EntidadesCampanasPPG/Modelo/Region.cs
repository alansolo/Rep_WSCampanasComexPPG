using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class Region
    {
        [DataMember]
        public int IdCampania { get; set; }
        [DataMember]
        public string NombreRegion { get; set; }
        [DataMember]
        public string PPGID { get; set; }
        [DataMember]
        public string Usuario { get; set; }
    }
}

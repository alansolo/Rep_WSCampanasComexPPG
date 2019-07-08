using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class LineaFamilia
    {
        [DataMember]
        public string Producto { get; set; }
        [DataMember]
        public string Linea { get; set; }
        //[DataMember]
        //public string Region { get; set; }
        [DataMember]
        public string DF { get; set; }
        [DataMember]
        public string Validacion { get; set; }
    }
}

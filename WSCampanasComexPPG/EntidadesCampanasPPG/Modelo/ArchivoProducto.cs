using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class ArchivoProducto
    {
        [DataMember]
        public string ClaveCampana { get; set; }
        [DataMember]
        public string UrlArchivo { get; set; }
    }
}

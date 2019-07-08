using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class SubCanal
    {
        [DataMember]
        public int ID
        { get; set; }

        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public int IdSubCanal
        { get; set; }

        [DataMember]
        public string Descripcion
        { get; set; }

        [DataMember]
        public string PPGID
        { get; set; }

        [DataMember]
        public string NombreUsuario
        { get; set; }

        [DataMember]
        public DateTime FechaCreacion
        { get; set; }

        [DataMember]
        public DateTime FechaModificacion
        { get; set; }
    }
}

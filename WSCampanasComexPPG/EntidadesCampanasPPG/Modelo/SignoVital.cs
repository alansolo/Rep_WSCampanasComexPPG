using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class SignoVital
    {
        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public string Descripcion
        { get; set; }

        [DataMember]
        public DateTime FechaCreacion
        { get; set; }

    }
}
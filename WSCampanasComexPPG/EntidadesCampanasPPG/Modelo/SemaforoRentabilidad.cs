using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class SemaforoRentabilidad
    {
        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public string Descripcion
        { get; set; }

        [DataMember]
        public decimal De
        { get; set; }

        [DataMember]
        public decimal Hasta
        { get; set; }

    }

}

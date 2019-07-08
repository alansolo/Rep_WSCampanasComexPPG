using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class Kroma
    {
        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public string Base
        { get; set; }

        [DataMember]
        public string Linea
        { get; set; }

        [DataMember]
        public string SobrePrecioKroma
        { get; set; }

        [DataMember]
        public decimal Porcentaje
        { get; set; }

    }

}

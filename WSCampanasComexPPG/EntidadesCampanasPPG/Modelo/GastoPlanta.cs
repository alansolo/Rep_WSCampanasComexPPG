using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class GastoPlanta
    {
        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public string Planta
        { get; set; }

        [DataMember]
        public decimal AnioAnterior
        { get; set; }

        [DataMember]
        public decimal AnioActual
        { get; set; }

    }

}

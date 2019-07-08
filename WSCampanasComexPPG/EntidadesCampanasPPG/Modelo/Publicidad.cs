using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class Publicidad
    {
        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public int IdPublicidad
        { get; set; }

        [DataMember]
        public decimal Monto
        { get; set; }
        [DataMember]
        public decimal MontoAnterior
        { get; set; }
        [DataMember]
        public string PublicidadDescripcion
        { get; set; }
        [DataMember]
        public string Comentario
        { get; set; }

    }
}
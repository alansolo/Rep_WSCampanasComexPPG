using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class Producto
    {
        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public int IdLineaProducto
        { get; set; }

        [DataMember]
        public int IdFamiliaEstelar
        { get; set; }

        [DataMember]
        public int IdMecanica
        { get; set; }

        [DataMember]
        public decimal CantidadOdescuento
        { get; set; }

        [DataMember]
        public string CapacidadProducto
        { get; set; }

        [DataMember]
        public int Alcance
        { get; set; }

        [DataMember]
        public int IdRol
        { get; set; }

        [DataMember]
        public string Observaciones
        { get; set; }

        [DataMember]
        public string SistemaAplicacion
        { get; set; }
    }
}
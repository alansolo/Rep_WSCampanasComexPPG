using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class GastoPlantaHistorico
    {
        [DataMember]
        public int Id
        { get; set; }

        [DataMember]
        public int IdPlanta
        { get; set; }

        [DataMember]
        public string Planta
        { get; set; }

        [DataMember]
        public decimal Porcentaje
        { get; set; }

        [DataMember]
        public string Periodo
        { get; set; }

        [DataMember]
        public string UsuarioCreacion
        { get; set; }

        [DataMember]
        public string UsuarioModificacion
        { get; set; }

        [DataMember]
        public DateTime FechaCreacion
        { get; set; }

        [DataMember]
        public DateTime FechaModificacion
        { get; set; }
    }
}

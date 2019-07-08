using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class FlujoActividad
    {
        [DataMember]
        public int IdCampania { get; set; }
        [DataMember]
        public string ClaveCampania { get; set; }
        [DataMember]
        public string NombreCampania { get; set; }
        [DataMember]
        public string MailResponsable { get; set; }
        [DataMember]
        public string MailResponsable2 { get; set; }
        [DataMember]
        public DateTime FechaInicio { get; set; }
        [DataMember]
        public DateTime FechaFin { get; set; }
        [DataMember]
        public int IDTarea { get; set; }
        [DataMember]
        public string TxtTarea { get; set; }
        [DataMember]
        public int TipoFlujo { get; set; }
        [DataMember]
        public string IdDependiente { get; set; }
    }
}

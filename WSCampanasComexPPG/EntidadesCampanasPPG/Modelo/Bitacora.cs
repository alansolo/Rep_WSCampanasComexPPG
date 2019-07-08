using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    [DataContract]
    public class Bitacora
    {
        [DataMember]
        public int IDCampania { get; set; }
        [DataMember]
        public string ClaveCampania { get; set; }
        [DataMember]
        public string NombreCampania { get; set; }
        [DataMember]
        public string LiderCampania { get; set; }
        [DataMember]
        public string Estatus { get; set; }
        [DataMember]
        public string Comentario { get; set; }
        [DataMember]
        public string FechaCreacion { get; set; }
        [DataMember]
        public DateTime FechaInicio { get; set; }
        [DataMember]
        public DateTime FechaFin { get; set; }
        [DataMember]
        public string FechaModificacion { get; set; }
        [DataMember]
        public string PPGID { get; set; }
        [DataMember]
        public string Usuario { get; set; }
        [DataMember]
        public int IdTarea { get; set; }
        [DataMember]
        public int TipoFlujo { get; set; }
        [DataMember]
        public string CorreoResponsable { get; set; }
        [DataMember]
        public int Completado { get; set; }
        [DataMember]
        public string Actividad { get; set; }
        [DataMember]
        public string FechaInicioCrono { get; set; }
        [DataMember]
        public string FechaFinCrono { get; set; }
        [DataMember]
        public string FechaInicioRealCrono { get; set; }
        [DataMember]
        public string FechaFinRealCrono { get; set; }
        [DataMember]
        public string PPGIDCrono { get; set; }
        [DataMember]
        public string PPGID2Crono { get; set; }
        [DataMember]
        public string CorreoCrono { get; set; }
        [DataMember]
        public string CorreoCrono2 { get; set; }
        [DataMember]
        public string NombreResponsable { get; set; }
        [DataMember]
        public string NombreResponsable2 { get; set; }
        [DataMember]
        public string TipoSubCanal { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{

    [DataContract]
    public class Rentabilidad
    {
        //ENTRADA
        [DataMember]
        public int IdCampania { get; set; }
        [DataMember]
        public string ClaveCampania { get; set; }
        [DataMember]
        public string PeriodoAnt { get; set; }
        [DataMember]
        public string PeriodoAct { get; set; }
        [DataMember]
        public int EsGuardar { get; set; }
        [DataMember]
        public string PPGID { get; set; }
        [DataMember]
        public string Usuario { get; set; }
        [DataMember]
        public string Estatus { get; set; }
        [DataMember]
        public List<GastoPlanta> ListGastoPlanta { get; set; }
        [DataMember]
        public List<SemaforoRentabilidad> ListSemaforoRentabilidad { get; set; }
        [DataMember]
        public List<ReporteCEO> ListReporteCEO { get; set; }
        [DataMember]
        public List<ReporteCEOMecanica> ListReporteCEOMecanica { get; set; }
        [DataMember]
        public List<ReporteCEOPublicidad> ListReporteCEOPublicidad { get; set; }
        [DataMember]
        public List<ReporteMKT> ListReporteMKT { get; set; }
        [DataMember]
        public List<ReporteSKU> ListReporteSKU { get; set; }
        [DataMember]
        public List<Campana> ListCampania { get; set; }
        [DataMember]
        public List<SKUPrecioCosto> ListSKUPrecioCosto { get; set; }
        [DataMember]
        public string SKU { get; set; }

        //SALIDA
        [DataMember]
        public decimal PorcentajeCompania { get; set; }
        [DataMember]
        public decimal PorcentajeNecesario { get; set; }
        [DataMember]
        public string Comentario { get; set; }

        //REPORTES
        [DataMember]
        public string UrlReporteCEOPDF { get; set; }
        [DataMember]
        public string UrlReporteCEOExcel { get; set; }
        [DataMember]
        public string UrlReporteMKTPDF { get; set; }
        [DataMember]
        public string UrlReporteMKTExcel { get; set; }
        [DataMember]
        public string UrlReporteSKUPDF { get; set; }
        [DataMember]
        public string UrlReporteSKUExcel { get; set; }

        //ERROR
        [DataMember]
        public string Mensaje { get; set; }
    }
}

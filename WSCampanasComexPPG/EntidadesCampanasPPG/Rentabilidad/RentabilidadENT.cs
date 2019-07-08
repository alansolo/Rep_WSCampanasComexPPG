using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EntidadesCampanasPPG.Modelo;

namespace EntidadesCampanasPPG.Modelo
{
    public class RentabilidadENT
    {
        public Rentabilidad Rentabilidad { get; set; }
        public string UrlReporteCEOPDF { get; set; }
        public string UrlReporteCEOExcel { get; set; }
        public string UrlReporteMKTPDF { get; set; }
        public string UrlReporteMKTExcel { get; set; }
        public string UrlReporteSKUPDF { get; set; }
        public string UrlReporteSKUExcel { get; set; }
        public string Mensaje { get; set; }
            
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class ReporteCEO
    {
        public string ClaveCampania { get; set; }
        public string Periodo { get; set; }
        public decimal SubtotalLitros { get; set; }
        public decimal SubtotalPiezas { get; set; }
        public decimal TotalLitrosPiezas { get; set; }
        public decimal Importe { get; set; }
        public decimal ImporteLitros { get; set; }
        public decimal ImportePiezas { get; set; }
        public decimal PrecioPromedioLitro { get; set; }
        public decimal CostoMP { get; set; }
        public decimal UtilidadMP{ get; set; }
        public decimal FactorUtilidad { get; set; }
        public decimal MP { get; set; }
        public decimal MargenUtilidad { get; set; }
        public decimal InversionPublicidad { get; set; }
        public decimal OtrosGastos { get; set; }
        public decimal GastosOperacion { get; set; }
        public decimal NotasCredito { get; set; }
        public decimal UtilidadConsideraMP { get; set; }
        public decimal UtilidadLitroPieza { get; set; }
        public decimal FactorUtilidadMP { get; set; }
        public decimal PorcenUtilidadConsideraMP { get; set; }
        public decimal PorcenIncrementoLitros { get; set; }
        public decimal PorcenIncrementoLitrosPresu { get; set; }
        public decimal Roi { get; set; }
        public decimal ImportePrecioPublico { get; set; }
        public decimal ImportePrecioPublicoSinIva { get; set; }
        public decimal UtilidadEnConcesionario { get; set; }
        public decimal MargenConcesionario { get; set; }
    }
}

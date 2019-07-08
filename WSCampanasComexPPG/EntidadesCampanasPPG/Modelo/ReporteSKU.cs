using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class ReporteSKU
    {
        public string ClaveCampania { get; set; }
        public string Articulo { get; set; }
        public string Descripcion { get; set; }
        public string Capacidad { get; set; }
        public decimal CostoMPAnterior { get; set; }
        public decimal GastoPlantaAnterior { get; set; }
        public decimal GastoKromaAñoAnterior { get; set; }
        public decimal CostoMPEnvConvAñoActual { get; set; }
        public decimal CostoPlantaAñoActual { get; set; }
        public decimal GastoKromaActual { get; set; }
        public decimal PrecioConcAnterior { get; set; }
        public decimal PrecioAntConPromocion { get; set; }
        public decimal InversionTintaAnt { get; set; }
        public decimal PrecioPublicoAntCP { get; set; }
        public decimal PrecioConc { get; set; }
        public decimal FactorUtilidadBruto { get; set; }
        public decimal PrecioPublicoSC { get; set; }
        public decimal InversionTintaActSC { get; set; }
        public decimal PorcMargenConcSC { get; set; }
        public decimal PrecioConce { get; set; }
        public decimal FactorUtilidadCC { get; set; }
        public decimal PrecioPublicoCC { get; set; }
        public decimal InversionTintaActCC { get; set; }
        public decimal PorcMargenConcCC { get; set; }
        public decimal LitrosAnterior { get; set; }
        public decimal PiezasVacioAnt { get; set; }
        public decimal PiezasAnt { get; set; }
        public decimal CostoMPEnvFab { get; set; }
        public decimal GastoPlantaAnt { get; set; }
        public decimal GastoKromaAnter { get; set; }
        public decimal UtilidadBrutaComexAnt { get; set; }
        public decimal ImporteCPrecioConcLitro { get; set; }
        public decimal ImporteCPrecioConcPiezas { get; set; }
        public decimal UtilidadConcAnterior { get; set; }
        public decimal InversionTintaAnterior { get; set; }
        public decimal ImporteCPrecioPublicoAnt { get; set; }
        public decimal LitrosActualSC { get; set; }
        public decimal LitrosVacioSC { get; set; }
        public decimal PiezasActualSC { get; set; }
        public decimal CostoMPEnvFabrica { get; set; }
        public decimal GastoPlantaActualSC { get; set; }
        public decimal GastoKromaActualSC { get; set; }
        public decimal UtilidadBrutaComexSC { get; set; }
        public decimal Importe { get; set; }
        public decimal ImporteCPrecioConcPiezasSC { get; set; }
        public decimal UtilidadConcSC { get; set; }
        public decimal InversionTintaSC { get; set; }
        public decimal ImporteCPrecioPublicoSC { get; set; }
        public decimal LitrosActualCC { get; set; }
        public decimal LitrosVacioCC { get; set; }
        public decimal PiezasActualCC { get; set; }
        public decimal CostoMPEnvFabCC { get; set; }
        public decimal GastoPlantaActualCC { get; set; }
        public decimal GastoKromaActualCC { get; set; }
        public decimal UtilidadBrutaComexCC { get; set; }
        public decimal ImporteCPrecioConcLitros { get; set; }
        public decimal ImporteCPrecioConcCCPiezas { get; set; }
        public decimal UtilidadConcCC { get; set; }
        public decimal InversionTintaCC { get; set; }
        public decimal ImporteCPrecioPublicoCC { get; set; }
        public decimal Anterior { get; set; }
        public decimal SinCampania { get; set; }
        public decimal ConPromocion { get; set; }
        public decimal Rentabilidad { get; set; }
        public string Comentario { get; set; }
        public decimal LitrosNecesarios { get; set; }
        public decimal PiezasNecesario { get; set; }
        public decimal Piezas { get; set; }
        public decimal CostoMPEnvNecesario { get; set; }
        public decimal CostoPlantaNecesario { get; set; }
        public decimal GastoKromaNecesario { get; set; }
        public decimal UtilidadBruta { get; set; }
        public decimal ImporteCPrecioConcLitrosNecesario { get; set; }
        public decimal ImporteCPrecioConcPiezasNecesario { get; set; }
        public decimal UtilidadConc { get; set; }
        public decimal TintaNecesario { get; set; }
        public decimal ImporteCPrecioPublicoNecesario { get; set; }
        public string SubArticulo { get; set; }
        public int Agrupador { get; set; }
        public int SubAgrupador { get; set; }
        public string Alcance { get; set; }
        public string SubAlcance { get; set; }
        public string SubDescripcion { get; set; }
        public int Orden { get; set; }
    }
}

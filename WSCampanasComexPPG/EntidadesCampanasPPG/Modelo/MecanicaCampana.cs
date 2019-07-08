using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntidadesCampanasPPG.Modelo
{
    public class MecanicaCampana
    {
        List<MecanicaRegalo> ListMecanicaRegalo { get; set; }
        List<MecanicaMultiplo> ListMecanicaMultiplo { get; set; }
        List<MecanicaDescuento> ListMecanicaDescuento { get; set; }
        List<MecanicaVolumen> ListMecanicaVolumen { get; set; }
        List<MecanicaKit> ListMecanicaKit { get; set; }
        List<MecanicaCombo> ListMecanicaCombo { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace EntidadesCampanasPPG.Modelo
{

    [DataContract]
    public class Campana
    {
        [DataMember]
        public int IdCampana
        { get; set; }

        [DataMember]
        public string Title
        { get; set; }

        [DataMember]
        public string FechaDocumento
        { get; set; }

        [DataMember]
        public string NombreCampa
        { get; set; }

        [DataMember]
        public int IdNegocioLider
        { get; set; }
        [DataMember]
        public string NegocioLider
        { get; set; }

        [DataMember]
        public string LiderCampa
        { get; set; }

        [DataMember]
        public string PPGIDLiderCampa
        { get; set; }

        [DataMember]
        public int IdSubcanal
        { get; set; }
        [DataMember]
        public string Subcanal
        { get; set; }

        [DataMember]
        public string FechaInicioSubCanal
        { get; set; }

        [DataMember]
        public string FechaFinSubCanal
        { get; set; }

        [DataMember]
        public string FechaInicioPublico
        { get; set; }

        [DataMember]
        public string FechaFinPublico
        { get; set; }

        [DataMember]
        public bool CampaExpress
        { get; set; }

        [DataMember]
        public string ObjetivoNC
        { get; set; }

        [DataMember]
        public string JustificacionNC
        { get; set; }

        [DataMember]
        public int IdMoneda
        { get; set; }

        [DataMember]
        public string Moneda
        { get; set; }

        [DataMember]
        public int IdTipoCampa
        { get; set; }
        [DataMember]
        public string TipoCampa
        { get; set; }

        [DataMember]
        public int IdTipoSell
        { get; set; }
        [DataMember]
        public string TipoSell
        { get; set; }
        [DataMember]
        public int IdAlcanceTerritorial
        { get; set; }
        [DataMember]
        public string AlcanceTerritorial
        { get; set; }
        //[DataMember]
        //public string ClientesOtrosCanales
        //{ get; set; }

        [DataMember]
        public string Status
        { get; set; }

        //[DataMember]
        //public string ExcMecanicas
        //{ get; set; }

        [DataMember]
        public string ObsPublicidad
        { get; set; }

        [DataMember]
        public string RegistraCampa
        { get; set; }

        [DataMember]
        public string PPGIDRegistraCampa
        { get; set; }

        [DataMember]
        public List<Kroma> ListaKroma
        { get; set; }

        [DataMember]
        public List<Publicidad> ListaPublicidad
        { get; set; }

        //[DataMember]
        //public List<Producto> ListaProducto
        //{ get; set; }

        [DataMember]
        public string UrlArchivoCronograma
        { get; set; }

        [DataMember]
        public string UrlArchivoProductos
        { get; set; }

        [DataMember]
        public List<Cronograma> ListaCronograma
        { get; set; }

        [DataMember]
        public int VersionCronograma
        { get; set; }

        [DataMember]
        public List<Region> ListaRegion
        { get; set; }

        [DataMember]
        public List<SubCanal> ListaSubCanal
        { get; set; }

        [DataMember]
        public int Excepcion
        { get; set; }

        [DataMember]
        public List<Bitacora> ListBitacora
        { get; set; }
        [DataMember]
        public List<MecanicaRegalo> ListMecanicaRegalo
        { get; set; }
        [DataMember]
        public List<MecanicaMultiplo> ListMecanicaMultiplo
        { get; set; }
        [DataMember]
        public List<MecanicaDescuento> ListMecanicaDescuento
        { get; set; }
        [DataMember]
        public List<MecanicaVolumen> ListMecanicaVolumen
        { get; set; }
        [DataMember]
        public List<MecanicaKit> ListMecanicaKit
        { get; set; }
        [DataMember]
        public List<MecanicaCombo> ListMecanicaCombo
        { get; set; }
        [DataMember]
        public List<Tienda> ListTienda
        { get; set; }
        [DataMember]
        public List<Rentabilidad> ListRentabilidad
        { get; set; }
        [DataMember]
        public string Estatus
        { get; set; }
        [DataMember]
        public int IdEstatus
        { get; set; }
        [DataMember]
        public decimal ImporteNotaCredito
        { get; set; }
        [DataMember]
        public string ComentarioVenta
        { get; set; }
        [DataMember]
        public bool CampaniaAnterior
        { get; set; }
        [DataMember]
        public List<SignoVital> ListaSignoVital
        { get; set; }

        [DataMember]
        public string TitleAnterior
        { get; set; }
        [DataMember]
        public string NombreCampaAnterior
        { get; set; }
        [DataMember]
        public string ObjetivoKroma
        { get; set; }
        [DataMember]
        public string JustificacionKroma
        { get; set; }
        [DataMember]
        public string TipoSubCanal
        { get; set; }
    }

}

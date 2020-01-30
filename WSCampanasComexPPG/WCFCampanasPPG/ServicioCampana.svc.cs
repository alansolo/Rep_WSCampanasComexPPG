using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using EntidadesCampanasPPG.Bitacora;
using EntidadesCampanasPPG.Campana;
using EntidadesCampanasPPG.FlujoActividad;
using EntidadesCampanasPPG.GastoPlanta;
using EntidadesCampanasPPG.Modelo;
using EntidadesCampanasPPG.Producto;
using EntidadesCampanasPPG.Reporte;
using NegocioCampanasPPG.Bitacora;
using NegocioCampanasPPG.Campana;

namespace WCFCampanasPPG
{
    // NOTA: puede usar el comando "Rename" del menú "Refactorizar" para cambiar el nombre de clase "Service1" en el código, en svc y en el archivo de configuración.
    // NOTE: para iniciar el Cliente de prueba WCF para probar este servicio, seleccione Service1.svc o Service1.svc.cs en el Explorador de soluciones e inicie la depuración.
    public class ServicioCampana : IServicioCampana
    {
        //PRIMER PASO PARA GUARDAR INFORMACION DE CAMPAÑA Y CRONOGRAMA
        public string GuardarCampanaCronograma(Campana campana)
        {
            string resultado = string.Empty;

            CampanaENT campanaENTReq = new CampanaENT();
            campanaENTReq.Campana = campana;

            CampanaENT campanaENTRes = new CampanaENT();

            CampanaNEG campanaNEG = new CampanaNEG();

            campanaENTRes = campanaNEG.GuardarCampanaCronograma(campanaENTReq);

            resultado = campanaENTRes.Mensaje;


            return resultado;
        }
        //SE GUARDA O ACTUALIZA SOLO CRONOGRAMA
        public string GuardarCronograma(Campana campana)
        {
            string resultado = string.Empty;

            CampanaENT campanaENTReq = new CampanaENT();
            campanaENTReq.Campana = campana;

            CampanaENT campanaENTRes = new CampanaENT();

            CampanaNEG campanaNEG = new CampanaNEG();

            campanaENTRes = campanaNEG.GuardarCampanaCronograma(campanaENTReq);

            resultado = campanaENTRes.Mensaje;


            return resultado;
        }
        //SEGUNDO PASO PARA GUARDAR INFORMACION DE CAMPAÑA Y ESCENARIOS
        public ProductoCampana GuardarProductoCampana(ArchivoProducto ArchivoProducto)
        {
            ProductoCampana productoCampana = new ProductoCampana();
            List<ProductoLinea> ListProductoLinea = new List<ProductoLinea>();

            ProductoLineaENT productoLineaENTReq = new ProductoLineaENT();
            productoLineaENTReq.ClaveCampana = ArchivoProducto.ClaveCampana;
            productoLineaENTReq.UrlArchivo = ArchivoProducto.UrlArchivo;

            ProductoLineaENT productoLineaENTRes = new ProductoLineaENT();

            CampanaNEG campanaNEG = new CampanaNEG();

            productoLineaENTRes = campanaNEG.GuardarProductoCampana(productoLineaENTReq);

            if (productoLineaENTRes.Mensaje == "OK")
            {
                productoCampana.ListLineaFamilia = productoLineaENTRes.ListLineaFamilia;
                productoCampana.ListMensaje = productoLineaENTRes.ListMensaje;
                productoCampana.ListSKUValidacion = new List<SKUValidacion>();
                productoCampana.Mensaje = productoLineaENTRes.Mensaje;
            }
            else
            {
                productoCampana.ListMensaje = productoLineaENTRes.ListMensaje;
                productoCampana.ListSKUValidacion = productoLineaENTRes.ListSkuValidacion;
                productoCampana.Mensaje = productoLineaENTRes.Mensaje;
            }

            return productoCampana;
        }
        //TERCER PASO PARA GUARDAR INFORMACION DE CAMPAÑA
        public string GuardarCampana(Campana Campana)
        {
            string resultado = string.Empty;

            CampanaENT campanaENTReq = new CampanaENT();
            campanaENTReq.Campana = Campana;

            CampanaENT campanaENTRes = new CampanaENT();

            CampanaNEG campanaNEG = new CampanaNEG();

            campanaENTRes = campanaNEG.GuardarCampana(campanaENTReq);

            if (campanaENTRes.Mensaje == "OK")
            {
                resultado = campanaENTRes.Mensaje;
            }
            else
            {
                resultado = campanaENTRes.Mensaje;
            }

            return resultado;
        }
        //MUESTRA INFORMACION DE CAMPAÑAS AÑO ANTERIOR
        public List<Campana> MostrarCampanaAnioAnterior(Campana Campana)
        {
            List<Campana> ListCampana = new List<Campana>();

            CampanaNEG campanaNEG = new CampanaNEG();

            CampanaBitacoraENT campanaENTReq = new CampanaBitacoraENT();
            campanaENTReq.ListCampana = new List<Campana>();
            campanaENTReq.ListCampana.Add(Campana);

            CampanaBitacoraENT campanaENTRes = new CampanaBitacoraENT();

            campanaENTRes = campanaNEG.MostrarCampanaDetAnioAnterior(campanaENTReq);

            if (campanaENTRes.ListCampana != null)
            {
                ListCampana = campanaENTRes.ListCampana;
            }

            return ListCampana;
        }
        //MUESTRA INFORMACION DE CAMPAÑAS AÑO ANTERIOR
        public List<Campana> MostrarCampanaDetAnioAnterior(Campana Campana)
        {
            List<Campana> ListCampana = new List<Campana>();

            CampanaNEG campanaNEG = new CampanaNEG();

            CampanaBitacoraENT campanaENTReq = new CampanaBitacoraENT();
            campanaENTReq.ListCampana = new List<Campana>();
            campanaENTReq.ListCampana.Add(Campana);

            CampanaBitacoraENT campanaENTRes = new CampanaBitacoraENT();

            campanaENTRes = campanaNEG.MostrarCampanaDetAnioAnterior(campanaENTReq);

            if (campanaENTRes.ListCampana != null)
            {
                ListCampana = campanaENTRes.ListCampana;
            }

            return ListCampana;
        }
        //MUESTRA INFORMACION DE LA CAMPAÑA
        public List<Campana> MostrarCampana(Campana Campana)
        {
            List<Campana> ListCampana = new List<Campana>();

            CampanaNEG campanaNEG = new CampanaNEG();

            CampanaBitacoraENT campanaENTReq = new CampanaBitacoraENT();
            campanaENTReq.ListCampana = new List<Campana>();
            campanaENTReq.ListCampana.Add(Campana);

            CampanaBitacoraENT campanaENTRes = new CampanaBitacoraENT();

            campanaENTRes = campanaNEG.MostrarCampana(campanaENTReq);

            if (campanaENTRes.ListCampana != null)
            {
                ListCampana = campanaENTRes.ListCampana;
            }

            return ListCampana;
        }
        //MUESTRA INFORMACION DE LA CAMPAÑA Y LOS ESCENARIOS
        public ProductoCampana MostrarProductoCampana()
        {
            ProductoCampana productoCampana = new ProductoCampana();

            productoCampana.ListLineaFamilia = new List<LineaFamilia>();

            LineaFamilia lineaFamilia = new LineaFamilia();

            productoCampana.ListLineaFamilia.Add(lineaFamilia);

            productoCampana.ListLineaFamilia.Add(lineaFamilia);

            return productoCampana;
        }
        //MUESTRA INFORMACION DE LA CAMPAÑA Y EL CRONOGRAMA
        public MostrarCronograma MostrarCampanaCronograma(Campana campana)
        {
            MostrarCronograma mostrarCronograma = new MostrarCronograma();

            CampanaENT campanaENTReq = new CampanaENT();
            campanaENTReq.Campana = campana;

            CampanaENT campanaENTRes = new CampanaENT();

            CampanaNEG campanaNEG = new CampanaNEG();

            campanaENTRes = campanaNEG.MostrarCampanaCronograma(campanaENTReq);

            mostrarCronograma = campanaENTRes.MostrarCronograma;


            return mostrarCronograma;
        }
        //REALIZA CALCULO DE RENTABILIDAD Y GENERA REPORTES
        public Rentabilidad MostrarRentabilidad(Rentabilidad Rentabilidad)
        {
            RentabilidadENT rentabilidadENTReq = new RentabilidadENT();
            rentabilidadENTReq.Rentabilidad = Rentabilidad;

            RentabilidadENT rentabilidadENTRes = new RentabilidadENT();

            CampanaNEG campanaNEG = new CampanaNEG();

            rentabilidadENTRes = campanaNEG.MostrarRentabilidad(rentabilidadENTReq);

            if (rentabilidadENTRes.Mensaje == "OK")
            {
                Rentabilidad = rentabilidadENTRes.Rentabilidad;
                Rentabilidad.Mensaje = rentabilidadENTRes.Mensaje;

                Rentabilidad.UrlReporteCEOPDF = rentabilidadENTRes.UrlReporteCEOPDF;
                Rentabilidad.UrlReporteCEOExcel = rentabilidadENTRes.UrlReporteCEOExcel;
                Rentabilidad.UrlReporteMKTPDF = rentabilidadENTRes.UrlReporteMKTPDF;
                Rentabilidad.UrlReporteMKTExcel = rentabilidadENTRes.UrlReporteMKTExcel;
                Rentabilidad.UrlReporteSKUPDF = rentabilidadENTRes.UrlReporteSKUPDF;
                Rentabilidad.UrlReporteSKUExcel = rentabilidadENTRes.UrlReporteSKUExcel;
            }
            else
            {
                Rentabilidad = new Rentabilidad();
                Rentabilidad.Mensaje = rentabilidadENTRes.Mensaje;

                Rentabilidad.UrlReporteCEOPDF = string.Empty;
                Rentabilidad.UrlReporteCEOExcel = string.Empty;
                Rentabilidad.UrlReporteMKTPDF = string.Empty;
                Rentabilidad.UrlReporteMKTExcel = string.Empty;
                Rentabilidad.UrlReporteSKUPDF = string.Empty;
                Rentabilidad.UrlReporteSKUExcel = string.Empty;
            }

            return Rentabilidad;
        }
        //REALIZA CALCULO DE RENTABILIDAD Y GENERA REPORTES 
        public Rentabilidad MostrarRentabilidadEjemplo()
        {
            Rentabilidad rentabilidad = new Rentabilidad();

            rentabilidad.ListSemaforoRentabilidad = new List<SemaforoRentabilidad>();
            rentabilidad.ListSemaforoRentabilidad.Add(new SemaforoRentabilidad());

            rentabilidad.ListGastoPlanta = new List<GastoPlanta>();
            rentabilidad.ListGastoPlanta.Add(new GastoPlanta());

            return rentabilidad;
        }
        //MUESTRA REPORTE DIRECTIVO
        public ReporteDirectivo MostrarReporteDirectivo(Campana Campana)
        {
            ReporteDirectivo reporteDirectivo = new ReporteDirectivo();

            ReporteENT reporteENTReq = new ReporteENT();
            reporteENTReq.ClaveCampana = Campana.Title;

            ReporteENT reporteENTRes = new ReporteENT();

            CampanaNEG campanaNEG = new CampanaNEG();

            reporteENTRes = campanaNEG.GenerarReporteDirectivo(reporteENTReq);

            if(reporteENTRes.Mensaje == "OK")
            {
                reporteDirectivo.Mensaje = reporteENTRes.Mensaje;
                reporteDirectivo.URLReporteDirectivo = reporteENTRes.UrlArchivo;
            }
            else
            {
                reporteDirectivo.Mensaje = reporteENTRes.Mensaje;
            }

            return reporteDirectivo;
        }
        //GUARDAR INFORMACION EN BITACORA
        public string GuardarBitacora(Bitacora Bitacora)
        {
            string resultado = string.Empty;

            BitacoraNEG bitacoraNEG = new BitacoraNEG();
            BitacoraENT bitacoraENTReq = new BitacoraENT();
            bitacoraENTReq.ListBitacora = new List<Bitacora>();
            bitacoraENTReq.ListBitacora.Add(Bitacora);

            BitacoraENT bitacoraENTRes = new BitacoraENT();

            bitacoraENTRes = bitacoraNEG.GuardarBitacora(bitacoraENTReq);

            if (bitacoraENTRes.Mensaje == "OK")
            {
                resultado = bitacoraENTRes.Mensaje;
            }
            else
            {
                resultado = bitacoraENTRes.Mensaje;
            }

            return resultado;
        }
        //MUESTRA INFORMACION DE BITACORA
        public List<Bitacora> MostrarBitacora(Bitacora Bitacora)
        {
            List<Bitacora> ListBitacora = new List<Bitacora>();

            BitacoraNEG bitacoraNEG = new BitacoraNEG();
            BitacoraENT bitacoraENTReq = new BitacoraENT();
            bitacoraENTReq.ListBitacora = new List<Bitacora>();

            bitacoraENTReq.ListBitacora.Add(Bitacora);

            BitacoraENT bitacoraENTRes = new BitacoraENT();

            bitacoraENTRes = bitacoraNEG.MostrarBitacora(bitacoraENTReq);

            if(bitacoraENTRes.Mensaje =="OK")
            {
                ListBitacora = bitacoraENTRes.ListBitacora;
            }

            return ListBitacora;
        }
        //MUESTRA APROBADOR DEPENDIENTE DE PREDECESOR
        public FlujoActividadAprobador MostrarAprobadorPredecesor(FlujoActividad flujoActividad)
        {
            FlujoActividadAprobador FlujoActividadAprobador = new FlujoActividadAprobador();

            CampanaNEG campanaNEG = new CampanaNEG();
            FlujoActividadENT flujoActividadENTReq = new FlujoActividadENT();
            flujoActividadENTReq.ListFlujoActividad = new List<FlujoActividad>();
            flujoActividadENTReq.ListFlujoActividad.Add(flujoActividad);

            FlujoActividadENT flujoActividadENTRes = new FlujoActividadENT();

            flujoActividadENTRes = campanaNEG.MostrarAprobadorPredecesor(flujoActividadENTReq);

            FlujoActividadAprobador.ListFlujoActividad = flujoActividadENTRes.ListFlujoActividad;
            FlujoActividadAprobador.Mensaje = flujoActividadENTRes.Mensaje;


            return FlujoActividadAprobador;
        }
        //MUESTRA BITACORA DETALLE CAMPANA
        public List<Campana> MostrarBitacoraDetalleCampana(Campana campana)
        {
            List<Campana> ListCampana = new List<Campana>();

            CampanaNEG campanaNEG = new CampanaNEG();

            CampanaBitacoraENT campanaENTReq = new CampanaBitacoraENT();
            campanaENTReq.ListCampana = new List<Campana>();
            campanaENTReq.ListCampana.Add(campana);

            CampanaBitacoraENT campanaENTRes = new CampanaBitacoraENT();

            campanaENTRes =  campanaNEG.MostrarCampanaBitacoraDetalle(campanaENTReq);

            if (campanaENTRes.ListCampana != null)
            {
                ListCampana = campanaENTRes.ListCampana;
            }

            return ListCampana;
        }
        //MUESTRA HISTORICO GASTO PLANTA
        public List<GastoPlantaHistorico> MostrarHistoricoGastoPlanta(GastoPlantaHistorico gastoPlanta)
        {
            List<GastoPlantaHistorico> ListGastoPlantaHistorico = new List<GastoPlantaHistorico>();

            CampanaNEG campanaNEG = new CampanaNEG();

            GastoPlantaENT gastoPlantaENTReq = new GastoPlantaENT();
            gastoPlantaENTReq.ListGastoPlantaHistorico = new List<GastoPlantaHistorico>();
            gastoPlantaENTReq.ListGastoPlantaHistorico.Add(gastoPlanta);

            GastoPlantaENT gastoPlantaENTRes = new GastoPlantaENT();

            gastoPlantaENTRes = campanaNEG.MostrarGastoPlantaHistorico(gastoPlantaENTReq);

            if(gastoPlantaENTRes != null && gastoPlantaENTRes.OK == 1)
            {
                ListGastoPlantaHistorico = gastoPlantaENTRes.ListGastoPlantaHistorico;
            }

            return ListGastoPlantaHistorico;
        }
    }
}

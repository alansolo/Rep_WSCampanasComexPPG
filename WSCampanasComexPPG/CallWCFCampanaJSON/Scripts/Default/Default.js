var app = angular.module("MyApp", []);

//funcion inicial para agregar las empresas
app.controller("MyController", function ($scope, $http, $window) {

    //var Campana =
    //{
    //    "CampaExpress": 0,
    //    //"ClientesOtrosCanales": "otros",
    //    //"ExcMecanicas": "1",
    //    //"FechaDocumento": "12/10/2018",
    //    "FechaFinPublico": "01/01/1900",
    //    "FechaFinSubCanal": "01/01/1900",
    //    "FechaInicioPublico": "01/01/1900",
    //    "FechaInicioSubCanal": "01/01/1900",
    //    "IdAlcanceTerritorial": 1,
    //    //"IdMoneda": 1,
    //    //"IdNegocioLider": 1,
    //    //"IdSubcanal": 1,
    //    "IdTipoCampa": 1,
    //    "IdTipoSell": 1,
    //    //"JustificacionNC": "justiNC",
    //    //"Kroma":
    //    //{
    //    //    "DF": 1,
    //    //    "DF17": 1,
    //    //    "DF6": 1,
    //    //    "DF9": 1,
    //    //    "FronteraNorteMonterrey": 1,
    //    //    "IdLineaAportacion": 1,
    //    //    "Justificacion": "justi",
    //    //    "RegionalOccidente": 1
    //    //},
    //    //"LiderCampa": "lider",
    //    //"ListaProducto":
    //    //    [
    //    //        {
    //    //            "Alcance": 1,
    //    //            "CantidadOdescuento": 1,
    //    //            "CapacidadProducto": "1",
    //    //            "IdFamiliaEstelar": 1,
    //    //            "IdLineaProducto": 1,
    //    //            "IdMecanica": 1,
    //    //            "IdTipoProducto": 1,
    //    //            "Observaciones": "obser",
    //    //            "SistemaAplicacion": "1"
    //    //        }
    //    //    ],
    //    //"ListaPublicidad":
    //    //    [
    //    //        {
    //    //            "IdPublicidad": 1,
    //    //            "Monto": 1
    //    //        }
    //    //    ],
    //    //"NombreCampa": "Campana 11",
    //    //"ObjetivoNC": "obje nc",
    //    //"ObsPublicidad": "obser",
    //    "PPGIDLiderCampa": null,
    //    //"PPGIDRegistraCampa": "dos12",
    //    //"RegistraCampa": "rodri",
    //    //"Status": 1,
    //    "Title": ""
    //};

    //var Campana =
    //{
    //    "CampaExpress": 0,
    //    "ClientesOtrosCanales": "otros",
    //    "ExcMecanicas": "1",
    //    //"FechaDocumento": "12/10/2018",
    //    "FechaFinPublico": "01/01/1900",
    //    "FechaFinSubCanal": "01/01/1900",
    //    "FechaInicioPublico": "01/01/1900",
    //    "FechaInicioSubCanal": "01/01/1900",
    //    "IdAlcanceTerritorial": 1,
    //    "IdMoneda": 1,
    //    "IdNegocioLider": 1,
    //    "IdSubcanal": 1,
    //    "IdTipoCampa": 1,
    //    "IdTipoSell": 1,
    //    "JustificacionNC": "justiNC",
    //    "Kroma":
    //    {
    //        "DF": 1,
    //        "DF17": 1,
    //        "DF6": 1,
    //        "DF9": 1,
    //        "FronteraNorteMonterrey": 1,
    //        "IdLineaAportacion": 1,
    //        "Justificacion": "justi",
    //        "RegionalOccidente": 1
    //    },
    //    "LiderCampa": "lider",
    //    "ListaProducto":
    //        [
    //            {
    //                "Alcance": 1,
    //                "CantidadOdescuento": 1,
    //                "CapacidadProducto": "1",
    //                "IdFamiliaEstelar": 1,
    //                "IdLineaProducto": 1,
    //                "IdMecanica": 1,
    //                "IdTipoProducto": 1,
    //                "Observaciones": "obser",
    //                "SistemaAplicacion": "1"
    //            }
    //        ],
    //    "ListaPublicidad":
    //        [
    //            {
    //                "IdPublicidad": 1,
    //                "Monto": 1
    //            }
    //        ],
    //    "NombreCampa": "Campana 11",
    //    "ObjetivoNC": "obje nc",
    //    "ObsPublicidad": "obser",
    //    "PPGIDLiderCampa": "uno12",
    //    "PPGIDRegistraCampa": "dos12",
    //    "RegistraCampa": "rodri",
    //    "Status": 1,
    //    "Title": "Clave 12"
    //};

    //var Campana =
    //{
    //    "Title": "2019-NA-7"
    //};

    var Campana =
    {
        "CampaExpress": true,
        "ClientesOtrosCanales": "eklkm",
        "ExcMecanicas": "1",
        "FechaDocumento": "04/12/2018",
        "FechaFinPublico": null,
        "FechaFinSubCanal": "20/12/2018",
        "FechaInicioPublico": null,
        "FechaInicioSubCanal": "14/12/2018",
        "IdAlcanceTerritorial": 5,
        "IdMoneda": 1,
        "IdNegocioLider": 1,
        "IdSubcanal": 1,
        "IdTipoCampa": 1,
        "IdTipoSell": 1,
        "JustificacionNC": "",
        "ListaKroma":
            [
                {
                    "IdFamilia": 1,
                    "Linea": "L1",
                    "SobrePrecioKroma": "DF1",
                    "Porcetaje": 0.04
                },
                {
                    "IdFamilia": 1,
                    "Linea": "L1",
                    "SobrePrecioKroma": "DF11",
                    "Porcetaje": 0.05
                },
                {
                    "IdFamilia": 1,
                    "Linea": "L1",
                    "SobrePrecioKroma": "DF7",
                    "Porcetaje": 0.03
                }
            ],
        "LiderCampa": "Perez Ortega, Julio Cesar [C]",
        "ListaRegion":
            [
                {
                    "NombreRegion": "Norte",
                    "PPGID": "ppgna\k697344",
                    "Usuario": "Solorzano Tapia, Alan Michel [C]"
                },
                {
                    "NombreRegion": "Sur",
                    "PPGID": "ppgna\k697344",
                    "Usuario": "Solorzano Tapia, Alan Michel [C]"
                }
            ],
        //"ListaProducto":
        //    [
        //        {
        //            "Alcance": 1,
        //            "CantidadOdescuento": 1,
        //            "CapacidadProducto": 3,
        //            "IdFamiliaEstelar": 1,
        //            "IdLineaProducto": 1,
        //            "IdMecanica": 6,
        //            "IdTipoProducto": 1,
        //            "Observaciones": "kowoer",
        //            "SistemaAplicacion": "1ee"
        //        },
        //        {
        //            "Alcance": 2,
        //            "CantidadOdescuento": 2,
        //            "CapacidadProducto": 1,
        //            "IdFamiliaEstelar": 2,
        //            "IdLineaProducto": 2,
        //            "IdMecanica": 5,
        //            "IdTipoProducto": 2,
        //            "Observaciones": "jt54t",
        //            "SistemaAplicacion": "1ww"
        //        }
        //    ],
        "ListaPublicidad":
            [
                { "Monto": 10, "IdPublicidad": 1, "MontoAnterior": 100 },
                { "Monto": 20, "IdPublicidad": 2, "MontoAnterior": 100 },
                { "Monto": 30, "IdPublicidad": 3, "MontoAnterior": 100 },
                { "Monto": 40, "IdPublicidad": 4, "MontoAnterior": 100 },
                { "Monto": 50, "IdPublicidad": 5, "MontoAnterior": 100 },
                { "Monto": 60, "IdPublicidad": 6, "MontoAnterior": 100 }
            ],
        "NombreCampa": "Prueba Region",
        "ObjetivoNC": "",
        "ObsPublicidad": "",
        "PPGIDLiderCampa": "ppgna\k697344",
        "PPGIDRegistraCampa": "ppgna\k697344",
        "RegistraCampa": "Solorzano Tapia, Alan Michel [C]",
        "Status": "Creación de Campaña",
        "Title": "2019-NA-160",
        "UrlArchivoCronograma": "https://one.web.ppg.com/la/comex/camp_publicidad/Documentos%20compartidos/03_Cronograma%20Nacional(1).xlsx"
    };

    //var Campana =
    //{
    //    "Title": null
    //};

    //var Campana =
    //{
    //    "CampaExpress": true,
    //    "CampaniaAnterior": false,
    //    "ComentarioVenta": "",
    //    "ClientesOtrosCanales": "Canal",
    //    "ExcMecanicas": "Mecanica",
    //    "FechaDocumento": "24/04/2019",
    //    "FechaFinPublico": "12/05/2019",
    //    "FechaFinSubCanal": "12/05/2019",
    //    "FechaInicioPublico": "06/05/2019",
    //    "FechaInicioSubCanal": "06/05/2019",
    //    "IdAlcanceTerritorial": 1,
    //    "IdMoneda": 1,
    //    "IdNegocioLider": 1,
    //    "IdSubcanal": 1,
    //    "IdTipoCampa": 1,
    //    "IdTipoSell": 3,
    //    "ImporteNotaCredito": "0",
    //    "JustificacionKroma": "",
    //    "JustificacionNC": "",
    //    "IdPublicidad": 6,
    //    "Monto": 0,
    //    "MontoAnterior": 0,
    //    "NombreCampa": "Campana",
    //    "NombreCampaAnterior": "",
    //    "ObjetivoNC": "",
    //    "LiderCampa": "Perez Ortega, Julio Cesar [C]",      
    //    "ObsPublicidad": "",
    //    "PPGIDLiderCampa": "M698854",
    //    "PPGIDRegistraCampa": "M698854",
    //    "RegistraCampa": "Rodrigo Campero",
    //    "Status": 1,
    //    "Title": "2019-NA-160",
    //    "UrlArchivoCronograma": "https://one.web.ppg.com/la/comex/camp_publicidad/Documentos%20compartidos/03_Cronograma%20Nacional(1).xlsx",
    //    "UrlArchivoProductos": "https://one.web.ppg.com/la/comex/camp_publicidad/DocumentosCampaas/Escenarios/ESC-2019-NA-129.xlsm",
    //    "ListaProducto":
    //        [
    //            {
    //                "Alcance": 1,
    //                "CantidadOdescuento": 10,
    //                "CapacidadProducto": "1",
    //                "IdFamiliaEstelar": 1,
    //                "IdLineaProducto": 1,
    //                "IdMecanica": 1,
    //                "IdTipoProducto": 1,
    //                "Observaciones": "Unico",
    //                "SistemaAplicacion": 1
    //            }
    //        ],
    //    "ListaPublicidad":
    //        [
    //            {
    //                "IdPublicidad": 1,
    //                "Monto": 0,
    //                "MontoAnterior": 0
    //            },
    //            {
    //                "IdPublicidad": 2,
    //                "Monto": 0,
    //                "MontoAnterior": 0
    //            },
    //            {
    //                "IdPublicidad": 3,
    //                "Monto": 0,
    //                "MontoAnterior": 0
    //            },
    //            {
    //                "IdPublicidad": 4,
    //                "Monto": 0,
    //                "MontoAnterior": 0
    //            },
    //            {
    //                "IdPublicidad": 5,
    //                "Monto": 0,
    //                "MontoAnterior": 0
    //            },
    //            {
    //                "IdPublicidad": 6,
    //                "Monto": 0,
    //                "MontoAnterior": 0
    //            }
    //        ],
    //    "ListaSignoVital":
    //        [
    //            "Artículos por Ticket"
    //        ]
    //        };

    //$http
    //    ({
    //        method: "POST",
    //        url: "https://cmxcampaniadev.nac.ppg.com/ServicioCampana.svc/GuardarCampana",
    //        //url: "http://localhost:3640/ServicioCampana.svc/GuardarCampana",
    //        data: JSON.stringify(Campana),
    //        dataType: "json",
    //        contentType: "json",
    //        header: { "contentType": "application/json" }
    //    }).then(function (dat) {
    //        var uno = dat.data;
    //    },
    //        function (er) {
    //            var dos = er.data;
    //        });

    //$http
    //({
    //    method: "POST",
    //    //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarCampana",
    //        url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/GuardarCampana",
    //        //url: "http://localhost:3640/ServicioCampana.svc/GuardarCampana",
    //    data: JSON.stringify(Campana),
    //    dataType: "json",
    //    contentType: "json",
    //    header: { "contentType": "application/json" }
    //}).then(function (dat) {
    //    var uno = dat.data;
    //},
    //    function (er) {
    //        var dos = er.data;
    //    });


    //$.ajax({
    //    type: "GET",
    //    //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/GuardarCampana",
    //    //url: "http://localhost:3640/ServicioCampana.svc/GuardarCampana",
    //    //url: "http://localhost:3640/ServicioCampana.svc/GuardarProductoCampana",
    //    url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarProductoCampana",
    //    //data: JSON.stringify(ArchivoProducto),
    //    //data: JSON.stringify(Campana),
    //    contentType: 'application/json; charset=utf-8',
    //    dataType: 'json',
    //    success: function (data) {
    //        alert("success..." + data);
    //    },
    //    error: function (xhr) {
    //        alert(xhr.responseText);
    //    }
    //});



    var Rentabilidad =
    {
        "Comentario": null,
        "IdCampania":0,
        "ClaveCampania": "2019-NA-3",
        "ListGastoPlanta": [
            { "AnioActual": 20.01, "AnioAnterior": 18, "Planta": "Tepexpan" },
            { "AnioActual": 15, "AnioAnterior": 13, "Planta": "Aga" },
            { "AnioActual": 27, "AnioAnterior": 25, "Planta": "FPU" },
            { "AnioActual": 0, "AnioAnterior": 0, "Planta": "Comprado" },
            { "AnioActual": 17, "AnioAnterior": 16, "Planta": "Gasto Kroma" }
        ],
        "ListSemaforoRentabilidad": [
            { "De": 1, "Descripcion": "Amarillo", "Hasta": 3 },
            { "De": 3.1, "Descripcion": "Rojo", "Hasta": 100 },
            { "De": 0, "Descripcion": "Verde", "Hasta": 0.9 }
        ],
        "PeriodoAct": "201901",
        "PeriodoAnt": "201801",
        "PorcentajeCompania": 0,
        "PorcentajeNecesario": 0,
        "EsGuardar": 0,
        "Estatus": "Analizar Campaña",
        "PPGID": "ppgna\u171461",
        "Usuario": "Perez Ortega, Julio Cesar [C]",
        "Mensaje": null
    };


    //var Rentabilidad =
    //{
    //    "Comentario": "",
    //    "IdCampania": 66,
    //    "ListGastoPlanta":
    //        [
    //            { "AnioActual": 18, "AnioAnterior": 20, "Planta": "Tepexpan" },
    //            { "AnioActual": 13, "AnioAnterior": 15, "Planta": "Aga" },
    //            { "AnioActual": 25, "AnioAnterior": 27, "Planta": "FPU" },
    //            { "AnioActual": 0, "AnioAnterior": 0, "Planta": "Comprado" },
    //            { "AnioActual": 16, "AnioAnterior": 17, "Planta": "Kroma" }
    //        ],
    //    "ListSemaforoRentabilidad":
    //        [
    //            { "De": 10, "Descripcion": "Verde", "Hasta": 0},
    //            { "De": 10, "Descripcion": "Amarillo", "Hasta": 40},
    //            { "De": 40, "Descripcion": "Rojo", "Hasta": 100}
    //        ],
    //    "Mensaje": null,
    //    "PeriodoAct": "201801",
    //    "PeriodoAnt": "201701",
    //    "PorcentajeCompania": 0,
    //    "PorcentajeNecesario": 0
    //};


    var ClaveCampana = "2019-RE-11";
    var UrlArchivo = "https://one.web.ppg.com/la/comex/camp_publicidad/Documentos%20compartidos/Prueba%202019-NA-300%20E1.xlsm";
    /*"https://one.web.ppg.com/la/comex/camp_publicidad/Documentos%20compartidos/Layout_Mecanica_EA2.xlsx";*/

    var ArchivoProducto =
    {
        "ClaveCampana": ClaveCampana,
        "UrlArchivo": UrlArchivo
    };

    var Bitacora =
    {
        "ClaveCampania": "2019-NA-195",
        "Estatus": "lista para parametrizar",
        "PPGID": "ppgna\u171461",
        "Usuario": "Perez Ortega, Julio Cesar [C]",
        "Comentario": "",
        "Completado": 1,
        "IdTarea": 14,
        "TipoFlujo": 3,
        "CorreoResponsable": ""
    };

    //var Bitacora =
    //{
    //    "IDCampania": 93
    //};

    var FlujoActividad =
    {
        "IdCampania": "164",
        "IDTarea": 14
    };

    var GastoPlantaHistorico =
    {
        "Periodo": "201901"
    };

    $.ajax({
        type: "POST",
        //url: "http://localhost:3639/ServicioCampana.svc/GuardarCampana",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/GuardarCampana",
        url: "http://localhost:3639/ServicioCampana.svc/GuardarProductoCampana",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/GuardarProductoCampana",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarCampanaCronograma",
        //url: "https://cmxcampaniatest.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarCampanaCronograma",
        //url: "http://localhost:3639/ServicioCampana.svc/GuardarCampanaCronograma",
        //url: "https://cmxcampaniatest.nac.ppg.com/WSCampania/ServicioCampana.svc/GuardarCampanaCronograma",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarRentabilidad",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarRentabilidad",
        //url: "https://cmxcampaniatest.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarRentabilidad",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarReporteDirectivo",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarReporteDirectivo",
        //url: "http://localhost:3639/ServicioCampana.svc/GuardarBitacora",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/GuardarBitacora",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarCampanaCronograma",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarCampanaCronograma",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarAprobadorPredecesor",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarAprobadorPredecesor",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarBitacora",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarBitacora",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarCampana",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarCampanaDetAnioAnterior",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarBitacoraDetalleCampana",
        //url: "http://localhost:3639/ServicioCampana.svc/MostrarHistoricoGastoPlanta",
        //url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/MostrarHistoricoGastoPlanta",
        data: JSON.stringify(ArchivoProducto),
        //data: JSON.stringify(Campana),
        //data: JSON.stringify(Rentabilidad),
        //data: JSON.stringify(Bitacora),
        //data: JSON.stringify(FlujoActividad),
        //data: JSON.stringify(GastoPlantaHistorico),
        contentType: 'application/json; charset=utf-8',
        dataType: 'json',
        success: function (data) {

            alert("success..." + data);
        },
        error: function (xhr) {
            alert(xhr.responseText);
        }
    });

    //$http
    //    ({
    //        method: "POST",
    //        url: "https://cmxcampaniadev.nac.ppg.com/WSCampania/ServicioCampana.svc/GuardarProductoCampana",
    //        data: JSON.stringify(ArchivoProducto),
    //        dataType: "json",
    //        contentType: "json",
    //        header: { "contentType": "application/json" }
    //    }).then(function (dat) {
    //        var uno = dat.data;
    //    },
    //    function (er) {
    //        var dos = er.data;
    //    });
});
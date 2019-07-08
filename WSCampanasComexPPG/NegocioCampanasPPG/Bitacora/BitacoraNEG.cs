using DatosCampanasPPG.BitacoraDAT;
using DatosCampanasPPG.Campana;
using DatosCampanasPPG.Catalogo;
using EntidadesCampanasPPG.Bitacora;
using EntidadesCampanasPPG.Modelo;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using UtilidadesCampanasPPG;

namespace NegocioCampanasPPG.Bitacora
{
    public class BitacoraNEG
    {
        public BitacoraENT GuardarBitacora(BitacoraENT bitacoraENTReq)
        {
            int respuesta = 0;
            BitacoraENT bitacoraENTRes = new BitacoraENT();
            BitacoraDAT bitacoraDAT = new BitacoraDAT();
            EntidadesCampanasPPG.Modelo.Bitacora bitacora = new EntidadesCampanasPPG.Modelo.Bitacora();

            CampanaDAT campanaDAT = new CampanaDAT();
            GrupoReporteGR grupoReporteGR = new GrupoReporteGR(); 

            DataSet dsReporteGR = new DataSet();
            DataTable dtReporteGR1 = new DataTable();
            DataTable dtReporteGR2 = new DataTable();
            
            DataTable dtParametro = new DataTable();
            ParametroDAT parametroDAT = new ParametroDAT();
            List<Parametro> ListParametro = new List<Parametro>();
            Parametro parametro = new Parametro();

            string estatusReporteGR = string.Empty;
            string nombreArchivo = string.Empty;
            string pathArchivoCompleto = string.Empty;
            string urlCompleto = string.Empty;
            string pathArchivo = string.Empty;
            string url = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;
            string nombreReporte = string.Empty;
            string nombreReporteFinal = string.Empty;

            string nombreHojaCircular = string.Empty;
            string nombreTituloCircular = string.Empty;

            DataSet dsReporteCircular = new DataSet();
            DataTable dtReporteCircular = new DataTable();
            DataTable dtReporteCircularRes = new DataTable();
            List<string> ListAlcance = new List<string>();
            List<string> ListAlcanceTotal = new List<string>();
            string alcancePrimario = string.Empty;
            bool esAlcancePrimario = false;

            EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();
            string claveCampana = string.Empty;
            string nombreCampana = string.Empty;
            string fechaInicioSubCanal = string.Empty;
            string fechaFinSubCanal = string.Empty;
            string fechaInicioPublico = string.Empty;
            string fechaFinPublico = string.Empty;

            try
            {
                bitacora = bitacoraENTReq.ListBitacora.FirstOrDefault();

                if(bitacora == null)
                {
                    bitacoraENTRes.Mensaje = "ERROR: No se agregaron todos los datos necesarios para guardar informacion de la bitacora.";

                    return bitacoraENTRes;
                }

                respuesta = bitacoraDAT.GuardarBitacora(bitacora);

                if (respuesta >= 1)
                {
                    //OBTENER PARAMETROS
                    dtParametro = parametroDAT.GetParametro(0, null);

                    ListParametro = dtParametro.AsEnumerable()
                                    .Select(n => new Parametro
                                    {
                                        Id = n.Field<int?>("Id").GetValueOrDefault(),
                                        Nombre = n.Field<string>("Nombre"),
                                        Valor = n.Field<string>("Valor")
                                    }).ToList();

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["EstatusReporteGR"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        estatusReporteGR = parametro.Valor;
                    }


                    if (bitacora.Estatus.ToUpper() == estatusReporteGR.ToUpper())
                    {
                        #region REPORTE GRAN RED

                        //////////////////
                        //REPORTE GRAN RED
                        grupoReporteGR = campanaDAT.MostrarReporteGR(bitacora.IDCampania);

                        if((grupoReporteGR.ListCampana != null && grupoReporteGR.ListCampana.Count > 0) 
                            ||(grupoReporteGR.ListReporteGR1 != null && grupoReporteGR.ListReporteGR1.Count > 0)
                            || (grupoReporteGR.ListReporteGR2 != null && grupoReporteGR.ListReporteGR2.Count > 0))
                        {

                            parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["DirectorioReporte"].ToString().ToUpper()).FirstOrDefault();
                            if (parametro != null)
                            {
                                pathArchivo = parametro.Valor;
                            }

                            url = string.Empty;
                            parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["UrlReporteGR"].ToString().ToUpper()).FirstOrDefault();
                            if (parametro != null)
                            {
                                url = HttpUtility.HtmlEncode(parametro.Valor);
                            }

                            parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["NombreReporteGR"].ToString().ToUpper()).FirstOrDefault();
                            if (parametro != null)
                            {
                                nombreReporte = parametro.Valor + bitacora.ClaveCampania;
                            }

                            parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                               ConfigurationManager.AppSettings["UsuarioSharePoint"].ToString().ToUpper()).FirstOrDefault();
                            if (parametro != null)
                            {
                                usuarioSharePoint = parametro.Valor;
                            }

                            parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PasswordSharePoint"].ToString().ToUpper()).FirstOrDefault();
                            if (parametro != null)
                            {
                                passwordSharePoint = parametro.Valor;
                            }


                            //DATOS CAMPAÑA
                            if (grupoReporteGR.ListCampana.Count > 0)
                            {
                                campana = grupoReporteGR.ListCampana.FirstOrDefault();

                                claveCampana = campana.Title;
                                nombreCampana = campana.NombreCampa;
                                fechaInicioSubCanal = campana.FechaInicioSubCanal;
                                fechaFinSubCanal = campana.FechaFinSubCanal;
                                fechaInicioPublico = campana.FechaInicioPublico;
                                fechaFinPublico = campana.FechaFinPublico;
                            }

                            //AGREGAR PRIMER REPORTE
                            dtReporteGR1.Columns.Add("Id Tienda");
                            dtReporteGR1.Columns.Add("Bill To");
                            dtReporteGR1.Columns.Add("Descripcion Region");
                            dtReporteGR1.Columns.Add("Zona Territorial");
                            dtReporteGR1.Columns.Add("Customer Name");
                            dtReporteGR1.Columns.Add("Segmento");

                            grupoReporteGR.ListReporteGR1.ForEach(n =>
                            {
                                dtReporteGR1.Rows.Add(n.IdTienda, n.BillTo, n.DescripcionRegion,
                                                    n.ZonaTerritorial, n.CustomerName, n.Segmento);
                            });

                            //AGREGAR SEGUNDO REPORTE
                            dtReporteGR2.Columns.Add("Articulo");
                            dtReporteGR2.Columns.Add("Codigo EPR");
                            dtReporteGR2.Columns.Add("Descripcion");
                            dtReporteGR2.Columns.Add("Envase Descripcion");
                            dtReporteGR2.Columns.Add("Mecanica");
                            dtReporteGR2.Columns.Add("Estatus Producto");
                            dtReporteGR2.Columns.Add("Segmento");

                            grupoReporteGR.ListReporteGR2.ForEach(n =>
                            {
                                dtReporteGR2.Rows.Add(n.Articulo, n.CodigoEPR, n.Descripcion,
                                                    n.EnvaseDescripcion, n.Mecanica, n.EstatusProducto, 
                                                    n.Segmento);
                            });

                            //CREAR ARCHIVO EXCEL
                            dtReporteGR1.TableName = "GranRed_1";
                            //dsReporteGR.Tables.Add(dtReporteGR1);

                            dtReporteGR2.TableName = "GranRed_2";
                            //dsReporteGR.Tables.Add(dtReporteGR2);


                            nombreArchivo = Guid.NewGuid().ToString() + ".xlsx";
                            pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);

                            //ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoCompleto, dsReporteGR);

                            using (ExcelPackage excel = new ExcelPackage())
                            {
                                excel.Workbook.Worksheets.Add("Reporte_GrandRed_1");
                                excel.Workbook.Worksheets.Add("Reporte_GrandRed_2");

                                var worksheetGranRed_1 = excel.Workbook.Worksheets["Reporte_GrandRed_1"];
                                var worksheetGranRed_2 = excel.Workbook.Worksheets["Reporte_GrandRed_2"];

                                //PAGINA 1
                                worksheetGranRed_1.SetValue("A1", "Reporte Gran Red 1");
                                worksheetGranRed_1.SetValue("A2", "Campaña:");
                                worksheetGranRed_1.SetValue("B2", nombreCampana);
                                worksheetGranRed_1.SetValue("A3", "Clave Campaña:");
                                worksheetGranRed_1.SetValue("B3", claveCampana);

                                worksheetGranRed_1.SetValue("A4", "Fecha Inicio Publico");
                                if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico != "01/01/0001")
                                {
                                    worksheetGranRed_1.SetValue("B4", fechaInicioPublico);
                                }
                                worksheetGranRed_1.SetValue("A5", "Fecha Fin Publico");
                                if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico != "01/01/0001")
                                {
                                    worksheetGranRed_1.SetValue("B5", fechaFinPublico);
                                }

                                worksheetGranRed_1.Cells["A1:B5"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                worksheetGranRed_1.Cells["A1:B5"].Style.Font.Bold = true;
                                worksheetGranRed_1.Cells["A1:B5"].Style.Font.Size = 14;

                                worksheetGranRed_1.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR1.Columns.Count + 64) + "7"].LoadFromDataTable(dtReporteGR1, true, OfficeOpenXml.Table.TableStyles.Light2);

                                worksheetGranRed_1.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR1.Columns.Count + 64) + "7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheetGranRed_1.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR1.Columns.Count + 64) + "7"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                worksheetGranRed_1.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR1.Columns.Count + 64) + "7"].Style.Font.Color.SetColor(Color.White);

                                //PAGINA 2
                                worksheetGranRed_2.SetValue("A1", "Reporte Gran Red 2");
                                worksheetGranRed_2.SetValue("A2", "Campaña:");
                                worksheetGranRed_2.SetValue("B2", nombreCampana);
                                worksheetGranRed_2.SetValue("A3", "Clave Campaña:");
                                worksheetGranRed_2.SetValue("B3", claveCampana);

                                worksheetGranRed_2.SetValue("A4", "Fecha Inicio Publico");
                                if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico != "01/01/0001")
                                {
                                    worksheetGranRed_2.SetValue("B4", fechaInicioPublico);
                                }
                                worksheetGranRed_2.SetValue("A5", "Fecha Fin Publico");
                                if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico != "01/01/0001")
                                {
                                    worksheetGranRed_2.SetValue("B5", fechaFinPublico);
                                }

                                worksheetGranRed_2.Cells["A1:B5"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                worksheetGranRed_2.Cells["A1:B5"].Style.Font.Bold = true;
                                worksheetGranRed_2.Cells["A1:B5"].Style.Font.Size = 14;

                                worksheetGranRed_2.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR2.Columns.Count + 64) + "7"].LoadFromDataTable(dtReporteGR2, true, OfficeOpenXml.Table.TableStyles.Light2);

                                worksheetGranRed_2.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR2.Columns.Count + 64) + "7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheetGranRed_2.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR2.Columns.Count + 64) + "7"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                worksheetGranRed_2.Cells["A7:" + Char.ConvertFromUtf32(dtReporteGR2.Columns.Count + 64) + "7"].Style.Font.Color.SetColor(Color.White);

                                FileInfo excelFile = new FileInfo(pathArchivoCompleto);
                                excel.SaveAs(excelFile);
                            }

                            urlCompleto = Path.Combine(url, nombreReporte + ".xlsx");

                            using (WebClient client = new WebClient())
                            {
                                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                                client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                                client.UploadFile(urlCompleto, "PUT", pathArchivoCompleto);
                            }
                        }

                        #endregion

                        #region REPORTE CIRCULAR

                        ////////////////////////
                        //REPORTE CIRCULAR
                        ///////////////////////

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                            ConfigurationManager.AppSettings["AlcancePrimario"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            alcancePrimario = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["DirectorioReporte"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathArchivo = parametro.Valor;
                        }

                        url = string.Empty;
                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                            ConfigurationManager.AppSettings["UrlReporteCircular"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            url = HttpUtility.HtmlEncode(parametro.Valor);
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                            ConfigurationManager.AppSettings["NombreReporteCircular"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            nombreReporte = parametro.Valor + bitacora.ClaveCampania;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                           ConfigurationManager.AppSettings["UsuarioSharePoint"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            usuarioSharePoint = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                            ConfigurationManager.AppSettings["PasswordSharePoint"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            passwordSharePoint = parametro.Valor;
                        }

                        //OBTENER REPORTE CIRCULAR
                        dsReporteCircular = campanaDAT.MostrarReporteCircular(bitacora.IDCampania);

                        if (dsReporteCircular != null && dsReporteCircular.Tables.Count > 0)
                        {
                            //OBTENER ALCANCES
                            foreach(DataTable dtAlcance in dsReporteCircular.Tables)
                            {                               
                                ListAlcance = dtAlcance.AsEnumerable().GroupBy(n => n["Alcance"]).Select(m => m.Key.ToString()).ToList();

                                if (ListAlcance.Count > 0)
                                {
                                    ListAlcanceTotal.AddRange(ListAlcance);
                                }
                            }

                            //CREAR ARCHIVO POR ALCANCE Y AGREGAR HOJAS POR MECANICA
                            
                            foreach (string alcance in ListAlcanceTotal.Distinct())
                            {
                                using (ExcelPackage excel = new ExcelPackage())
                                {
                                    nombreReporteFinal = nombreReporte + "-" + alcance.ToUpper();

                                    nombreArchivo = Guid.NewGuid().ToString() + ".xlsx";
                                    pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);

                                    esAlcancePrimario = false;

                                    //IDENTIFICA SI ES NACIONAL
                                    if (alcance.ToUpper().Trim() == alcancePrimario)
                                    {
                                        esAlcancePrimario = true; 
                                    }

                                    dtReporteCircular = new DataTable();
                                    dtReporteCircularRes = new DataTable();

                                    //REGALO
                                    if (dsReporteCircular.Tables.Count > 0 && dsReporteCircular.Tables[0].Rows.Count > 0)
                                    {
                                        dtReporteCircular = dsReporteCircular.Tables[0];                                       

                                        dtReporteCircularRes = dtReporteCircular.Clone();

                                        //AGREGA DATOS NACIONAL
                                        if (dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).CopyToDataTable());
                                        }
                                        
                                        //AGREGA DATOS DIFERENTE DE NACIONAL
                                        if (!esAlcancePrimario && dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        nombreHojaCircular = "REGALO";

                                        excel.Workbook.Worksheets.Add(nombreHojaCircular);

                                        var worksheetCircular = excel.Workbook.Worksheets[nombreHojaCircular];

                                        worksheetCircular.SetValue("A1", "Padre");
                                        worksheetCircular.Cells["A1:B1"].Merge = true;
                                        worksheetCircular.SetValue("C1", "Hijo");
                                        worksheetCircular.Cells["C1:D1"].Merge = true;
                                        worksheetCircular.SetValue("E1", "Alcance");

                                        worksheetCircular.Cells["A1:E1"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A1:E1"].Style.Font.Bold = true;
                                        worksheetCircular.Cells["A1:E1"].Style.Font.Size = 14;
                                        worksheetCircular.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].LoadFromDataTable(dtReporteCircularRes, true, OfficeOpenXml.Table.TableStyles.Light2);

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Font.Color.SetColor(Color.White);
                                    }

                                    esAlcancePrimario = false;

                                    //IDENTIFICA SI ES NACIONAL
                                    if (alcance.ToUpper().Trim() == alcancePrimario)
                                    {
                                        esAlcancePrimario = true;
                                    }

                                    dtReporteCircular = new DataTable();
                                    dtReporteCircularRes = new DataTable();

                                    //MULTIPLO
                                    if (dsReporteCircular.Tables.Count > 1 && dsReporteCircular.Tables[1].Rows.Count > 0)
                                    {
                                        dtReporteCircular = dsReporteCircular.Tables[1];

                                        dtReporteCircularRes = dtReporteCircular.Clone();

                                        //AGREGA DATOS NACIONAL
                                        if(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).CopyToDataTable());
                                        }
                                        
                                        //AGREGA DATOS DIFERENTE DE NACIONAL
                                        if (!esAlcancePrimario && dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        nombreHojaCircular = "MULTIPLO";

                                        excel.Workbook.Worksheets.Add(nombreHojaCircular);

                                        var worksheetCircular = excel.Workbook.Worksheets[nombreHojaCircular];

                                        worksheetCircular.SetValue("A1", "Padre");
                                        worksheetCircular.Cells["A1:B1"].Merge = true;
                                        worksheetCircular.SetValue("C1", "Alcance");

                                        worksheetCircular.Cells["A1:C1"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A1:C1"].Style.Font.Bold = true;
                                        worksheetCircular.Cells["A1:C1"].Style.Font.Size = 14;
                                        worksheetCircular.Cells["A1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].LoadFromDataTable(dtReporteCircularRes, true, OfficeOpenXml.Table.TableStyles.Light2);

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Font.Color.SetColor(Color.White);
                                    }

                                    esAlcancePrimario = false;

                                    //IDENTIFICA SI ES NACIONAL
                                    if (alcance.ToUpper().Trim() == alcancePrimario)
                                    {
                                        esAlcancePrimario = true;
                                    }

                                    dtReporteCircular = new DataTable();
                                    dtReporteCircularRes = new DataTable();

                                    //DESCUENTO
                                    if (dsReporteCircular.Tables.Count > 2 && dsReporteCircular.Tables[2].Rows.Count > 0)
                                    {
                                        dtReporteCircular = dsReporteCircular.Tables[2];

                                        dtReporteCircularRes = dtReporteCircular.Clone();

                                        //AGREGA DATOS NACIONAL
                                        if (dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        //AGREGA DATOS DIFERENTE DE NACIONAL
                                        if (!esAlcancePrimario && dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        nombreHojaCircular = "DESCUENTO";

                                        excel.Workbook.Worksheets.Add(nombreHojaCircular);

                                        var worksheetCircular = excel.Workbook.Worksheets[nombreHojaCircular];

                                        worksheetCircular.SetValue("A1", "Padre");
                                        worksheetCircular.Cells["A1:B1"].Merge = true;
                                        worksheetCircular.SetValue("C1", "Alcance");

                                        worksheetCircular.Cells["A1:C1"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A1:C1"].Style.Font.Bold = true;
                                        worksheetCircular.Cells["A1:C1"].Style.Font.Size = 14;
                                        worksheetCircular.Cells["A1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].LoadFromDataTable(dtReporteCircularRes, true, OfficeOpenXml.Table.TableStyles.Light2);

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Font.Color.SetColor(Color.White);
                                    }

                                    esAlcancePrimario = false;

                                    //IDENTIFICA SI ES NACIONAL
                                    if (alcance.ToUpper().Trim() == alcancePrimario)
                                    {
                                        esAlcancePrimario = true;
                                    }

                                    dtReporteCircular = new DataTable();
                                    dtReporteCircularRes = new DataTable();

                                    //VOLUMEN
                                    if (dsReporteCircular.Tables.Count > 3 && dsReporteCircular.Tables[3].Rows.Count > 0)
                                    {
                                        dtReporteCircular = dsReporteCircular.Tables[3];

                                        dtReporteCircularRes = dtReporteCircular.Clone();

                                        //AGREGA DATOS NACIONAL
                                        if (dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        //AGREGA DATOS DIFERENTE DE NACIONAL
                                        if (!esAlcancePrimario && dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        nombreHojaCircular = "VOLUMEN";

                                        excel.Workbook.Worksheets.Add(nombreHojaCircular);

                                        var worksheetCircular = excel.Workbook.Worksheets[nombreHojaCircular];

                                        worksheetCircular.SetValue("A1", "Padre");
                                        worksheetCircular.Cells["A1:B1"].Merge = true;
                                        worksheetCircular.SetValue("C1", "De");
                                        worksheetCircular.SetValue("D1", "Hasta");
                                        worksheetCircular.SetValue("E1", "Descuento");
                                        worksheetCircular.SetValue("F1", "Alcance");

                                        worksheetCircular.Cells["A1:F1"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A1:F1"].Style.Font.Bold = true;
                                        worksheetCircular.Cells["A1:F1"].Style.Font.Size = 14;
                                        worksheetCircular.Cells["A1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].LoadFromDataTable(dtReporteCircularRes, true, OfficeOpenXml.Table.TableStyles.Light2);

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Font.Color.SetColor(Color.White);
                                    }

                                    esAlcancePrimario = false;

                                    //IDENTIFICA SI ES NACIONAL
                                    if (alcance.ToUpper().Trim() == alcancePrimario)
                                    {
                                        esAlcancePrimario = true;
                                    }

                                    dtReporteCircular = new DataTable();
                                    dtReporteCircularRes = new DataTable();

                                    //KIT
                                    if (dsReporteCircular.Tables.Count > 4 && dsReporteCircular.Tables[4].Rows.Count > 0)
                                    {
                                        dtReporteCircular = dsReporteCircular.Tables[4];

                                        dtReporteCircularRes = dtReporteCircular.Clone();

                                        //AGREGA DATOS NACIONAL
                                        if (dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        //AGREGA DATOS DIFERENTE DE NACIONAL
                                        if (!esAlcancePrimario && dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        nombreHojaCircular = "KIT";

                                        excel.Workbook.Worksheets.Add(nombreHojaCircular);

                                        var worksheetCircular = excel.Workbook.Worksheets[nombreHojaCircular];

                                        worksheetCircular.SetValue("A1", "Padre");
                                        worksheetCircular.Cells["A1:B1"].Merge = true;
                                        worksheetCircular.SetValue("C1", "Hijo");
                                        worksheetCircular.Cells["C1:D1"].Merge = true;
                                        worksheetCircular.SetValue("E1", "Descuento");
                                        worksheetCircular.SetValue("F1", "Importe");
                                        worksheetCircular.SetValue("G1", "Alcance");

                                        worksheetCircular.Cells["A1:G1"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A1:G1"].Style.Font.Bold = true;
                                        worksheetCircular.Cells["A1:G1"].Style.Font.Size = 14;
                                        worksheetCircular.Cells["A1:G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].LoadFromDataTable(dtReporteCircularRes, true, OfficeOpenXml.Table.TableStyles.Light2);

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Font.Color.SetColor(Color.White);
                                    }

                                    esAlcancePrimario = false;

                                    //IDENTIFICA SI ES NACIONAL
                                    if (alcance.ToUpper().Trim() == alcancePrimario)
                                    {
                                        esAlcancePrimario = true;
                                    }

                                    dtReporteCircular = new DataTable();
                                    dtReporteCircularRes = new DataTable();

                                    //COMBO
                                    if (dsReporteCircular.Tables.Count > 5 && dsReporteCircular.Tables[5].Rows.Count > 0)
                                    {
                                        dtReporteCircular = dsReporteCircular.Tables[5];

                                        dtReporteCircularRes = dtReporteCircular.Clone();

                                        //AGREGA DATOS NACIONAL
                                        if (dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcancePrimario.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        //AGREGA DATOS DIFERENTE DE NACIONAL
                                        if (!esAlcancePrimario && dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).ToList().Count > 0)
                                        {
                                            dtReporteCircularRes.Merge(dtReporteCircular.AsEnumerable().Where(n => n["Alcance"].ToString().ToUpper().Trim() == alcance.ToUpper().Trim()).CopyToDataTable());
                                        }

                                        nombreHojaCircular = "COMBO";

                                        excel.Workbook.Worksheets.Add(nombreHojaCircular);

                                        var worksheetCircular = excel.Workbook.Worksheets[nombreHojaCircular];

                                        worksheetCircular.SetValue("A1", "Padre");
                                        worksheetCircular.Cells["A1:B1"].Merge = true;
                                        worksheetCircular.SetValue("C1", "Madre");
                                        worksheetCircular.Cells["C1:D1"].Merge = true;
                                        worksheetCircular.SetValue("E1", "Hijo");
                                        worksheetCircular.Cells["E1:F1"].Merge = true;
                                        worksheetCircular.SetValue("G1", "Alcance");

                                        worksheetCircular.Cells["A1:G1"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A1:G1"].Style.Font.Bold = true;
                                        worksheetCircular.Cells["A1:G1"].Style.Font.Size = 14;
                                        worksheetCircular.Cells["A1:G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].LoadFromDataTable(dtReporteCircularRes, true, OfficeOpenXml.Table.TableStyles.Light2);

                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                                        worksheetCircular.Cells["A2:" + Char.ConvertFromUtf32(dtReporteCircularRes.Columns.Count + 64) + "2"].Style.Font.Color.SetColor(Color.White);
                                    }

                                    FileInfo excelFile = new FileInfo(pathArchivoCompleto);
                                    excel.SaveAs(excelFile);
                                }

                                urlCompleto = Path.Combine(url, nombreReporteFinal + ".xlsx");

                                using (WebClient client = new WebClient())
                                {
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                                    client.UploadFile(urlCompleto, "PUT", pathArchivoCompleto);
                                }

                            }

                        }

                        #endregion

                    }

                    bitacoraENTRes.Mensaje = "OK";
                }
                else
                {
                    bitacoraENTRes.Mensaje = "ERROR: Ocurrio un problema inesperado, identifique si se guardo correctamente la informacion, o consulte al administrador de sistemas.";

                    ArchivoLog.EscribirLog(null, "ERROR: Service - MostrarRentabilidad, Ocurrio un problema en el SP de Guardar Bitacora. ");
                }
            }
            catch (Exception ex)
            {
                bitacoraENTRes.Mensaje = "ERROR: " + ex.Message;

                ArchivoLog.EscribirLog(null, "ERROR: Service - MostrarRentabilidad, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return bitacoraENTRes;
        }
        public BitacoraENT MostrarBitacora(BitacoraENT bitacoraENTReq)
        {
            BitacoraENT bitacoraENTRes = new BitacoraENT();
            BitacoraDAT bitacoraDAT = new BitacoraDAT();
            EntidadesCampanasPPG.Modelo.Bitacora bitacora = new EntidadesCampanasPPG.Modelo.Bitacora();
            List<EntidadesCampanasPPG.Modelo.Bitacora> ListBitacora = new List<EntidadesCampanasPPG.Modelo.Bitacora>();
            DataTable dtBitacora = new DataTable();

            try
            {
                bitacora = bitacoraENTReq.ListBitacora.FirstOrDefault();

                if (bitacora == null)
                {
                    bitacoraENTRes.Mensaje = "ERROR: No se agregaron todos los datos necesarios para guardar informacion de la bitacora.";

                    return bitacoraENTRes;
                }

                dtBitacora = bitacoraDAT.MostrarBitacora(bitacora);

                ListBitacora = dtBitacora.AsEnumerable()
                                       .Select(row => new EntidadesCampanasPPG.Modelo.Bitacora
                                       {
                                           IDCampania = row.Field<int?>("ID").GetValueOrDefault(),
                                           ClaveCampania = row.Field<string>("Camp_Number"),
                                           NombreCampania = row.Field<string>("Nombre_Camp"),
                                           LiderCampania = row.Field<string>("Lider_Campania"),
                                           Estatus = row.Field<string>("Estatus"),
                                           Comentario = row.Field<string>("Comentario"),
                                           FechaCreacion = row.Field<DateTime?>("FechaCreacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                           //FechaInicio = row.Field<DateTime?>("FechaInicio").GetValueOrDefault(),
                                           //FechaFin = row.Field<DateTime?>("FechaFin").GetValueOrDefault(),
                                           FechaModificacion  = row.Field<DateTime?>("FechaModificacion").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                           PPGID = row.Field<string>("PPGID"),
                                           Usuario = row.Field<string>("NombreUsuario"),
                                           IdTarea = row.Field<int?>("ID_Tarea").GetValueOrDefault(),
                                           TipoFlujo = row.Field<int?>("TipoFlujo").GetValueOrDefault(),
                                           CorreoResponsable = row.Field<string>("CorreoUsuario"),
                                           Completado = row.Field<int?>("Completado").GetValueOrDefault(),
                                           Actividad = row.Field<string>("Actividad"),
                                           FechaInicioCrono = row.Field<DateTime?>("FechaInicioCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                           FechaInicioRealCrono = row.Field<DateTime?>("FechaInicioRealCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                           FechaFinCrono = row.Field<DateTime?>("FechaFinCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                           FechaFinRealCrono = row.Field<DateTime?>("FechaFinRealCrono").GetValueOrDefault().ToString("dd/MM/yyyy"),
                                           PPGIDCrono = row.Field<string>("PPGIDCrono"),
                                           PPGID2Crono = row.Field<string>("PPGID2Crono"),
                                           CorreoCrono = row.Field<string>("CorreoCrono"),
                                           CorreoCrono2 = row.Field<string>("CorreoCrono2"),
                                           NombreResponsable = row.Field<string>("NombreResponsable"),
                                           NombreResponsable2 = row.Field<string>("NombreResponsable_2"),
                                           TipoSubCanal = row.Field<string>("TipoSubCanal")
                                       }).ToList();

                bitacoraENTRes.ListBitacora = ListBitacora;
                bitacoraENTRes.Mensaje = "OK";
            }
            catch (Exception ex)
            {
                bitacoraENTRes.Mensaje = "ERROR: " + ex.Message;

                ArchivoLog.EscribirLog(null, "ERROR: Service - MostrarBitacora, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return bitacoraENTRes;
        }

    }
}

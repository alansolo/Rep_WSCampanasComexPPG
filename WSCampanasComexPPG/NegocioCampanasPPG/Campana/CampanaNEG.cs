using DatosCampanasPPG.Campana;
using EntidadesCampanasPPG.Modelo;
using EntidadesCampanasPPG.Producto;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.IO;
using System.Reflection;
using LinqToExcel;
using System.Configuration;
using System.Net;
using net.sf.mpxj;
using net.sf.mpxj.mpp;
using net.sf.mpxj.reader;
using EntidadesCampanasPPG.Reporte;
using UtilidadesCampanasPPG;
using System.DirectoryServices.AccountManagement;
using NegocioCampanasPPG.Ldap;
using DatosCampanasPPG.Catalogo;
using EntidadesCampanasPPG.FlujoActividad;
using EntidadesCampanasPPG.Campana;
using EntidadesCampanasPPG.GastoPlanta;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Spire.Pdf.HtmlConverter;
using System.Threading;
using ExpertPdf.HtmlToPdf;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System.Threading.Tasks;


namespace NegocioCampanasPPG.Campana
{
    public class CampanaNEG
    {
        public CampanaENT GuardarCampana(CampanaENT campanaENTReq)
        {
            int respuesta = 0;
            CampanaENT campanaENTRes = new CampanaENT();
            CampanaDAT campanaDAT = new CampanaDAT();
            EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();

            try
            {
                campana = campanaENTReq.Campana;

                respuesta = campanaDAT.GuardarCampanaCompleto(campana);

                if (respuesta == 1)
                {
                    campanaENTRes.Mensaje = "OK";
                }
                else
                {
                    campanaENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se guardo la informacion de Campañas, intente de nuevo, o consulte al administrador de sistemas.";

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCampana, Message: No se ejecuto correctametne el SP para guardar la informacion de Campañas.");
                }
            }
            catch (Exception ex)
            {
                campanaENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se guardo la informacion de Campañas, intente de nuevo, o consulte al administrador de sistemas.";

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCampana, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return campanaENTRes;
        }
        public CampanaENT GuardarCronograma(CampanaENT campanaENTReq)
        {
            int respuesta = 0;

            string url = string.Empty;
            string nombreArchivo = string.Empty;
            string pathArchivo = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;
            string pathArchivoCompleto = string.Empty;

            string incluirSi = string.Empty;
            string incluirNo = string.Empty;
            string tipoFlujoArchivo = string.Empty;
            string tipoFlujoAprobacion = string.Empty;
            string tipoFlujoInformativo = string.Empty;
            string tipoFlujoActualizar = string.Empty;

            CampanaENT campanaENTRes = new CampanaENT();

            LdapNEG ldapNEG = new LdapNEG();

            ParametroDAT parametroDAT = new ParametroDAT();
            DataTable dtParametro = new DataTable();
            List<Parametro> ListParametro = new List<Parametro>();
            Parametro parametro = new Parametro();

            List<Cronograma> ListCronograma = new List<Cronograma>();

            try
            {
                dtParametro = parametroDAT.GetParametro(0, null);

                ListParametro = dtParametro.AsEnumerable()
                                .Select(n => new Parametro
                                {
                                    Id = n.Field<int?>("Id").GetValueOrDefault(),
                                    Nombre = n.Field<string>("Nombre"),
                                    Valor = n.Field<string>("Valor")
                                }).ToList();

                EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();
                campana = campanaENTReq.Campana;

                CampanaDAT campanaDAT = new CampanaDAT();

                if (!string.IsNullOrEmpty(campana.UrlArchivoCronograma))
                {

                    url = HttpUtility.HtmlEncode(campana.UrlArchivoCronograma);
                    nombreArchivo = Guid.NewGuid().ToString() + System.IO.Path.GetExtension(url);

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["DirectorioLayoutCronograma"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        pathArchivo = parametro.Valor;
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

                    pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);


                    //DATOS PARA CATALOGO DE CRONOGRAMA
                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["IncluirSi"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        incluirSi = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["IncluirNo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        incluirNo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoArchivo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoArchivo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoAprobacion"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoAprobacion = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoInformativo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoInformativo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoActualizar"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoActualizar = parametro.Valor;
                    }


                    using (WebClient client = new WebClient())
                    {
                        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        client.DownloadFile(url, pathArchivoCompleto);
                    }

                    //LEER ARCHIVO EXCEL

                    IFormatProvider culture = new CultureInfo("es-MX", true);

                    try
                    {
                        var book = new ExcelQueryFactory(pathArchivoCompleto);

                        ListCronograma = book.Worksheet("Cronograma").AsEnumerable()
                                           .Select(row => new Cronograma
                                           {
                                               IDTarea = row["ID"].Cast<int>(),
                                               IDPadre = row["ID_Padre"].Cast<int>(),
                                               Actividad = row["Actividad"],
                                               Duracion = row["Duracion"].Cast<decimal>(),
                                               FechaInicio = row["Inicio"].Cast<DateTime>(),
                                               FechaFin = row["Final"].Cast<DateTime>(),
                                               TiempoOptimista = row["T_Optimista"].Cast<int>(),
                                               TiempoPesimista = row["T_Pesimista"].Cast<int>(),
                                               Predecesor = row["Predecesor"],
                                               Correo = row["Correo_Responsable"],
                                               Correo_2 = row["Correo_Responsable_2"],
                                               PorcentajeUsuario = row["Porcentaje"].Cast<decimal>(),
                                               PorcentajeSistema = 0,
                                               TipoFlujo = row["Tipo_Flujo"],
                                               IdTipoFlujo = row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoArchivo.ToUpper() ? 1 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoAprobacion.ToUpper() ? 2 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoInformativo.ToUpper() ? 3 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoActualizar.ToUpper() ? 4 : 0,
                                               Incluir = row["Incluir"],
                                               IdIncluir = row["Incluir"].ToString().ToUpper() == incluirSi.ToUpper() ? 1 :
                                                                row["Incluir"].ToString().ToUpper() == incluirNo.ToUpper() ? 0 : 0,
                                               EstatusEnvio = 0,
                                               UsuarioCreacion = campana.PPGIDRegistraCampa
                                           }).ToList();
                    }
                    catch (Exception ex)
                    {
                        campanaENTRes.OK = 0;

                        campanaENTRes.Mensaje = "ERROR: No se guardo la informacion de Campaña Cronograma, el archivo no tiene el formato correcto, no existe la Hoja \"Cronograma\" o los datos del archivo no son correctos.";

                        ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCronograma: Cargar Archivo Excel Hoja Cronograma, Source:" + ex.Source + ", Message:" + ex.Message);

                        campanaENTRes.MostrarCronograma = new MostrarCronograma();

                        return campanaENTRes;

                    }

                    //AGREGAR PROCESO PARA OBTENER EL NOMBRE DE LOS RESPONSABLE
                    #region ActiveDirectory

                    List<UsuarioLdap> ListUsuarioLdap = new List<UsuarioLdap>();
                    ListUsuarioLdap = ldapNEG.GetUsuarioLdap(ListParametro);

                    #endregion

                    if (ListCronograma.Count > 0)
                    {
                        List<Cronograma> ListCronogramaExcluir = ListCronograma.Where(n => n.IdIncluir == 0 && n.IDPadre == 1).ToList();

                        ListCronograma.RemoveAll(n => n.IdIncluir == 0);

                        ListCronogramaExcluir.ForEach(m =>
                        {
                            ListCronograma.RemoveAll(n => n.IDPadre == m.IDTarea);
                        });

                        UsuarioLdap usuarioLdap;

                        ListCronograma.ForEach(n =>
                        {
                            if (!string.IsNullOrEmpty(n.Correo))
                            {
                                usuarioLdap = ListUsuarioLdap.Where(m => m.Email == n.Correo.ToLower()).FirstOrDefault();

                                if (usuarioLdap != null)
                                {
                                    n.NombreResponsable = usuarioLdap.Nombre;
                                    n.PPGID = usuarioLdap.PPGID;
                                }
                            }

                            if (!string.IsNullOrEmpty(n.Correo_2))
                            {
                                usuarioLdap = ListUsuarioLdap.Where(m => m.Email == n.Correo_2.ToLower()).FirstOrDefault();

                                if (usuarioLdap != null)
                                {
                                    n.NombreResponsable_2 = usuarioLdap.Nombre;
                                    n.PPGID_2 = usuarioLdap.PPGID;
                                }
                            }

                            n.Padre = ListCronograma.Where(m => m.IDPadre == n.IDTarea).Count() > 0 ? true : false;
                        });
                    }

                    campana.ListaCronograma = ListCronograma;


                    respuesta = campanaDAT.GuardarCronogramaCompleto(campana);

                    if (respuesta == 1)
                    {
                        campanaENTRes.OK = 1;

                        campanaENTRes.Mensaje = "OK";
                    }
                    else
                    {
                        campanaENTRes.OK = 0;

                        campanaENTRes.Mensaje = "ERROR: Ocurrio un problema inesperado, no se guardo la informacion de Campaña y Cronograma, intente de nuevo o consulte al administrador de sistemas.";

                        ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCampanaCronograma, Message: Ocurrio un problema en el SP para guardar la informacion de Campaña y Cronograma.");
                    }

                }
                else
                {
                    campanaENTRes.OK = 0;

                    campanaENTRes.Mensaje = "ERROR: No se agrego la URL del archivo de cronograma.";

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCronograma, Message: No se agrego la URL del archivo de cronograma.");
                }
            }
            catch(Exception ex)
            {
                campanaENTRes.OK = 0;

                campanaENTRes.Mensaje = "ERROR: Ocurrio un problema inesperado, no se guardo la informacion de Campaña y Cronograma, intente de nuevo o consulte al administrador de sistemas.";

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCronograma, Source:" + ex.Source + ", Message:" + ex.Message);

            }

            return campanaENTRes;
        }
        public CampanaENT GuardarCampanaCronograma(CampanaENT campanaENTReq)
        {
            int respuesta = 0;
            string url = string.Empty;
            string nombreArchivo = string.Empty;
            string pathArchivo = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;
            string pathArchivoCompleto = string.Empty;
            CampanaENT campanaENTRes = new CampanaENT();

            List<Cronograma> ListCronograma = new List<Cronograma>();
            string responsable = string.Empty;
            string predecesor = string.Empty;

            string incluirSi = string.Empty;
            string incluirNo = string.Empty;
            string tipoFlujoArchivo = string.Empty;
            string tipoFlujoAprobacion = string.Empty;
            string tipoFlujoInformativo = string.Empty;
            string tipoFlujoActualizar = string.Empty;

            LdapNEG ldapNEG = new LdapNEG();

            ParametroDAT parametroDAT = new ParametroDAT();
            DataTable dtParametro = new DataTable();
            List<Parametro> ListParametro = new List<Parametro>();
            Parametro parametro = new Parametro();

            List<EntidadesCampanasPPG.Modelo.Campana> ListCampanaExiste = new List<EntidadesCampanasPPG.Modelo.Campana>(); 

            try
            {
                dtParametro = parametroDAT.GetParametro(0, null);

                ListParametro = dtParametro.AsEnumerable()
                                .Select(n => new Parametro
                                {
                                    Id = n.Field<int?>("Id").GetValueOrDefault(),
                                    Nombre = n.Field<string>("Nombre"),
                                    Valor = n.Field<string>("Valor")
                                }).ToList();

                EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();
                campana = campanaENTReq.Campana;

                CampanaDAT campanaDAT = new CampanaDAT();

                if (!string.IsNullOrEmpty(campana.UrlArchivoCronograma))
                {
                    
                    url = HttpUtility.HtmlEncode(campana.UrlArchivoCronograma);
                    nombreArchivo = Guid.NewGuid().ToString() + System.IO.Path.GetExtension(url);

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() == 
                        ConfigurationManager.AppSettings["DirectorioLayoutCronograma"].ToString().ToUpper()).FirstOrDefault();
                    if(parametro != null)
                    {
                        pathArchivo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["UsuarioSharePoint"].ToString().ToUpper()).FirstOrDefault();
                    if(parametro != null)
                    {
                        usuarioSharePoint = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["PasswordSharePoint"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        passwordSharePoint = parametro.Valor;
                    }

                    pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);


                    //DATOS PARA CATALOGO DE CRONOGRAMA
                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["IncluirSi"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        incluirSi = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["IncluirNo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        incluirNo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoArchivo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoArchivo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoAprobacion"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoAprobacion = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoInformativo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoInformativo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoActualizar"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoActualizar = parametro.Valor;
                    }


                    using (WebClient client = new WebClient())
                    {
                        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        client.DownloadFile(url, pathArchivoCompleto);
                    }

                    //LEER ARCHIVO EXCEL

                    IFormatProvider culture = new CultureInfo("es-MX", true);

                    try
                    {
                        var book = new ExcelQueryFactory(pathArchivoCompleto);

                        ListCronograma = book.Worksheet("Cronograma").AsEnumerable()
                                           .Select(row => new Cronograma
                                           {
                                               IDTarea = row["ID"].Cast<int>(),
                                               IDPadre = row["ID_Padre"].Cast<int>(),
                                               Actividad = row["Actividad"],
                                               Duracion = row["Duracion"].Cast<decimal>(),
                                               FechaInicio = row["Inicio"].Cast<DateTime>(),
                                               FechaFin = row["Final"].Cast<DateTime>(),
                                               TiempoOptimista = row["T_Optimista"].Cast<int>(),
                                               TiempoPesimista = row["T_Pesimista"].Cast<int>(),
                                               Predecesor = row["Predecesor"],
                                               Correo = row["Correo_Responsable"],
                                               Correo_2 = row["Correo_Responsable_2"],
                                               PorcentajeUsuario = row["Porcentaje"].Cast<decimal>(),
                                               PorcentajeSistema = 0,
                                               TipoFlujo = row["Tipo_Flujo"],
                                               IdTipoFlujo = row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoArchivo.ToUpper() ? 1 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoAprobacion.ToUpper() ? 2 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoInformativo.ToUpper() ? 3 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoActualizar.ToUpper() ? 4 : 0,
                                               Incluir = row["Incluir"],
                                               IdIncluir = row["Incluir"].ToString().ToUpper() == incluirSi.ToUpper() ? 1 :
                                                                row["Incluir"].ToString().ToUpper() == incluirNo.ToUpper() ? 0 : 0,
                                               EstatusEnvio = 0,
                                               UsuarioCreacion = campana.PPGIDRegistraCampa
                                           }).ToList();
                    }
                    catch (Exception ex)
                    {
                        campanaENTRes.OK = 0;

                        campanaENTRes.Mensaje = "ERROR: No se guardo la informacion de Campaña Cronograma, el archivo no tiene el formato correcto, no existe la Hoja \"Cronograma\" o los datos del archivo no son correctos.";

                        ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCampanaCronograma: Cargar Archivo Excel Hoja Cronograma, Source:" + ex.Source + ", Message:" + ex.Message);

                        campanaENTRes.MostrarCronograma = new MostrarCronograma();

                        return campanaENTRes;

                    }

                    //AGREGAR PROCESO PARA OBTENER EL NOMBRE DE LOS RESPONSABLE
                    #region ActiveDirectory

                    List<UsuarioLdap> ListUsuarioLdap = new List<UsuarioLdap>();
                    ListUsuarioLdap = ldapNEG.GetUsuarioLdap(ListParametro);

                    #endregion

                    if (ListCronograma.Count > 0)
                    {
                        List<Cronograma> ListCronogramaExcluir = ListCronograma.Where(n => n.IdIncluir == 0 && n.IDPadre == 1).ToList();

                        ListCronograma.RemoveAll(n => n.IdIncluir == 0);

                        ListCronogramaExcluir.ForEach(m =>
                        {
                            ListCronograma.RemoveAll(n => n.IDPadre == m.IDTarea);
                        });

                        UsuarioLdap usuarioLdap;

                        ListCronograma.ForEach(n =>
                        {
                            if (!string.IsNullOrEmpty(n.Correo))
                            {
                                usuarioLdap = ListUsuarioLdap.Where(m => m.Email == n.Correo.ToLower()).FirstOrDefault();

                                if (usuarioLdap != null)
                                {
                                    n.NombreResponsable = usuarioLdap.Nombre;
                                    n.PPGID = usuarioLdap.PPGID;
                                }
                            }

                            if (!string.IsNullOrEmpty(n.Correo_2))
                            {
                                usuarioLdap = ListUsuarioLdap.Where(m => m.Email == n.Correo_2.ToLower()).FirstOrDefault();

                                if (usuarioLdap != null)
                                {
                                    n.NombreResponsable_2 = usuarioLdap.Nombre;
                                    n.PPGID_2 = usuarioLdap.PPGID;
                                }
                            }

                            n.Padre = ListCronograma.Where(m => m.IDPadre == n.IDTarea).Count() > 0 ? true : false;
                        });
                    }

                    campana.ListaCronograma = ListCronograma;

                    //if (campana.Excepcion == 1)
                    //{
                    //    respuesta = campanaDAT.GuardarCampanaCronogramaCompleto(campana);
                    //}
                    //else
                    //{
                    //    //VALIDAR SI YA EXISTE LA CAMPAÑA
                    //    ListCampanaExiste = campanaDAT.MostrarCampana(campana);

                    //    if (ListCampanaExiste.Count > 0)
                    //    {
                    //        campana.ListaCronograma.ForEach(n =>
                    //        {
                    //            n.VersionCronograma = ListCampanaExiste.Max(m => m.VersionCronograma);
                    //        });

                    //        respuesta = campanaDAT.EditCampanaCronogramaCompleto(campana);
                    //    }
                    //    else
                    //    {
                    //        respuesta = campanaDAT.GuardarCampanaCronogramaCompleto(campana);
                    //    }
                    //}

                    respuesta = campanaDAT.GuardarCampanaCronogramaCompleto(campana);

                    if (respuesta == 1)
                    {
                        campanaENTRes.OK = 1;

                        campanaENTRes.Mensaje = "OK";
                    }
                    else
                    {
                        campanaENTRes.OK = 0;

                        campanaENTRes.Mensaje = "ERROR: Ocurrio un problema inesperado, no se guardo la informacion de Campaña y Cronograma, intente de nuevo o consulte al administrador de sistemas.";

                        ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCampanaCronograma, Message: Ocurrio un problema en el SP para guardar la informacion de Campaña y Cronograma.");
                    }
                }
                else
                {
                    campanaENTRes.OK = 0;

                    campanaENTRes.Mensaje = "ERROR: No se agrego la URL del archivo de cronograma.";

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCampanaCronograma, Message: No se agrego la URL del archivo de cronograma.");
                }
            }
            catch(Exception ex)
            {
                campanaENTRes.OK = 0;

                campanaENTRes.Mensaje = "ERROR: Ocurrio un problema inesperado, no se guardo la informacion de Campaña y Cronograma, intente de nuevo o consulte al administrador de sistemas.";

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarCampanaCronograma, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return campanaENTRes;
        }
        public CampanaENT MostrarCampanaCronograma(CampanaENT campanaENTReq)
        {
            CampanaENT campanaENTRes = new CampanaENT();

            bool correcto = true;
            bool correctoWarning = true;

            string url = string.Empty;
            string nombreArchivo = string.Empty;
            string pathArchivo = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;
            string pathArchivoCompleto = string.Empty;

            List<Cronograma> ListCronograma = new List<Cronograma>();
            List<Cronograma> ListCronogramaValidacion = new List<Cronograma>();

            string incluirSi = string.Empty;
            string incluirNo = string.Empty;
            string tipoFlujoArchivo = string.Empty;
            string tipoFlujoAprobacion = string.Empty;
            string tipoFlujoInformativo = string.Empty;
            string tipoFlujoActualizar = string.Empty;

            ParametroDAT parametroDAT = new ParametroDAT();
            DataTable dtParametro = new DataTable();
            List<Parametro> ListParametro = new List<Parametro>();
            Parametro parametro = new Parametro();

            DateTime dateCronograma = new DateTime();

            try
            {
                dtParametro = parametroDAT.GetParametro(0, null);

                ListParametro = dtParametro.AsEnumerable()
                                .Select(n => new Parametro
                                {
                                    Id = n.Field<int?>("Id").GetValueOrDefault(),
                                    Nombre = n.Field<string>("Nombre"),
                                    Valor = n.Field<string>("Valor")
                                }).ToList();

                EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();
                campana = campanaENTReq.Campana;

                if (!string.IsNullOrEmpty(campana.UrlArchivoCronograma))
                {
                    //LEER ARCHIVO
                    url = HttpUtility.HtmlEncode(campana.UrlArchivoCronograma);
                    nombreArchivo = Guid.NewGuid().ToString() + System.IO.Path.GetExtension(url);                

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["DirectorioLayoutCronograma"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        pathArchivo = parametro.Valor;
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

                    pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);


                    //DATOS PARA CATALOGO DE CRONOGRAMA
                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["IncluirSi"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        incluirSi = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["IncluirNo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        incluirNo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoArchivo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoArchivo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoAprobacion"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoAprobacion = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoInformativo"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoInformativo = parametro.Valor;
                    }

                    parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["TipoFlujoActualizar"].ToString().ToUpper()).FirstOrDefault();
                    if (parametro != null)
                    {
                        tipoFlujoActualizar = parametro.Valor;
                    }


                    using (WebClient client = new WebClient())
                    {
                        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                        client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                        client.DownloadFile(url, pathArchivoCompleto);
                    }

                    IFormatProvider culture = new CultureInfo("es-MX", true);

                    var book = new ExcelQueryFactory(pathArchivoCompleto);

                    try
                    {
                        ListCronograma = book.Worksheet("Cronograma").AsEnumerable()
                                           .Select(row => new Cronograma
                                           {
                                               IDTarea = row["ID"].Cast<int>(),
                                               IDPadre = row["ID_Padre"].Cast<int>(),
                                               Actividad = row["Actividad"],
                                               Duracion = row["Duracion"].Cast<int>(),
                                               FechaInicio = row["Inicio"].Cast<DateTime>(),
                                               StrFechaInicio = row["Inicio"].Cast<DateTime>().ToString("dd/MM/yyyy"),
                                               FechaFin = row["Final"].Cast<DateTime>(),
                                               StrFechaFin = row["Final"].Cast<DateTime>().ToString("dd/MM/yyyy"),
                                               TiempoOptimista = row["T_Optimista"].Cast<int>(),
                                               TiempoPesimista = row["T_Pesimista"].Cast<int>(),
                                               Predecesor = row["Predecesor"],
                                               Correo = row["Correo_Responsable"],
                                               Correo_2 = row["Correo_Responsable_2"],
                                               PorcentajeUsuario = row["Porcentaje"].Cast<decimal>(),
                                               PorcentajeSistema = 0,
                                               TipoFlujo = row["Tipo_Flujo"],
                                               IdTipoFlujo = row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoArchivo.ToUpper() ? 1 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoAprobacion.ToUpper() ? 2 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoInformativo.ToUpper() ? 3 :
                                                                row["Tipo_Flujo"].ToString().ToUpper() == tipoFlujoActualizar.ToUpper() ? 4 : 0,
                                               Incluir = row["Incluir"],
                                               IdIncluir = row["Incluir"].ToString().ToUpper() == incluirSi.ToUpper() ? 1 :
                                                                row["Incluir"].ToString().ToUpper() == incluirNo.ToUpper() ? 0 : 0,
                                               EstatusEnvio = 0,
                                               UsuarioCreacion = campana.PPGIDRegistraCampa
                                           }).ToList();
                    }
                    catch (Exception ex)
                    {
                        campanaENTRes.OK = 0;

                        campanaENTRes.Mensaje = "ERROR: No se guardo la informacion de Campaña Cronograma, el archivo no tiene el formato correcto, no existe la Hoja \"Cronograma\" o los datos del archivo no son correctos.";

                        ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja Cronograma, Source:" + ex.Source + ", Message:" + ex.Message);

                        campanaENTRes.MostrarCronograma = new MostrarCronograma();

                        campanaENTRes.MostrarCronograma.OK = 0;

                        campanaENTRes.MostrarCronograma.Mensaje = "ERROR: No se guardo la informacion de Campaña Cronograma, el archivo no tiene el formato correcto, no existe la Hoja \"Cronograma\" o los datos del archivo no son correctos.";

                        return campanaENTRes;

                    }


                    //VALIDAR DATOS COMPLETOS
                    ValidarDatos(ref correcto, ListCronograma);

                    //VALIDAR PREDECESOR
                    ValidarPredecesor(ref correctoWarning, ListCronograma);

                    //VALIDAR FECHA CON LA ACTUAL
                    ValidarFechasVsActual(ref correctoWarning, ListCronograma);

                    //VALIDAR FECHAS REALES
                    ValidarFechasReales(ref correcto, ListCronograma);

                    //VALIDAR FECHAS PROGRAMADAS
                    ValidarFechasProgramadas(ref correctoWarning, ListCronograma);

                    campanaENTRes.MostrarCronograma = new MostrarCronograma();
                    campanaENTRes.MostrarCronograma.ListaCronograma = ListCronograma;
                    campanaENTRes.MostrarCronograma.Correcto = correcto;
                    campanaENTRes.MostrarCronograma.CorrectoWarning = correctoWarning;
                    campanaENTRes.MostrarCronograma.OK = 1;
                    campanaENTRes.MostrarCronograma.Mensaje = "OK";

                    dateCronograma = ListCronograma.Min(n => n.FechaInicio);

                    if (dateCronograma != null)
                    {
                        campanaENTRes.MostrarCronograma.FechaInicio = dateCronograma.ToString("dd/MM/yyyy");
                    }

                    dateCronograma = ListCronograma.Max(n => n.FechaFin);

                    if (dateCronograma != null)
                    {
                        campanaENTRes.MostrarCronograma.FechaFinal = dateCronograma.ToString("dd/MM/yyyy");
                    }

                    campanaENTRes.OK = 1;
                    campanaENTRes.Mensaje = "OK";
                }
                else
                {
                    campanaENTRes.MostrarCronograma = new MostrarCronograma();
                    campanaENTRes.MostrarCronograma.OK = 0;
                    campanaENTRes.MostrarCronograma.Mensaje = "ERROR: No se agrego la URL del archivo de cronograma."; ;

                    campanaENTRes.OK = 0;
                    campanaENTRes.Mensaje = "ERROR: No se agrego la URL del archivo de cronograma.";

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampanaCronograma, Message: No se agrego la URL del archivo de cronograma.");
                }
            }
            catch(Exception ex)
            {
                campanaENTRes.MostrarCronograma = new MostrarCronograma();
                campanaENTRes.MostrarCronograma.OK = 0;
                campanaENTRes.MostrarCronograma.Mensaje = "Error: Ocurrio un error inesperado, no se pudo obtener la informacion de Campaña y Cronograma, intente de nuevo o consulte a su administrador de sistemas.";


                campanaENTRes.OK = 0;
                campanaENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se pudo obtener la informacion de Campaña y Cronograma, intente de nuevo o consulte a su administrador de sistemas.";

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampanaCronograma, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return campanaENTRes;
        }
        public ProductoLineaENT GuardarProductoCampana(ProductoLineaENT productoLineaENTReq)
        {
            List<LineaFamilia> ListLineaFamilia = new List<LineaFamilia>();
            string ClaveCampana = string.Empty;
            string url = string.Empty;
            string nombreArchivo = string.Empty;
            string pathArchivo = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;
            string pathArchivoCompleto = string.Empty;

            List<MecanicaRegalo> ListMecanicaRegalo = new List<MecanicaRegalo>();
            List<MecanicaMultiplo> ListMecanicaMultiplo = new List<MecanicaMultiplo>();
            List<MecanicaDescuento> ListMecanicaDescuento = new List<MecanicaDescuento>();
            List<MecanicaVolumen> ListMecanicaVolumen = new List<MecanicaVolumen>();
            List<MecanicaKit> ListMecanicaKit = new List<MecanicaKit>();
            List<MecanicaCombo> ListMecanicaCombo = new List<MecanicaCombo>();
            List<Tienda> ListTienda = new List<Tienda>();
            List<Tienda> ListTiendaExclusion = new List<Tienda>();
            List<Alcance> ListAlcance = new List<Alcance>();

            ProductoLineaENT productoLineaENTRes = new ProductoLineaENT();
            
            ParametroDAT parametroDAT = new ParametroDAT();
            DataTable dtParametro = new DataTable();
            List<Parametro> ListParametro = new List<Parametro>();
            Parametro parametro = new Parametro();

            string Mensaje = string.Empty;

            productoLineaENTRes.ListMensaje = new List<string>();

            int idTienda = 0;

            int intResultado = 0;
            decimal deciResultado = 0;

            int id = 0;

            //DataTable dtTienda = new DataTable();

            try
            {
                dtParametro = parametroDAT.GetParametro(0, null);

                ListParametro = dtParametro.AsEnumerable()
                                .Select(n => new Parametro
                                {
                                    Id = n.Field<int?>("Id").GetValueOrDefault(),
                                    Nombre = n.Field<string>("Nombre"),
                                    Valor = n.Field<string>("Valor")
                                }).ToList();

                ClaveCampana = productoLineaENTReq.ClaveCampana;
                url = HttpUtility.HtmlEncode(productoLineaENTReq.UrlArchivo);
                nombreArchivo = Guid.NewGuid().ToString() + System.IO.Path.GetExtension(url);

                parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["DirectorioLayoutCronograma"].ToString().ToUpper()).FirstOrDefault();
                if (parametro != null)
                {
                    pathArchivo = parametro.Valor;
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

                pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);

                using (WebClient client = new WebClient())
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                    client.DownloadFile(url, pathArchivoCompleto);
                }


                //CARGAR ARCHIVO EXCEL CON ESCENARIOS
                var book = new ExcelQueryFactory(pathArchivoCompleto);

                try
                {
                    //var ListaHojasExcel = book.GetWorksheetNames();

                    //REGALO
                    #region REGALO

                    ListMecanicaRegalo = book.Worksheet("REGALO").AsEnumerable()
                                   .Select(row => new MecanicaRegalo
                                   {
                                       ClaveCampana = ClaveCampana,
                                       Familia = row["Submarca"] + "-" + row["Tipo_Marca"],
                                       Submarca = row["Submarca"],
                                       Tipo_Marca = row["Tipo_Marca"],
                                       SKU = row["SKU"],
                                       Descripcion = row["Descripcion"],
                                       Tipo = row["Tipo"],
                                       Grupo_Regalo = !int.TryParse(row["Grupo_Regalo"], out intResultado) ? 0 : intResultado,
                                       NumeroHijo = !int.TryParse(row["Numero_Hijo"], out intResultado) ? 0 : intResultado,
                                       Alcance = row["Alcance"],
                                       Capacidad = row["Capacidad"],
                                       Dinamica = row["Dinamica"],
                                       VLitrosAnioAnt = !decimal.TryParse(row["Ventas_Litros_Año_Anterior"], out deciResultado) ? 0 : deciResultado,
                                       PLitrosSinCamp = !decimal.TryParse(row["Presupuesto_Litros_Sin_Campaña"], out deciResultado) ? 0 : deciResultado,
                                       PLitrosConCamp = !decimal.TryParse(row["Presupuesto_Litros_Con_Campaña"], out deciResultado) ? 0 : deciResultado
                                   }).ToList();

                    ListMecanicaRegalo.RemoveAll(n => 
                                        string.IsNullOrEmpty(n.Submarca.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.SKU.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Tipo.ToString().Trim()) &&
                                        n.Grupo_Regalo <= 0 &&
                                        string.IsNullOrEmpty(n.Alcance.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) &&
                                        n.PLitrosConCamp <= 0);

                    if (ListMecanicaRegalo
                            .Where(n => string.IsNullOrEmpty(n.Submarca.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.SKU.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Tipo.ToString().Trim()) ||
                                        n.Grupo_Regalo <= 0 ||
                                        //n.NumeroHijo == 0 &&
                                        string.IsNullOrEmpty(n.Alcance.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) ||
                                        //n.VLitrosAnioAnt == 0 &&
                                        //n.PLitrosSinCamp == 0 &&
                                        n.PLitrosConCamp <= 0
                                    ).Count() != 1)
                    {
                        //Submarca
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.Submarca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Submarca\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Submarca\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo_Marca
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.Tipo_Marca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo_Marca\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Tipo_Marca\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //SKU
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.SKU.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"SKU\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"SKU\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Descripcion
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.Descripcion.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Descripcion\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Descripcion\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.Tipo.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Tipo\", Hoja \"REGALO\"");

                            //return productoLineaENTRes;
                        }

                        //Grupo Regalo
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => n.Grupo_Regalo == 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Grupo_Regalo\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Grupo_Regalo\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Numero Hijo
                        //if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => n.NumeroHijo == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Numero_Hijo\", Hoja \"REGALO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Numero_Hijo\", Hoja \"REGALO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Alcance
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.Alcance.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Alcance\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Alcance\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Capacidad
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.Capacidad.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Capacidad\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Capacidad\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Dinamica
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => string.IsNullOrEmpty(n.Dinamica.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Dinamica\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Dinamica\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Venta Litros Año Anterior
                        //if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => n.VLitrosAnioAnt == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Ventas_Litros_Año_Anterior\", Hoja \"REGALO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Ventas_Litros_Año_Anterior\", Hoja \"REGALO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Presupuesto Litros Sin Campaña
                        //if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => n.PLitrosSinCamp == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"REGALO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"REGALO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Presupuesto Litros Con Campaña
                        if (ListMecanicaRegalo.Count > 0 && ListMecanicaRegalo.Where(n => n.PLitrosConCamp <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"REGALO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"REGALO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }
                    }


                    ListMecanicaRegalo = ListMecanicaRegalo
                            .Where(n => !string.IsNullOrEmpty(n.Submarca.ToString()) &&
                                        !string.IsNullOrEmpty(n.Tipo_Marca.ToString()) &&
                                        !string.IsNullOrEmpty(n.SKU.ToString()) &&
                                        !string.IsNullOrEmpty(n.Descripcion.ToString()) &&
                                        !string.IsNullOrEmpty(n.Tipo.ToString()) &&
                                        n.Grupo_Regalo > 0 &&
                                        //n.NumeroHijo != 0 &&
                                        !string.IsNullOrEmpty(n.Alcance.ToString()) &&
                                        !string.IsNullOrEmpty(n.Capacidad.ToString()) &&
                                        !string.IsNullOrEmpty(n.Dinamica.ToString()) &&
                                        //n.VLitrosAnioAnt == 0 &&
                                        //n.PLitrosSinCamp == 0 &&
                                        n.PLitrosConCamp > 0
                                    ).ToList();


                    #endregion
                }
                catch(Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"REGALO\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;

                }

                try
                {
                    //MULTIPLO
                    #region MULTIPLO           

                    id = 1;

                    ListMecanicaMultiplo = book.Worksheet("MULTIPLO").AsEnumerable()
                                    .Select(row => new MecanicaMultiplo
                                    {
                                        ClaveCampana = ClaveCampana,
                                        Familia = row["Submarca"] + "-" + row["Tipo_Marca"],
                                        Submarca = row["Submarca"],
                                        Tipo_Marca = row["Tipo_Marca"],
                                        SKU = row["SKU"],
                                        Descripcion = row["Descripcion"],
                                        Alcance = row["Alcance"],
                                        Capacidad = row["Capacidad"],
                                        Dinamica = row["Dinamica"],
                                        Multiplo_Padre = !int.TryParse(row["Multiplo_Padre"], out intResultado) ? 0 : intResultado,
                                        Multiplo_Hijo = !int.TryParse(row["Multiplo_Hijo"], out intResultado) ? 0 : intResultado,
                                        Punto_Venta = row["Punto_Venta"].Cast<string>(),
                                        VLitrosAnioAnt = !decimal.TryParse(row["Ventas_Litros_Año_Anterior"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosSinCamp = !decimal.TryParse(row["Presupuesto_Litros_Sin_Campaña"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosConCamp = !decimal.TryParse(row["Presupuesto_Litros_Con_Campaña"], out deciResultado) ? 0 : deciResultado,
                                        Grupo_Multiplo = id++                                      
                                    }).ToList();


                    ListMecanicaMultiplo.RemoveAll(n => 
                                        string.IsNullOrEmpty(n.Submarca.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.SKU.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Alcance.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) &&
                                        n.Multiplo_Padre <= 0 &&
                                        n.Multiplo_Hijo <= 0 &&
                                        string.IsNullOrEmpty(n.Punto_Venta.ToString().Trim()) &&
                                        n.PLitrosConCamp <= 0);

                    if (ListMecanicaMultiplo
                            .Where(n => 
                                        string.IsNullOrEmpty(n.Submarca.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.SKU.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Alcance.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) ||
                                        n.Multiplo_Padre <= 0 ||
                                        n.Multiplo_Hijo <= 0 ||
                                        string.IsNullOrEmpty(n.Punto_Venta.ToString().Trim()) ||
                                        //n.VLitrosAnioAnt == 0 &&
                                        //n.PLitrosSinCamp == 0 &&
                                        n.PLitrosConCamp <= 0
                                    ).Count() != 1)
                    {
                        //Submarca
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => string.IsNullOrEmpty(n.Submarca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Submarca\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Submarca\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo_Marca
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => string.IsNullOrEmpty(n.Tipo_Marca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo_Marca\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Tipo_Marca\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //SKU
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => string.IsNullOrEmpty(n.SKU.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"SKU\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"SKU\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Descripcion
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => string.IsNullOrEmpty(n.Descripcion.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Descripcion\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Descripcion\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Alcance
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => string.IsNullOrEmpty(n.Alcance.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Alcance\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Alcance\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Capacidad
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => string.IsNullOrEmpty(n.Capacidad.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Capacidad\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Capacidad\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Multiplo_Padre
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => n.Multiplo_Padre == 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Multiplo_Padre\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Multiplo_Padre\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Multiplo_Hijo
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => n.Multiplo_Hijo == 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Multiplo_Hijo\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Multiplo_Hijo\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Punto_Venta
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => string.IsNullOrEmpty(n.Punto_Venta.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Punto_Venta\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja REGALO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Punto_Venta\", Hoja \"REGALO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Ventas_Litros_Año_Anterior
                        //if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => n.VLitrosAnioAnt == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Ventas_Litros_Año_Anterior\", Hoja \"MULTIPLO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Ventas_Litros_Año_Anterior\", Hoja \"MULTIPLO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Presupuesto_Litros_Sin_Campaña
                        //if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => n.PLitrosSinCamp == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"MULTIPLO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"MULTIPLO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Presupuesto_Litros_Con_Campaña
                        if (ListMecanicaMultiplo.Count > 0 && ListMecanicaMultiplo.Where(n => n.PLitrosConCamp <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"MULTIPLO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + "Carga Hoja \"MULTIPLO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"MULTIPLO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }
                    }


                    ListMecanicaMultiplo = ListMecanicaMultiplo
                            .Where(n => 
                                        !string.IsNullOrEmpty(n.Submarca.ToString()) &&
                                        !string.IsNullOrEmpty(n.Tipo_Marca.ToString()) &&
                                        !string.IsNullOrEmpty(n.SKU.ToString()) &&
                                        !string.IsNullOrEmpty(n.Descripcion.ToString()) &&
                                        !string.IsNullOrEmpty(n.Alcance.ToString()) &&
                                        !string.IsNullOrEmpty(n.Capacidad.ToString()) &&
                                        !string.IsNullOrEmpty(n.Dinamica.ToString()) &&
                                        n.Multiplo_Padre > 0 &&
                                        n.Multiplo_Hijo > 0 &&
                                        !string.IsNullOrEmpty(n.Punto_Venta.ToString()) &&
                                        //n.VLitrosAnioAnt == 0 &&
                                        //n.PLitrosSinCamp == 0 &&
                                        n.PLitrosConCamp > 0
                                    ).ToList();


                    #endregion
                }
                catch(Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"MULTIPLO\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja MULTIPLO, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }

                try
                {
                    //DESCUENTO
                    #region DESCUENTO

                    id = 1;

                    ListMecanicaDescuento = book.Worksheet("DESCUENTO").AsEnumerable()
                                    .Select(row => new MecanicaDescuento
                                    {
                                        ClaveCampana = ClaveCampana,
                                        Familia = row["Submarca"] + "-" + row["Tipo_Marca"],
                                        Submarca = row["Submarca"],
                                        Tipo_Marca = row["Tipo_Marca"],
                                        SKU = row["SKU"],
                                        Descripcion = row["Descripcion"],
                                        Alcance = row["Alcance"],
                                        Capacidad = row["Capacidad"],
                                        Dinamica = row["Dinamica"],
                                        Porcentaje = !decimal.TryParse(row["Porcentaje"], out deciResultado) ? 0 : deciResultado,
                                        Importe = !decimal.TryParse(row["Importe"], out deciResultado) ? 0 : deciResultado,
                                        VLitrosAnioAnt = !decimal.TryParse(row["Ventas_Litros_Año_Anterior"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosSinCamp = !decimal.TryParse(row["Presupuesto_Litros_Sin_Campaña"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosConCamp = !decimal.TryParse(row["Presupuesto_Litros_Con_Campaña"], out deciResultado) ? 0 : deciResultado,
                                        Grupo_Descuento = id++
                                    }).ToList();

                    ListMecanicaDescuento.RemoveAll(n => 
                                        string.IsNullOrEmpty(n.Submarca.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.SKU.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Alcance.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) &&
                                        string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) &&
                                        n.Porcentaje <= 0 &&
                                        n.Importe <= 0 &&
                                        n.PLitrosConCamp <= 0);

                    if (ListMecanicaDescuento
                            .Where(n => 
                                        string.IsNullOrEmpty(n.Submarca.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.SKU.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Alcance.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) ||
                                        string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) ||
                                        (n.Porcentaje <= 0 ||
                                        n.Importe <= 0) ||
                                        //n.VLitrosAnioAnt == 0 &&
                                        //n.PLitrosSinCamp == 0 &&
                                        n.PLitrosConCamp <= 0
                                    ).Count() != 1)
                    {
                        //Submarca
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => string.IsNullOrEmpty(n.Submarca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Submarca\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Submarca\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo_Marca
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => string.IsNullOrEmpty(n.Tipo_Marca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo_Marca\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Tipo_Marca\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //SKU
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => string.IsNullOrEmpty(n.SKU.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"SKU\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"SKU\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Descripcion
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => string.IsNullOrEmpty(n.Descripcion.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Descripcion\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Descripcion\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            return productoLineaENTRes;
                        }

                        //Alcance
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => string.IsNullOrEmpty(n.Alcance.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Alcance\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Alcance\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Capacidad
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => string.IsNullOrEmpty(n.Capacidad.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Capacidad\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Capacidad\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Dinamica
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => string.IsNullOrEmpty(n.Dinamica.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Dinamica\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Dinamica\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Porcentaje o Importe
                        if (ListMecanicaDescuento.Count > 0 && (ListMecanicaDescuento.Where(n => n.Porcentaje <= 0).Count() > 0 &&
                            ListMecanicaDescuento.Where(n => n.Importe <= 0).Count() > 0))
                        {
                            Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, debe agregar informacion en la columna \"Porcentaje\" o \"Importe\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "debe agregar informacion en la columna \"Porcentaje\" o \"Importe\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Importe
                        //if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => n.Importe == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Importe\", Hoja \"DESCUENTO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Importe\", Hoja \"DESCUENTO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Ventas_Litros_Año_Anterior
                        //if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => n.VLitrosAnioAnt == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Ventas_Litros_Año_Anterior\", Hoja \"DESCUENTO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Ventas_Litros_Año_Anterior\", Hoja \"DESCUENTO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Presupuesto_Litros_Sin_Campaña
                        //if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => n.PLitrosSinCamp == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"DESCUENTO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"DESCUENTO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Presupuesto_Litros_Con_Campaña
                        if (ListMecanicaDescuento.Count > 0 && ListMecanicaDescuento.Where(n => n.PLitrosConCamp <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"DESCUENTO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + "Carga Hoja \"DESCUENTO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"DESCUENTO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                    }

                    ListMecanicaDescuento = ListMecanicaDescuento
                            .Where(n => 
                                        !string.IsNullOrEmpty(n.Submarca.ToString()) &&
                                        !string.IsNullOrEmpty(n.Tipo_Marca.ToString()) &&
                                        !string.IsNullOrEmpty(n.SKU.ToString()) &&
                                        !string.IsNullOrEmpty(n.Descripcion.ToString()) &&
                                        !string.IsNullOrEmpty(n.Alcance.ToString()) &&
                                        !string.IsNullOrEmpty(n.Capacidad.ToString()) &&
                                        !string.IsNullOrEmpty(n.Dinamica.ToString()) &&
                                        (n.Porcentaje > 0 ||
                                        n.Importe > 0) &&
                                        //n.VLitrosAnioAnt == 0 &&
                                        //n.PLitrosSinCamp == 0 &&
                                        n.PLitrosConCamp > 0
                            ).ToList();

                    #endregion
                }
                catch(Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"DESCUENTO\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja DESCUENTO, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }

                try
                {
                    //VOLUMEN
                    #region VOLUMEN

                    ListMecanicaVolumen = book.Worksheet("VOLUMEN").AsEnumerable()
                                   .Select(row => new MecanicaVolumen
                                   {
                                       ClaveCampana = ClaveCampana,
                                       Familia = row["Submarca"] + "-" + row["Tipo_Marca"],
                                       Submarca = row["Submarca"],
                                       Tipo_Marca = row["Tipo_Marca"],
                                       SKU = row["SKU"],
                                       Descripcion = row["Descripcion"],
                                       Bloque = !int.TryParse(row["De"], out intResultado) ? 0 : intResultado,
                                       Alcance = row["Alcance"],
                                       Capacidad = row["Capacidad"],
                                       Dinamica = row["Dinamica"],
                                       De = !decimal.TryParse(row["De"], out deciResultado) ? 0 : deciResultado,
                                       Hasta = !decimal.TryParse(row["Hasta"], out deciResultado) ? 0 : deciResultado,
                                       Descuento = !decimal.TryParse(row["Descuento"], out deciResultado) ? 0 : deciResultado,
                                       VLitrosAnioAnt = !decimal.TryParse(row["Ventas_Litros_Año_Anterior"], out deciResultado) ? 0 : deciResultado,
                                       PLitrosSinCamp = !decimal.TryParse(row["Presupuesto_Litros_Sin_Campaña"], out deciResultado) ? 0 : deciResultado,
                                       PLitrosConCamp = !decimal.TryParse(row["Presupuesto_Litros_Con_Campaña"], out deciResultado) ? 0 : deciResultado
                                   }).ToList();


                    ListMecanicaVolumen.RemoveAll(n => 
                                       string.IsNullOrEmpty(n.Submarca.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.SKU.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) &&
                                       n.Bloque <= 0 &&
                                       string.IsNullOrEmpty(n.Alcance.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) &&
                                       n.De <= 0 &&
                                       n.Hasta <= 0 &&
                                       n.Descuento <= 0 &&
                                       n.PLitrosConCamp <= 0);

                    if (ListMecanicaVolumen
                           .Where(n => 
                                       string.IsNullOrEmpty(n.Submarca.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.SKU.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) ||
                                       n.Bloque <= 0 ||
                                       string.IsNullOrEmpty(n.Alcance.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) ||
                                       n.De <= 0 ||
                                       n.Hasta <= 0 ||
                                       n.Descuento <= 0 ||
                                       //n.VLitrosAnioAnt == 0 &&
                                       //n.PLitrosSinCamp == 0 &&
                                       n.PLitrosConCamp <= 0
                           ).Count() != 1)
                    {
                        //Submarca
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => string.IsNullOrEmpty(n.Submarca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Submarca\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Submarca\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo_Marca
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => string.IsNullOrEmpty(n.Tipo_Marca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo_Marca\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Tipo_Marca\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //SKU
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => string.IsNullOrEmpty(n.SKU.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"SKU\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja SKU, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"SKU\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Descripcion
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => string.IsNullOrEmpty(n.Descripcion.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Descripcion\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Descripcion\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Bloque
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.Bloque == 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Bloque\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Bloque\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Alcance
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => string.IsNullOrEmpty(n.Alcance.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Alcance\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Alcance\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Capacidad
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => string.IsNullOrEmpty(n.Capacidad.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Capacidad\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Capacidad\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Dinamica
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => string.IsNullOrEmpty(n.Dinamica.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Dinamica\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Dinamica\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //De
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.De < 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"De\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"De\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Hasta
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.Hasta < 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Hasta\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Hasta\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //VALIDAR DE Y HASTA
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.De > n.Hasta && n.Hasta != 0).Count() > 0)
                        {
                            Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, la columna \"De\" es mayor a \"Hasta\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + " la columna \"De\" es mayor a \"Hasta\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Descuento
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.Descuento <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Descuento\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Descuento\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Ventas_Litros_Año_Anterior
                        //if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.VLitrosAnioAnt == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Ventas_Litros_Año_Anterior\", Hoja \"VOLUMEN\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Ventas_Litros_Año_Anterior\", Hoja \"VOLUMEN\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Presupuesto_Litros_Sin_Campaña
                        //if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.PLitrosSinCamp == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"VOLUMEN\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"VOLUMEN\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Presupuesto_Litros_Con_Campaña
                        if (ListMecanicaVolumen.Count > 0 && ListMecanicaVolumen.Where(n => n.PLitrosConCamp <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"VOLUMEN\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + "Carga Hoja \"VOLUMEN\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"VOLUMEN\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }
                    }

                    ListMecanicaVolumen = ListMecanicaVolumen
                           .Where(n => 
                                       !string.IsNullOrEmpty(n.Submarca.ToString()) &&
                                       !string.IsNullOrEmpty(n.Tipo_Marca.ToString()) &&
                                       !string.IsNullOrEmpty(n.SKU.ToString()) &&
                                       !string.IsNullOrEmpty(n.Descripcion.ToString()) &&
                                       n.Bloque > 0 &&
                                       !string.IsNullOrEmpty(n.Alcance.ToString()) &&
                                       !string.IsNullOrEmpty(n.Capacidad.ToString()) &&
                                       !string.IsNullOrEmpty(n.Dinamica.ToString()) &&
                                       n.De >= 0 &&
                                       n.Hasta >= 0 &&
                                       n.Descuento > 0 &&
                                       //n.VLitrosAnioAnt == 0 &&
                                       //n.PLitrosSinCamp == 0 &&
                                       n.PLitrosConCamp > 0
                           ).ToList();


                    #endregion
                }
                catch(Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"VOLUMEN\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja VOLUMEN, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }

                try
                {
                    //KIT
                    #region KIT

                    ListMecanicaKit = book.Worksheet("KIT").AsEnumerable()
                                    .Select(row => new MecanicaKit
                                    {
                                        ClaveCampana = ClaveCampana,
                                        Familia = row["Submarca"] + "-" + row["Tipo_Marca"],
                                        Submarca = row["Submarca"],
                                        Tipo_Marca = row["Tipo_Marca"],
                                        SKU = row["SKU"],
                                        Grupo_Kit = !int.TryParse(row["Grupo_Kit"], out intResultado) ? 0 : intResultado,
                                        Tipo = row["Tipo"],
                                        NumeroHijo = !int.TryParse(row["Numero_Hijo"], out intResultado) ? 0 : intResultado,
                                        Descripcion = row["Descripcion"],
                                        Alcance = row["Alcance"],
                                        Capacidad = row["Capacidad"],
                                        Dinamica = row["Dinamica"],
                                        Porcentaje = !decimal.TryParse(row["Porcentaje"], out deciResultado) ? 0 : deciResultado,
                                        Importe = !decimal.TryParse(row["Importe"], out deciResultado) ? 0 : deciResultado,
                                        VLitrosAnioAnt = !decimal.TryParse(row["Ventas_Litros_Año_Anterior"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosSinCamp = !decimal.TryParse(row["Presupuesto_Litros_Sin_Campaña"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosConCamp = !decimal.TryParse(row["Presupuesto_Litros_Con_Campaña"], out deciResultado) ? 0 : deciResultado
                                    }).ToList();

                    ListMecanicaKit.RemoveAll(n => 
                                       string.IsNullOrEmpty(n.Submarca.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.SKU.ToString().Trim()) &&
                                       n.Grupo_Kit <= 0 &&
                                       string.IsNullOrEmpty(n.Tipo.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Alcance.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) &&
                                       n.Porcentaje <= 0 &&
                                       n.Importe <= 0 &&
                                       n.PLitrosConCamp <= 0);

                    if (ListMecanicaKit
                           .Where(n => 
                                       string.IsNullOrEmpty(n.Submarca.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.SKU.ToString().Trim()) ||
                                       n.Grupo_Kit == 0 ||
                                       string.IsNullOrEmpty(n.Tipo.ToString().Trim()) ||
                                       //n.NumeroHijo == 0 &&
                                       string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Alcance.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) ||
                                       n.Porcentaje <= 0 ||
                                       n.Importe <= 0 ||
                                       //n.VLitrosAnioAnt == 0 &&
                                       //n.PLitrosSinCamp == 0 &&
                                       n.PLitrosConCamp <= 0
                           ).Count() != 1)
                    {
                        //Submarca
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.Submarca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Submarca\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Submarca\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo_Marca
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.Tipo_Marca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo_Marca\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Tipo_Marca\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //SKU
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.SKU.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"SKU\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"SKU\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Grupo_Kit
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => n.Grupo_Kit == 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Grupo_Kit\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Grupo_Kit\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.Tipo.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Tipo\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Numero_Hijo
                        //if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => n.NumeroHijo == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Numero_Hijo\", Hoja \"KIT\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Numero_Hijo\", Hoja \"KIT\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Descripcion
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.Descripcion.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Descripcion\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Descripcion\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Alcance
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.Alcance.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Alcance\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Alcance\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Capacidad
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.Capacidad.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Capacidad\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Capacidad\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Dinamica
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => string.IsNullOrEmpty(n.Dinamica.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Dinamica\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Dinamica\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Porcentaje o Importe
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => n.Porcentaje <= 0 && n.Importe <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Porcentaje\" o \"Importe\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "debe agregar informacion en la columna \"Porcentaje\" o \"Importe\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Importe
                        //if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => n.Importe == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Importe\", Hoja \"KIT\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Importe\", Hoja \"KIT\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Ventas_Litros_Año_Anterior
                        //if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => n.VLitrosAnioAnt == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Ventas_Litros_Año_Anterior\", Hoja \"KIT\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Ventas_Litros_Año_Anterior\", Hoja \"KIT\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Presupuesto_Litros_Sin_Campaña
                        //if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => n.PLitrosSinCamp == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"KIT\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"KIT\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Presupuesto_Litros_Con_Campaña
                        if (ListMecanicaKit.Count > 0 && ListMecanicaKit.Where(n => n.PLitrosConCamp <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"KIT\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + "Carga Hoja \"KIT\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"KIT\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }
                    }


                    ListMecanicaKit = ListMecanicaKit
                           .Where(n => 
                                       !string.IsNullOrEmpty(n.Submarca.ToString()) &&
                                       !string.IsNullOrEmpty(n.Tipo_Marca.ToString()) &&
                                       !string.IsNullOrEmpty(n.SKU.ToString()) &&
                                       n.Grupo_Kit > 0 &&
                                       !string.IsNullOrEmpty(n.Tipo.ToString()) &&
                                       //n.NumeroHijo == 0 &&
                                       !string.IsNullOrEmpty(n.Descripcion.ToString()) &&
                                       !string.IsNullOrEmpty(n.Alcance.ToString()) &&
                                       !string.IsNullOrEmpty(n.Capacidad.ToString()) &&
                                       !string.IsNullOrEmpty(n.Dinamica.ToString()) &&
                                       (n.Porcentaje > 0 ||
                                       n.Importe > 0) &&
                                       //n.VLitrosAnioAnt == 0 &&
                                       //n.PLitrosSinCamp == 0 &&
                                       n.PLitrosConCamp > 0
                           ).ToList();

                    #endregion
                }
                catch(Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"KIT\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja KIT, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }

                try
                {
                    //COMBO
                    #region COMBO

                    ListMecanicaCombo = book.Worksheet("COMBO").AsEnumerable()
                                    .Select(row => new MecanicaCombo
                                    {
                                        ClaveCampana = ClaveCampana,
                                        Familia = row["Submarca"] + "-" + row["Tipo_Marca"],
                                        Submarca = row["Submarca"],
                                        Tipo_Marca = row["Tipo_Marca"],
                                        SKU = row["SKU"],
                                        Descripcion = row["Descripcion"],
                                        Tipo = row["Tipo"],
                                        Grupo_Combo = !int.TryParse(row["Grupo_Combo"], out intResultado) ? 0 : intResultado,
                                        NumeroPadre = !int.TryParse(row["Numero_Padre"], out intResultado) ? 0 : intResultado,
                                        NumeroHijo = !int.TryParse(row["Numero_Hijo"], out intResultado) ? 0 : intResultado,
                                        Alcance = row["Alcance"],
                                        Capacidad = row["Capacidad"],
                                        Dinamica = row["Dinamica"],
                                        VLitrosAnioAnt = !decimal.TryParse(row["Ventas_Litros_Año_Anterior"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosSinCamp = !decimal.TryParse(row["Presupuesto_Litros_Sin_Campaña"], out deciResultado) ? 0 : deciResultado,
                                        PLitrosConCamp = !decimal.TryParse(row["Presupuesto_Litros_Con_Campaña"], out deciResultado) ? 0 : deciResultado
                                    }).ToList();


                    ListMecanicaCombo.RemoveAll(n => 
                                       string.IsNullOrEmpty(n.Submarca.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.SKU.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Tipo.ToString().Trim()) &&
                                       n.Grupo_Combo <= 0 &&
                                       n.NumeroPadre <= 0 &&
                                       n.NumeroHijo <= 0 &&
                                       string.IsNullOrEmpty(n.Alcance.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) &&
                                       string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) &&
                                       n.PLitrosConCamp <= 0);

                    if (ListMecanicaCombo
                           .Where(n => 
                                       string.IsNullOrEmpty(n.Submarca.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Tipo_Marca.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.SKU.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Descripcion.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Tipo.ToString().Trim()) ||
                                       n.Grupo_Combo <= 0 ||
                                       n.NumeroPadre <= 0 ||
                                       n.NumeroHijo <= 0 ||
                                       string.IsNullOrEmpty(n.Alcance.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Capacidad.ToString().Trim()) ||
                                       string.IsNullOrEmpty(n.Dinamica.ToString().Trim()) ||
                                       //n.VLitrosAnioAnt == 0 &&
                                       //n.PLitrosSinCamp == 0 &&
                                       n.PLitrosConCamp <= 0
                           ).Count() != 1)
                    {
                        //Submarca
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.Submarca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Submarca\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Submarca\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo_Marca
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.Tipo_Marca.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo_Marca\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Tipo_Marca\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //SKU
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.SKU.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"SKU\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"SKU\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Descripcion
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.Descripcion.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Descripcion\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Descripcion\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.Tipo.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Tipo\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Grupo_Combo
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => n.Grupo_Combo == 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Grupo_Combo\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Grupo_Combo\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Numero_Padre o Numero_Hijo
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => n.NumeroPadre <= 0 && n.NumeroHijo <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, debe agreagar informacion en la columna \"Numero_Padre\" o \"Numero_Hijo\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "debe agregar informacion en la columna \"Numero_Padre\" o \"Numero_Hijo\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Numero_Hijo
                        //if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => n.NumeroHijo == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Numero_Hijo\", Hoja \"COMBO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Numero_Hijo\", Hoja \"COMBO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Alcance
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.Alcance.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Alcance\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Alcance\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Capacidad
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.Capacidad.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Capacidad\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Capacidad\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Dinamica
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => string.IsNullOrEmpty(n.Dinamica.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Dinamica\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Dinamica\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Ventas_Litros_Año_Anterior
                        //if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => n.VLitrosAnioAnt == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Ventas_Litros_Año_Anterior\", Hoja \"COMBO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Ventas_Litros_Año_Anterior\", Hoja \"COMBO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Presupuesto_Litros_Sin_Campaña
                        //if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => n.PLitrosSinCamp == 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"COMBO\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Sin_Campaña\", Hoja \"COMBO\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Presupuesto_Litros_Con_Campaña
                        if (ListMecanicaCombo.Count > 0 && ListMecanicaCombo.Where(n => n.PLitrosConCamp <= 0).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"COMBO\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + "Carga Hoja \"COMBO\"" + ", Message:" + "falta informacion en la columna \"Presupuesto_Litros_Con_Campaña\", Hoja \"COMBO\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }
                    }


                    ListMecanicaCombo = ListMecanicaCombo
                           .Where(n => 
                                       !string.IsNullOrEmpty(n.Submarca.ToString()) &&
                                       !string.IsNullOrEmpty(n.Tipo_Marca.ToString()) &&
                                       !string.IsNullOrEmpty(n.SKU.ToString()) &&
                                       !string.IsNullOrEmpty(n.Descripcion.ToString()) &&
                                       !string.IsNullOrEmpty(n.Tipo.ToString()) &&
                                       n.Grupo_Combo > 0 &&
                                       (n.NumeroPadre > 0 ||
                                       n.NumeroHijo > 0) &&
                                       !string.IsNullOrEmpty(n.Alcance.ToString()) &&
                                       !string.IsNullOrEmpty(n.Capacidad.ToString()) &&
                                       !string.IsNullOrEmpty(n.Dinamica.ToString()) &&
                                       //n.VLitrosAnioAnt == 0 &&
                                       //n.PLitrosSinCamp == 0 &&
                                       n.PLitrosConCamp > 0
                           ).ToList();

                    #endregion
                }
                catch(Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"COMBO\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja COMBO, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }

                try
                {
                    //CLIENTES
                    #region CLIENTES

                    ListTienda = book.Worksheet("CLIENTES").AsEnumerable()
                                   .Select(row => new Tienda
                                   {
                                       //IdTienda = row["Num"].Cast<int>(),
                                       BillTo = row["Bill_To"],
                                       CustomerName = row["Customer_Name"],
                                       //Region = row["Region"],
                                       //DescripcionRegion = row["Descripcion_Region"],
                                       //DescripcionZona = row["Descripcion_Zona"],
                                       //ClaveSobreprecio = row["Clave_Sobreprecio"],
                                       Segmento = row["Segmento"],
                                       SubCanal = row["SubCanal"]//,
                                       //Exclusion = row["Exclusion"]
                                   }).ToList();

                    ListTienda.RemoveAll(n => //n.IdTienda <= 0 &&
                                      string.IsNullOrEmpty(n.BillTo.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.CustomerName.ToString().Trim()) &&
                                      //string.IsNullOrEmpty(n.Region.ToString()) &&
                                      //string.IsNullOrEmpty(n.DescripcionRegion.ToString()) &&
                                      //string.IsNullOrEmpty(n.DescripcionZona.ToString()) &&
                                      //string.IsNullOrEmpty(n.ClaveSobreprecio.ToString())
                                      string.IsNullOrEmpty(n.Segmento.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.SubCanal.ToString().Trim()) //&&
                                      //string.IsNullOrEmpty(n.Exclusion.ToString())
                                      );

                    if (ListTienda
                          .Where(n => //n.IdTienda <= 0 ||
                                      string.IsNullOrEmpty(n.BillTo.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.CustomerName.ToString().Trim()) ||
                                      //string.IsNullOrEmpty(n.Region.ToString()) ||
                                      //string.IsNullOrEmpty(n.DescripcionRegion.ToString()) ||
                                      //string.IsNullOrEmpty(n.DescripcionZona.ToString()) ||
                                      //string.IsNullOrEmpty(n.ClaveSobreprecio.ToString()) //&&
                                      string.IsNullOrEmpty(n.Segmento.ToString().Trim()) ||                  
                                      string.IsNullOrEmpty(n.SubCanal.ToString().Trim()) //||
                                      //string.IsNullOrEmpty(n.Exclusion.ToString())
                          ).Count() != 1)
                    {
                        ////ID_Tienda
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => n.IdTienda <= 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"ID_Tienda\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"ID_Tienda\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Bill_To
                        if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.BillTo.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Bill_To\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Bill_To\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Customer_Name
                        if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.CustomerName.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Customer_Name\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Customer_Name\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Region
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.Region.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Region\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Region\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Descripcion_Region
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.DescripcionRegion.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Descripcion_Region\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Descripcion_Region\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Descripcion_Zona
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.DescripcionZona.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Descripcion_Zona\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Descripcion_Zona\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Clave_Sobreprecio
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.ClaveSobreprecio.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Clave_Sobreprecio\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Clave_Sobreprecio\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Subcanal
                        if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.SubCanal.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Subcanal\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Subcanal\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Segmento
                        if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.Segmento.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Segmento\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Segmento\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }
                       
                        ////Exclusion
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.Exclusion.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Exclusion\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Exclusion\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}
                    }

                    ListTienda = ListTienda
                          .Where(n => //n.IdTienda > 0 &&
                                      !string.IsNullOrEmpty(n.BillTo.ToString()) &&
                                      !string.IsNullOrEmpty(n.CustomerName.ToString()) &&
                                      //!string.IsNullOrEmpty(n.Region.ToString()) &&
                                      //!string.IsNullOrEmpty(n.DescripcionRegion.ToString()) &&
                                      //!string.IsNullOrEmpty(n.DescripcionZona.ToString()) &&
                                      //!string.IsNullOrEmpty(n.ClaveSobreprecio.ToString())
                                      !string.IsNullOrEmpty(n.Segmento.ToString()) &&
                                      !string.IsNullOrEmpty(n.SubCanal.ToString()) //&&
                                      //!string.IsNullOrEmpty(n.Exclusion.ToString()) 
                          ).ToList();

                    idTienda = 0;

                    ListTienda.ForEach(n =>
                    {
                        n.IdTienda = ++idTienda;
                        n.Exclusion = (idTienda + 1).ToString();
                    });

                    if(ListTienda.Count == 0)
                    {
                        ListTienda.Add(new Tienda());
                    }

                    #endregion
                }
                catch(Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"CLIENTES\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }

                try
                {
                    //EXCLUSION
                    #region EXCLUSION

                    ListTiendaExclusion = book.Worksheet("EXCLUSIONES").AsEnumerable()
                                   .Select(row => new Tienda
                                   {
                                       //IdTienda = row["ID_Tienda"].Cast<int>(),
                                       BillTo = row["Clave"],
                                       CustomerName = row["Descripcion/Nombre"],                                      
                                       SubCanal = row["SubCanal"],
                                       Segmento = row["Segmento"],
                                       //Exclusion = row["Exclusion"],
                                       Region = row["Tipo"],
                                       //DescripcionRegion = row["Descripcion_Region"],
                                       //DescripcionZona = "Exclusion"//,
                                       //ClaveSobreprecio = row["Clave_Sobreprecio"],
                                   }).ToList();

                    ListTiendaExclusion.RemoveAll(n => //n.IdTienda <= 0 &&
                                      string.IsNullOrEmpty(n.BillTo.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.CustomerName.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.SubCanal.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.Segmento.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.Region.ToString().Trim())
                                      //string.IsNullOrEmpty(n.DescripcionRegion.ToString()) &&
                                      //string.IsNullOrEmpty(n.DescripcionZona.ToString()) &&
                                      //string.IsNullOrEmpty(n.ClaveSobreprecio.ToString())
                                      );

                    if (ListTiendaExclusion
                          .Where(n => //n.IdTienda <= 0 ||
                                      string.IsNullOrEmpty(n.BillTo.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.CustomerName.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.SubCanal.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.Segmento.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.Region.ToString().Trim()) //||
                                      //string.IsNullOrEmpty(n.ClaveSobreprecio.ToString()) //&&
                                                                                          //string.IsNullOrEmpty(n.Segmento.ToString()) &&                  
                                                                                          //string.IsNullOrEmpty(n.SubCanal.ToString()) &&
                                                                                          //string.IsNullOrEmpty(n.Exclusion.ToString())
                          ).Count() != 1)
                    {
                        ////ID_Tienda
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => n.IdTienda <= 0).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"ID_Tienda\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"ID_Tienda\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Bill_To/Clave
                        if (ListTiendaExclusion.Count > 0 && ListTiendaExclusion.Where(n => string.IsNullOrEmpty(n.BillTo.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Clave\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Clave\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Customer_Name
                        if (ListTiendaExclusion.Count > 0 && ListTiendaExclusion.Where(n => string.IsNullOrEmpty(n.CustomerName.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Customer_Name\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Customer_Name\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Exclusion
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.Exclusion.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Exclusion\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Exclusion\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //Subcanal
                        if (ListTiendaExclusion.Count > 0 && ListTiendaExclusion.Where(n => string.IsNullOrEmpty(n.SubCanal.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Subcanal\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Subcanal\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Segmento
                        if (ListTiendaExclusion.Count > 0 && ListTiendaExclusion.Where(n => string.IsNullOrEmpty(n.Segmento.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Segmento\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Segmento\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Tipo
                        if (ListTiendaExclusion.Count > 0 && ListTiendaExclusion.Where(n => string.IsNullOrEmpty(n.Region.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Tipo\", Hoja \"CLIENTES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Tipo\", Hoja \"CLIENTES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        ////Descripcion_Region
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.DescripcionRegion.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Descripcion_Region\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Descripcion_Region\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        ////Descripcion_Zona
                        //if (ListTienda.Count > 0 && ListTienda.Where(n => string.IsNullOrEmpty(n.DescripcionZona.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Descripcion_Zona\", Hoja \"CLIENTES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + "Carga Hoja \"CLIENTES\"" + ", Message:" + "falta informacion en la columna \"Descripcion_Zona\", Hoja \"CLIENTES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}
                    }


                    ListTiendaExclusion = ListTiendaExclusion
                          .Where(n => //n.IdTienda > 0 &&
                                      !string.IsNullOrEmpty(n.BillTo.ToString()) &&
                                      !string.IsNullOrEmpty(n.CustomerName.ToString()) &&
                                      !string.IsNullOrEmpty(n.SubCanal.ToString()) &&
                                      !string.IsNullOrEmpty(n.Segmento.ToString()) &&
                                      !string.IsNullOrEmpty(n.Region.ToString()) //&&
                                      //!string.IsNullOrEmpty(n.ClaveSobreprecio.ToString())
                          ).ToList();

                    idTienda = 0;

                    ListTiendaExclusion.ForEach(n =>
                    {
                        n.IdTienda = ++idTienda;
                        n.Exclusion = (-1).ToString();
                        n.DescripcionZona = "Exclusion";
                    });

                    if (ListTiendaExclusion.Count == 0)
                    {
                        ListTiendaExclusion.Add(new Tienda());
                    }

                    #endregion
                }
                catch (Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"CLIENTES\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja CLIENTES, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }

                try
                {
                    //ALCANCES
                    #region ALCANCES                 

                    ListAlcance = book.Worksheet("ALCANCES").AsEnumerable()
                                    .Select(row => new Alcance
                                    {
                                        //IdCampana = row["Bill_To"],
                                        //IdAlcance = row["Bill_To"],                                       
                                        ProductoEstelar = row["Producto_Estelar"],
                                        Observaciones = row["Alcance"],
                                        MecanicaPromocional = row["Mecanica_Promocional"],                                     
                                        Factor = row["Factor"],
                                        UnidadNegocio = row["Segmento"],
                                    }).ToList();

                    ListAlcance.RemoveAll(n => 
                                      string.IsNullOrEmpty(n.ProductoEstelar.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.Observaciones.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.MecanicaPromocional.ToString().Trim()) &&                                     
                                      //string.IsNullOrEmpty(n.Factor.ToString().Trim()) &&
                                      string.IsNullOrEmpty(n.UnidadNegocio.ToString().Trim())
                                      );

                    if (ListAlcance
                          .Where(n => 
                                      string.IsNullOrEmpty(n.ProductoEstelar.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.Observaciones.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.MecanicaPromocional.ToString().Trim()) ||                                      
                                      //string.IsNullOrEmpty(n.Factor.ToString().Trim()) ||
                                      string.IsNullOrEmpty(n.UnidadNegocio.ToString().Trim())
                          ).Count() != 1)
                    {
                        

                        //Producto_Estelar
                        if (ListAlcance.Count > 0 && ListAlcance.Where(n => string.IsNullOrEmpty(n.ProductoEstelar.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Producto_Estelar\", Hoja \"ALCANCES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja ALCANCES, Source:" + "Carga Hoja \"ALCANCES\"" + ", Message:" + "falta informacion en la columna \"Producto_Estelar\", Hoja \"ALCANCES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Alcance
                        if (ListAlcance.Count > 0 && ListAlcance.Where(n => string.IsNullOrEmpty(n.Observaciones.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Alcance\", Hoja \"ALCANCES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja ALCANCES, Source:" + "Carga Hoja \"ALCANCES\"" + ", Message:" + "falta informacion en la columna \"Alcance\", Hoja \"ALCANCES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }

                        //Mecanica_Promocional
                        if (ListAlcance.Count > 0 && ListAlcance.Where(n => string.IsNullOrEmpty(n.MecanicaPromocional.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Mecanica\", Hoja \"ALCANCES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja ALCANCES, Source:" + "Carga Hoja \"ALCANCES\"" + ", Message:" + "falta informacion en la columna \"Mecanica\", Hoja \"ALCANCES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }                     

                        ////Factor
                        //if (ListAlcance.Count > 0 && ListAlcance.Where(n => string.IsNullOrEmpty(n.Observaciones.ToString())).Count() > 0)
                        //{
                        //    Mensaje = "ERROR: Columna \"Factor\", Hoja \"ALCANCES\".";

                        //    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja ALCANCES, Source:" + "Carga Hoja \"ALCANCES\"" + ", Message:" + "falta informacion en la columna \"Factor\", Hoja \"ALCANCES\"");

                        //    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                        //    return productoLineaENTRes;
                        //}

                        //UnidadNegocio
                        if (ListAlcance.Count > 0 && ListAlcance.Where(n => string.IsNullOrEmpty(n.UnidadNegocio.ToString())).Count() > 0)
                        {
                            Mensaje = "ERROR: Columna \"Segmento\", Hoja \"ALCANCES\".";

                            productoLineaENTRes.ListMensaje.Add(Mensaje);

                            ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja ALCANCES, Source:" + "Carga Hoja \"ALCANCES\"" + ", Message:" + "falta informacion en la columna \"Segmento\", Hoja \"ALCANCES\"");

                            productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                            //return productoLineaENTRes;
                        }
                    }


                    ListAlcance = ListAlcance
                          .Where(n => 
                                      !string.IsNullOrEmpty(n.ProductoEstelar.ToString()) &&
                                      !string.IsNullOrEmpty(n.Observaciones.ToString()) &&
                                      !string.IsNullOrEmpty(n.MecanicaPromocional.ToString()) &&                                     
                                      //!string.IsNullOrEmpty(n.Factor.ToString()) &&
                                      !string.IsNullOrEmpty(n.UnidadNegocio.ToString())
                          ).ToList();

                    #endregion
                }
                catch (Exception ex)
                {
                    Mensaje = "ERROR: No se guardo la informacion de Campaña Escenarios, el archivo no tiene el formato correcto, Hoja \"ALCANCES\".";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana: Cargar Archivo Excel Hoja ALCANCES, Source:" + ex.Source + ", Message:" + ex.Message);

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();

                    //return productoLineaENTRes;
                }


                //SE VALIDA QUE NO TENGA ERRORES
                if(productoLineaENTRes.ListMensaje.Count > 0)
                {
                    productoLineaENTRes.Mensaje = "ERROR";

                    return productoLineaENTRes;
                }

                CampanaDAT campanaDAT = new CampanaDAT();

                ListLineaFamilia = campanaDAT.GuardarProdutoCampanaCompleto(ClaveCampana, ListMecanicaRegalo, ListMecanicaMultiplo, ListMecanicaDescuento,
                                                                        ListMecanicaVolumen, ListMecanicaKit, ListMecanicaCombo, ListTienda, ListTiendaExclusion, 
                                                                        ListAlcance);

                if (ListLineaFamilia != null)
                {
                    productoLineaENTRes.Mensaje = "OK";
                    productoLineaENTRes.ListLineaFamilia = ListLineaFamilia;
                }
                else
                {
                    Mensaje = "ERROR: No se encontro informacion de los sobrepecios, intente nuevamente o consulte al administrador de sistema.";

                    productoLineaENTRes.ListMensaje.Add(Mensaje);

                    productoLineaENTRes.Mensaje = Mensaje;

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana, Message: No se encontro informacion de los sobrepecios, intente nuevamente o consulte al administrador de sistema.");

                    productoLineaENTRes.ListLineaFamilia = new List<LineaFamilia>();
                }


                //if (respuesta == 1)
                //{
                //    Mensaje = "OK";
                //}
                //else
                //{
                //    Mensaje = "ERROR: Ocurrio un problema inesperado, identifique si se guardo correctamente la informacion, o consulte al administrador de sistemas.";
                //}
            }
            catch (Exception ex)
            {
                Mensaje = "ERROR: Ocurrio un error inesperado, no se guardo la informacion de Campaña Escenarios, intente nuevamente o consulte a su administrador de sistemas.";

                productoLineaENTRes.ListMensaje.Add(Mensaje);

                productoLineaENTRes.Mensaje = Mensaje;

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - GuardarProductoCampana, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return productoLineaENTRes;
        }
        public RentabilidadENT MostrarRentabilidad(RentabilidadENT rentabilidadENTReq)
        {
            RentabilidadENT rentabilidadENTRes = new RentabilidadENT();
            Rentabilidad rentabilidad = new Rentabilidad();

            List<ReporteCEO> ListReporteCEO = new List<ReporteCEO>();
            List<ReporteCEOMecanica> ListReporteCEOMecanica = new List<ReporteCEOMecanica>();
            List<ReporteCEOPublicidad> ListReporteCEOPublicidad = new List<ReporteCEOPublicidad>();
            List<ReporteMKT> ListReporteMKT = new List<ReporteMKT>();
            List<ReporteMKT> ListReporteMKTOrden = new List<ReporteMKT>();
            List<ReporteSKU> ListReporteSKU = new List<ReporteSKU>();
            List<EntidadesCampanasPPG.Modelo.Campana> ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();
            List<SKUPrecioCosto> ListSKUPrecioCosto = new List<SKUPrecioCosto>();

            StringBuilder html = new StringBuilder();
            List<string> ListHtmlCEO = new List<string>();
            List<string> ListHtmlMercadotecnia = new List<string>();
            List<string> ListHtmlSKU = new List<string>();

            int RegPrimerPaginaCEOMerca = 40;
            int RegPrimerPaginaCEOPubli = 40;
            int contRegCEO = 0;

            bool primerPaginaMercadotecnia = true;
            int RegPrimerPaginaMercadotecnia = 30;
            int RegSegundoPaginaMercadotecnia = 40;
            int contRegMercadotecnia = 0;

            bool primerPaginaSKU= true;
            int RegPrimerPaginaSKU = 35;
            int RegSegundoPaginaSKU = 40;
            int contRegSKU = 0;

            StringBuilder inicioHtml = new StringBuilder();
            StringBuilder finHtml = new StringBuilder();
            StringBuilder inicioDatosHtml = new StringBuilder();
            StringBuilder finDatosHtml = new StringBuilder();
            StringBuilder tablaInicioHtml = new StringBuilder();

            //DataSet dsReporteCEO = new DataSet();
            DataTable dtReporteCEO = new DataTable();
            DataTable dtReporteCEOMecanica = new DataTable();
            DataTable dtReporteCEOPublicidad = new DataTable();

            //DataSet dsReporteMKT = new DataSet();
            DataTable dtReporteMKT = new DataTable();

            //DataSet dsReporteSKU = new DataSet();
            DataTable dtReporteSKU = new DataTable();

            string nombreArchivo = string.Empty;
            string pathArchivo = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;

            string pathArchivoCompletoCEO = string.Empty;
            string pathArchivoCompletoMercadotecnia = string.Empty;
            string pathArchivoCompletoSKU = string.Empty;

            string pathArchivoCompletoCEOEx = string.Empty;
            string pathArchivoCompletoMercadotecniaEx = string.Empty;
            string pathArchivoCompletoSKUEx = string.Empty;

            string url = string.Empty;

            string urlCompletoCEO = string.Empty;
            string urlCompletoMercadotecnia = string.Empty;
            string urlCompletoSKU= string.Empty;

            string urlCompletoCEOEx = string.Empty;
            string urlCompletoMercadotecniaEx = string.Empty;
            string urlCompletoSKUEx = string.Empty;

            string nombreReporte = string.Empty;

            double widthPDF = 0;
            double heigthPDF = 0;
            double widthPDFVertical = 0;
            double heigthPDFVertical = 0;

            DataTable dtParametro = new DataTable();
            List<Parametro> ListParametro = new List<Parametro>();
            ParametroDAT parametroDAT = new ParametroDAT();
            Parametro parametro = new Parametro();

            string pathImgComex = string.Empty;
            string pathImgPPG = string.Empty;
            string pathImgHeader = string.Empty;
            string pathImgFooter = string.Empty;

            string claveCampana = string.Empty;
            string nombreCampana = string.Empty;
            string fechaInicioSubCanal = string.Empty;
            string fechaFinSubCanal = string.Empty;
            string fechaInicioPublico = string.Empty;
            string fechaFinPublico = string.Empty;

            //bool columnStyle = true;
            string columnStyleHtml = string.Empty;
            const string etiquetaSubtotal = "SUBTOTAL";

            const string etiquetaRentable = "RENTABLE";
            const string etiqueteNoRentable = "NO RENTABLE";
            const string etiquetaRevisar = "REVISAR";

            EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();

            List<string> ListAlcance = new List<string>();
            List<int> ListPeriodo = new List<int>();
            List<ArticuloCantidad> ListArticuloCantidad = new List<ArticuloCantidad>();
            List<string> ListDatosCeoMecanica = new List<string>();
            ReporteCEOMecanica ReporteCEOMecanicaTemp = new ReporteCEOMecanica();

            List<int> ListaRegBlancoAnalisis = new List<int>() { 0, 1, 2, 4, 5, 7, 9, 11, 12, 13, 14, 17, 19, 20, 21, 23, 24 };
            List<int> ListaRegAzulAnalisis = new List<int>() { 3, 6, 8, 10, 15, 16, 18, 22, 25 };

            int RegNoSell_1 = 0;
            int RegNoSell_2 = 1;
            int RegSellIn = 2;
            int NumRegSellIn = 20;
            int RegSellOut = 22;
            int NumRegSellOut = 4;

            bool PrimerSellIn = false;
            bool PrimerSellOut = false;

            List<PeriodoCEOMecanica> ListPeriodoAntAct = new List<PeriodoCEOMecanica>();
            PeriodoCEOMecanica periodoMecanica = new PeriodoCEOMecanica();
            bool periodoStyle = true;

            decimal totalPublicidadPAnt = 0;
            decimal totalPublicidadPAct = 0;

            //bool autoDetectPage = true;

            string htmlReporteCEO = string.Empty;
            string htmlReporteMercadotecnia = string.Empty;
            string htmlReporteSKU = string.Empty;

            PdfHtmlLayoutFormat layoutCEO = new PdfHtmlLayoutFormat();
            PdfPageSettings pageCEO = new PdfPageSettings();
            PdfDocument docCEO = new PdfDocument();
            PdfDocument docCEOImg = new PdfDocument();

            PdfHtmlLayoutFormat layoutMercadotecnia = new PdfHtmlLayoutFormat();
            PdfPageSettings pageMercadotecnia = new PdfPageSettings();
            PdfDocument docMercadotecnia = new PdfDocument();
            PdfDocument docMercadotecniaImg = new PdfDocument();

            PdfHtmlLayoutFormat layoutSKU = new PdfHtmlLayoutFormat();
            PdfPageSettings pageSKU = new PdfPageSettings();
            PdfDocument docSKU = new PdfDocument();
            PdfDocument docSKUImg = new PdfDocument();

            //PdfImage headerImage;
            //PdfImage footerImage;

            int contRegPDF = 0;
            //int maxRegPDF = 100;

            int j = 0;

            try
            {
                rentabilidad = rentabilidadENTReq.Rentabilidad;

                CampanaDAT campanaDAT = new CampanaDAT();

                rentabilidad = campanaDAT.MostrarRentabilidad(rentabilidad);

                if (rentabilidad.Mensaje == "OK")
                {
                    //DATOS DE RENTABILIDAD
                    rentabilidadENTRes.Rentabilidad = rentabilidad;

                    ListReporteCEO = rentabilidad.ListReporteCEO;
                    ListReporteCEOMecanica = rentabilidad.ListReporteCEOMecanica;
                    ListReporteCEOPublicidad = rentabilidad.ListReporteCEOPublicidad;
                    ListReporteMKT = rentabilidad.ListReporteMKT;
                    ListReporteSKU = rentabilidad.ListReporteSKU;
                    ListCampana = rentabilidad.ListCampania;
                    ListSKUPrecioCosto = rentabilidad.ListSKUPrecioCosto;
                    

                    #region DATOS CAMPAÑA

                    if(ListCampana.Count > 0)
                    {
                        campana = ListCampana.FirstOrDefault();

                        claveCampana = campana.Title;
                        nombreCampana = campana.NombreCampa;
                        fechaInicioSubCanal = campana.FechaInicioSubCanal;
                        fechaFinSubCanal = campana.FechaFinSubCanal;
                        fechaInicioPublico = campana.FechaInicioPublico;
                        fechaFinPublico = campana.FechaFinPublico;
                    }

                    #endregion


                    #region REPORTE CEO

                    ListHtmlCEO = new List<string>();

                    //REPORTE CEO
                    if (ListReporteCEO.Count > 0)
                    {
                        html = new StringBuilder();
                        inicioHtml = new StringBuilder();
                        inicioDatosHtml = new StringBuilder();
                        finDatosHtml = new StringBuilder();
                        finHtml = new StringBuilder();
                        ListHtmlCEO = new List<string>();

                        //OBTENER PARAMETROS
                        dtParametro = parametroDAT.GetParametro(0, null);

                        ListParametro = dtParametro.AsEnumerable()
                                        .Select(n => new Parametro
                                        {
                                            Id = n.Field<int?>("Id").GetValueOrDefault(),
                                            Nombre = n.Field<string>("Nombre"),
                                            Valor = n.Field<string>("Valor")
                                        }).ToList();

                        //PDF
                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgComex"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgComex = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgPPG"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgPPG = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgHeader"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgHeader = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgFooter"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgFooter = parametro.Valor;
                        }                   

                        //var pngBinaryDataComex = File.ReadAllBytes(pathImgComex);
                        //var ImgDataURIComex = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataComex);
                        //var ImgHtmlComex = String.Format("<img src='{0}' style='float:right;vertical-align:top;' width='120' height='50'>", ImgDataURIComex);

                        //var pngBinaryDataPPG = File.ReadAllBytes(pathImgPPG);
                        //var ImgDataURIPPG = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataPPG);
                        var pngBinaryDataHeaer = File.ReadAllBytes(pathImgHeader);
                        var ImgDataURIHeader = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataHeaer);
                        var pngBinaryDataFooter = File.ReadAllBytes(pathImgFooter);
                        var ImgDataURIFooter = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataFooter);

                        //var ImgHtmlPPG = String.Format("<img src='{0}' style='float:left;vertical-align:top;' width='70px' height='50px'>", ImgDataURIPPG);
                        //var ImgHtmlPPGFondo = String.Format("<div style='height: 500px; position: relative; background-repeat: no-repeat; background-size: 100% 33%; background-position: center; background-image: url(\"{0}\")';><div style='font-family:arial;'>", ImgDataURIPPG);
                        //var ImgHtmlHeader = String.Format("<div style='width: 100%; height: 50px; position: absolute; top: 1px;'><img src='{0}' width='100%' height='100%'></div>", ImgDataURIHeader);
                        //var ImgHtmlFooter = String.Format("<footer style='width: 100%; height: 50px; flow: static(footer);'><img src='{0}' width='100%' height='100%'></footer>", ImgDataURIFooter);

                        //html.AppendLine(ImgHtmlComex);
                        //html.AppendLine("<div style='font-family: arial; position: relative; height: 100%;'>");
                        //html.AppendLine(ImgHtmlHeader);
                        //html.AppendLine(ImgHtmlPPG);   
                        html.AppendLine("<div style='font-family: arial; position: relative;'>");
                        html.AppendLine("<center><h3>Reporte CEO</h3></center>");
                        html.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                        html.AppendLine("<tbody>");
                        //REGISTRO 1
                        html.AppendLine("<tr style='background-color:#ffffff'>");
                        html.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Campaña</td>");
                        html.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + nombreCampana + "</td>");
                        html.AppendLine("</tr>");
                        //REGISTRO 2
                        html.AppendLine("<tr style='background-color:#CEE3F6;'>");
                        html.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Clave Campaña</td>");
                        html.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + claveCampana + "</td>");
                        html.AppendLine("</tr>");
                        //REGISTRO 3
                        if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal.ToString() != "01/01/0001")
                        {
                            html.AppendLine("<tr style='background-color:#ffffff'>");
                            html.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Inicio SubCanal</td>");
                            html.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + fechaInicioSubCanal + "</td>");
                            html.AppendLine("</tr>");
                        }
                        //REGISTRO 4
                        if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal.ToString() != "01/01/0001")
                        {
                            html.AppendLine("<tr style='background-color:#CEE3F6'>");
                            html.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Fin SubCanal</td>");
                            html.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + fechaFinSubCanal + "</td>");
                            html.AppendLine("</tr>");
                        }
                        //REGISTRO 5
                        if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico.ToString() != "01/01/0001")
                        {
                            html.AppendLine("<tr style='background-color:#ffffff'>");
                            html.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Inicio Publico</td>");
                            html.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + fechaInicioPublico + "</td>");
                            html.AppendLine("</tr>");
                        }
                        //REGISTRO 6
                        if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico.ToString() != "01/01/0001")
                        {
                            html.AppendLine("<tr style='background-color:#CEE3F6'>");
                            html.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Fin Publico</td>");
                            html.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + fechaFinPublico + "</td>");
                            html.AppendLine("</tr>");
                        }

                        html.AppendLine("</tbody>");
                        html.AppendLine("</table>");

                        html.AppendLine("</div>");
                        //html.AppendLine("<br />");
                        //html.AppendLine("<br />");


                        #region REPORTE CEO ANALISIS

                        /////////////////////////////////
                        //PAGINA UNO REPORTE CEO ANALISIS

                        html.AppendLine("<div style='font-family: arial; position: relative;'>");
                        html.AppendLine("<center><h3>1) Analisis</h3></center>");
                        html.AppendLine("<center>");
                        html.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                        html.AppendLine("<thead style='background-color:#3B89D8; font-size:13px; color:white;'>");
                        html.AppendLine("<tr>");

                        ////AGREGAR COLUMNAS
                        //html.AppendLine("<th>NOMBRE CAMPANIA</th>");
                        //html.AppendLine("<th>PERIODO</th>");
                        //html.AppendLine("<th>SUBTOTAL LITROS</th>");
                        //html.AppendLine("<th>SUBTOTAL PIEZAS</th>");
                        //html.AppendLine("<th>TOTAL LITROS / PIEZAS</th>");
                        //html.AppendLine("<th>IMPORTE</th>");
                        //html.AppendLine("<th>IMPORTE LITROS</th>");
                        //html.AppendLine("<th>IMPORTE PZAS</th>");
                        //html.AppendLine("<th>PRECIO PROMEDIO X LT</th>");
                        //html.AppendLine("<th>COSTO MATERIA PRIMA + ENVASE + FABRICACION</th>");
                        //html.AppendLine("<th>UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION</th>");
                        //html.AppendLine("<th>FACTOR UTILIDAD</th>");
                        ////html.AppendLine("<th>MATERIA PRIMA + ENVASE + FABRICACION</th>");
                        //html.AppendLine("<th>% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACIÓN</th>");
                        //html.AppendLine("<th>INVERSION PUBLICITARIA</th>");
                        //html.AppendLine("<th>OTROS GASTOS DE OPERACIÓN PLANTA</th>");
                        //html.AppendLine("<th>GASTOS DE OPERACIÓN GENERALES KROMA</th>");
                        //html.AppendLine("<th>NOTAS DE CRÉDITO</th>");
                        //html.AppendLine("<th>UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION GASTOS DE OPERACIÓN GENERALES KROMA</th>");
                        //html.AppendLine("<th>UTILIDAD POR LT/PZ</th>");
                        //html.AppendLine("<th>FACTOR DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION</th>");
                        //html.AppendLine("<th>% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS</th>");
                        //html.AppendLine("<th>% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR</th>");
                        //html.AppendLine("<th>% INCREMENTO EN LITROS Vs PRESUPUESTO</th>");
                        //html.AppendLine("<th>ROI</th>");

                        html.AppendLine("<th style='width:25px;'></th>");
                        html.AppendLine("<th style='text-align:left; width:235px; background-color:#3B89D8; border: 0.5px solid black;'>Descripcion</th>");
                        html.AppendLine("<th style='text-align:left; width:95px; background-color:#2E8B57; border: 0.5px solid black;'>Periodo Anterior</th>");
                        html.AppendLine("<th style='width:95px; background-color:#3B89D8; border: 0.5px solid black;'>Periodo Actual Sin Campaña</th>");
                        html.AppendLine("<th style='width:95px; background-color:#3B89D8; border: 0.5px solid black;'>Periodo Actual Con Campaña</th>");
                        //html.AppendLine("<th style='width:70px; background-color:#3B89D8; border: 0.5px solid black;'>Periodo Actual Con Campaña Necesario</th>");
                        //html.AppendLine("<th style='width:70px; background-color:#3B89D8; border: 0.5px solid black;'>Ventas Real</th>");

                        //FIN COLUMNAS
                        html.AppendLine("</tr>");
                        html.AppendLine("</thead>");

                        ////EXCEL
                        //dtReporteCEO.Columns.Add("NOMBRE CAMPANIA");
                        //dtReporteCEO.Columns.Add("PERIODO");
                        //dtReporteCEO.Columns.Add("SUBTOTAL LITROS");
                        //dtReporteCEO.Columns.Add("SUBTOTAL PIEZAS");
                        //dtReporteCEO.Columns.Add("TOTAL LITROS / PIEZAS");
                        //dtReporteCEO.Columns.Add("IMPORTE");
                        //dtReporteCEO.Columns.Add("IMPORTE LITROS");
                        //dtReporteCEO.Columns.Add("IMPORTE PZAS");
                        //dtReporteCEO.Columns.Add("PRECIO PROMEDIO X LT");
                        //dtReporteCEO.Columns.Add("COSTO MATERIA PRIMA + ENVASE + FABRICACION");
                        //dtReporteCEO.Columns.Add("UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION");
                        //dtReporteCEO.Columns.Add("FACTOR UTILIDAD");
                        ////dtReporteCEO.Columns.Add("MATERIA PRIMA + ENVASE + FABRICACION");
                        //dtReporteCEO.Columns.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACIÓN");
                        //dtReporteCEO.Columns.Add("INVERSION PUBLICITARIA");
                        //dtReporteCEO.Columns.Add("OTROS GASTOS DE OPERACIÓN PLANTA");
                        //dtReporteCEO.Columns.Add("GASTOS DE OPERACIÓN GENERALES KROMA");
                        //dtReporteCEO.Columns.Add("NOTAS DE CRÉDITO");
                        //dtReporteCEO.Columns.Add("UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION GASTOS DE OPERACIÓN GENERALES KROMA");
                        //dtReporteCEO.Columns.Add("UTILIDAD POR LT/PZ");
                        //dtReporteCEO.Columns.Add("FACTOR DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION");
                        //dtReporteCEO.Columns.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS");
                        //dtReporteCEO.Columns.Add("% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR");
                        //dtReporteCEO.Columns.Add("% INCREMENTO EN LITROS Vs PRESUPUESTO");
                        //dtReporteCEO.Columns.Add("ROI");

                        dtReporteCEO.Columns.Add("Descripcion");
                        dtReporteCEO.Columns.Add("Periodo Anterior");
                        dtReporteCEO.Columns.Add("Periodo Actual Sin Campaña");
                        dtReporteCEO.Columns.Add("Periodo Actual Con Campaña");
                        //dtReporteCEO.Columns.Add("Periodo Actual Con Campaña Necesario");
                        //dtReporteCEO.Columns.Add("Ventas Real");

                        //AGREGAR COLUMNA
                        for (int i = 0; i < ListReporteCEO.Count - 2 && (i + 1) < dtReporteCEO.Columns.Count; i++)
                        {
                            if (i == 0)
                            {
                                dtReporteCEO.Rows.Add("SUBTOTAL LITROS", ListReporteCEO.ElementAt(i).SubtotalLitros.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("SUBTOTAL PIEZAS", ListReporteCEO.ElementAt(i).SubtotalPiezas.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("TOTAL LITROS / PIEZAS", ListReporteCEO.ElementAt(i).TotalLitrosPiezas.ToString("#,##0"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("IMPORTE", ListReporteCEO.ElementAt(i).Importe.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("IMPORTE LITROS", ListReporteCEO.ElementAt(i).ImporteLitros.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("IMPORTE PZAS", ListReporteCEO.ElementAt(i).ImportePiezas.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("PRECIO PROMEDIO X LT", ListReporteCEO.ElementAt(i).PrecioPromedioLitro.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("COSTO MATERIA PRIMA + ENVASE + FABRICACION", ListReporteCEO.ElementAt(i).CostoMP.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION", ListReporteCEO.ElementAt(i).UtilidadMP.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("FACTOR UTILIDAD MATERIA PRIMA + ENVASE + FABRICACION", ListReporteCEO.ElementAt(i).FactorUtilidad.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACIÓN", ListReporteCEO.ElementAt(i).MargenUtilidad.ToString("#,##0.00%"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("INVERSION PUBLICITARIA", ListReporteCEO.ElementAt(i).InversionPublicidad.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("OTROS GASTOS DE OPERACIÓN PLANTA", ListReporteCEO.ElementAt(i).OtrosGastos.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("GASTOS DE OPERACIÓN GENERALES KROMA", ListReporteCEO.ElementAt(i).GastosOperacion.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("NOTAS DE CRÉDITO", ListReporteCEO.ElementAt(i).NotasCredito.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS DE OPERACIÓN GENERALES KROMA + GASTOS DE OPERACIÓN PLANTA + NOTAS DE CRÉDITO", ListReporteCEO.ElementAt(i).UtilidadConsideraMP.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("UTILIDAD POR LT/PZ", ListReporteCEO.ElementAt(i).UtilidadLitroPieza.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("FACTOR DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS DE OPERACIÓN GENERALES KROMA + GASTOS DE OPERACIÓN PLANTA", ListReporteCEO.ElementAt(i).FactorUtilidadMP.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS DE OPERACIÓN GENERALES KROMA + GASTOS DE OPERACIÓN PLANTA", ListReporteCEO.ElementAt(i).PorcenUtilidadConsideraMP.ToString("#,##0.00%"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("% INCREMENTO EN LITROS Vs PRESUPUESTO", ListReporteCEO.ElementAt(i).PorcenIncrementoLitros.ToString("#,##0.00%"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR", ListReporteCEO.ElementAt(i).PorcenIncrementoLitrosPresu.ToString("#,##0.00%"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("ROI", ListReporteCEO.ElementAt(i).Roi.ToString("#,##0.00%"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("IMPORTE PRECIO PÚBLICO", ListReporteCEO.ElementAt(i).ImportePrecioPublico.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("IMPORTE PRECIO PÚBLICO SIN IVA", ListReporteCEO.ElementAt(i).ImportePrecioPublicoSinIva.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("UTILIDAD EN CONCESIONARIO", ListReporteCEO.ElementAt(i).UtilidadEnConcesionario.ToString("#,##0.00"), "", "");//, "", "");
                                dtReporteCEO.Rows.Add("MARGEN CONCESIONARIO", ListReporteCEO.ElementAt(i).MargenConcesionario.ToString("#,##0.00%"), "", "");//, "", "");
                            }
                            else
                            {
                                dtReporteCEO.Rows[0][i + 1] = ListReporteCEO.ElementAt(i).SubtotalLitros.ToString("#,##0.00");
                                dtReporteCEO.Rows[1][i + 1] = ListReporteCEO.ElementAt(i).SubtotalPiezas.ToString("#,##0.00");
                                dtReporteCEO.Rows[2][i + 1] = ListReporteCEO.ElementAt(i).TotalLitrosPiezas.ToString("#,##0");
                                dtReporteCEO.Rows[3][i + 1] = ListReporteCEO.ElementAt(i).Importe.ToString("#,##0.00");
                                dtReporteCEO.Rows[4][i + 1] = ListReporteCEO.ElementAt(i).ImporteLitros.ToString("#,##0.00");
                                dtReporteCEO.Rows[5][i + 1] = ListReporteCEO.ElementAt(i).ImportePiezas.ToString("#,##0.00");
                                dtReporteCEO.Rows[6][i + 1] = ListReporteCEO.ElementAt(i).PrecioPromedioLitro.ToString("#,##0.00");
                                dtReporteCEO.Rows[7][i + 1] = ListReporteCEO.ElementAt(i).CostoMP.ToString("#,##0.00");
                                dtReporteCEO.Rows[8][i + 1] = ListReporteCEO.ElementAt(i).UtilidadMP.ToString("#,##0.00");
                                dtReporteCEO.Rows[9][i + 1] = ListReporteCEO.ElementAt(i).FactorUtilidad.ToString("#,##0.00");
                                dtReporteCEO.Rows[10][i + 1] = ListReporteCEO.ElementAt(i).MargenUtilidad.ToString("#,##0.00%");
                                dtReporteCEO.Rows[11][i + 1] = ListReporteCEO.ElementAt(i).InversionPublicidad.ToString("#,##0.00");
                                dtReporteCEO.Rows[12][i + 1] = ListReporteCEO.ElementAt(i).OtrosGastos.ToString("#,##0.00");
                                dtReporteCEO.Rows[13][i + 1] = ListReporteCEO.ElementAt(i).GastosOperacion.ToString("#,##0.00");
                                dtReporteCEO.Rows[14][i + 1] = ListReporteCEO.ElementAt(i).NotasCredito.ToString("#,##0.00");
                                dtReporteCEO.Rows[15][i + 1] = ListReporteCEO.ElementAt(i).UtilidadConsideraMP.ToString("#,##0.00");
                                dtReporteCEO.Rows[16][i + 1] = ListReporteCEO.ElementAt(i).UtilidadLitroPieza.ToString("#,##0.00");
                                dtReporteCEO.Rows[17][i + 1] = ListReporteCEO.ElementAt(i).FactorUtilidadMP.ToString("#,##0.00");
                                dtReporteCEO.Rows[18][i + 1] = ListReporteCEO.ElementAt(i).PorcenUtilidadConsideraMP.ToString("#,##0.00%");
                                dtReporteCEO.Rows[19][i + 1] = ListReporteCEO.ElementAt(i).PorcenIncrementoLitros.ToString("#,##0.00%");
                                dtReporteCEO.Rows[20][i + 1] = ListReporteCEO.ElementAt(i).PorcenIncrementoLitrosPresu.ToString("#,##0.00%");
                                dtReporteCEO.Rows[21][i + 1] = ListReporteCEO.ElementAt(i).Roi.ToString("#,##0.00%");
                                dtReporteCEO.Rows[22][i + 1] = ListReporteCEO.ElementAt(i).ImportePrecioPublico.ToString("#,##0.00");
                                dtReporteCEO.Rows[23][i + 1] = ListReporteCEO.ElementAt(i).ImportePrecioPublicoSinIva.ToString("#,##0.00");
                                dtReporteCEO.Rows[24][i + 1] = ListReporteCEO.ElementAt(i).UtilidadEnConcesionario.ToString("#,##0.00");
                                dtReporteCEO.Rows[25][i + 1] = ListReporteCEO.ElementAt(i).MargenConcesionario.ToString("#,##0.00%");
                            }
                        }

                        //LLENAR DATOS REPORTE CEO
                        html.AppendLine("<tbody style='font-size:11.5px;'>");

                        //columnStyle = true;
                        j = 0;
                        PrimerSellIn = true;
                        PrimerSellOut = true;

                        foreach (DataRow dr in dtReporteCEO.Rows)
                        {
                            //if(columnStyle)
                            //{
                            //    columnStyleHtml = "style='background-color:#ffffff'";
                            //    columnStyle = !columnStyle;
                            //}
                            //else
                            //{
                            //    columnStyleHtml = "style='background-color:#CEE3F6'";
                            //    columnStyle = !columnStyle;
                            //}

                            if (ListaRegBlancoAnalisis.Where(n => n == j).Count() > 0)
                            {
                                columnStyleHtml = "background-color:#ffffff;";
                            }
                            else if (ListaRegAzulAnalisis.Where(n => n == j).Count() > 0)
                            {
                                columnStyleHtml = "background-color:#CEE3F6;";
                            }
                            else
                            {
                                columnStyleHtml = string.Empty;
                            }

                            //PDF
                            html.AppendLine("<tr style='" + columnStyleHtml + "'>");

                            if (j == RegNoSell_1)
                            {
                                html.AppendLine("<td style='" + columnStyleHtml + " border: 0.5px solid black;'></td>");
                            }
                            else if(j == RegNoSell_2)
                            {
                                html.AppendLine("<td style='" + columnStyleHtml + " border: 0.5px solid black;'></td>");
                            }
                            else if (j == RegSellIn)
                            {
                                if (PrimerSellIn)
                                {
                                    html.AppendLine("<td rowspan='" + NumRegSellIn + "' style='" + columnStyleHtml + " border: 0.5px solid black;'>SELL IN</td>");

                                    PrimerSellIn = false;
                                }
                                //else
                                //{
                                //    html.AppendLine("<td style='" + columnStyleHtml + " border: 0.5px solid black; transform: rotate(-90deg); display: inline-block;'></td>");
                                //}
                            }
                            else if (j == RegSellOut)
                            {
                                if (PrimerSellOut)
                                {
                                    html.AppendLine("<td rowspan='" + NumRegSellOut + "' style='" + columnStyleHtml + " border: 0.5px solid black;'>SELL OUT</td>");

                                    PrimerSellOut = false;
                                }
                                //else
                                //{
                                //    html.AppendLine("<td style='" + columnStyleHtml + " border: 0.5px solid black; transform: rotate(-90deg); display: inline-block;'></td>");
                                //}
                            }

                            html.AppendLine("<td style='" + columnStyleHtml + " text-align:left; border: 0.5px solid black;'>" + dr[0].ToString() + "</td>");
                            html.AppendLine("<td style='" + columnStyleHtml + " text-align:right; border: 0.5px solid black;'>" + dr[1].ToString() + "</td>");
                            html.AppendLine("<td style='" + columnStyleHtml + " text-align:right; border: 0.5px solid black;'>" + dr[2].ToString() + "</td>");
                            html.AppendLine("<td style='" + columnStyleHtml + " text-align:right; border: 0.5px solid black;'>" + dr[3].ToString() + "</td>");
                            //html.AppendLine("<td style='" + columnStyleHtml + " text-align: right; border: 0.5px solid black;'>" + dr[4].ToString() + "</td>");
                            //html.AppendLine("<td style='" + columnStyleHtml + " text-align: right; border: 0.5px solid black;'>" + dr[5].ToString() + "</td>");

                            html.AppendLine("</tr>");

                            ////PDF 2 PRUEBA
                            //html.AppendLine("<tr " + columnStyleHtml + ">");

                            //html.AppendLine("<td style='border: 0.5px solid black;'>" + dr[0].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[1].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[2].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[3].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[4].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[5].ToString() + "</td>");

                            //html.AppendLine("</tr>");

                            ////PDF 3 PRUEBA
                            //html.AppendLine("<tr " + columnStyleHtml + ">");

                            //html.AppendLine("<td style='border: 0.5px solid black;'>" + dr[0].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[1].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[2].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[3].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[4].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[5].ToString() + "</td>");

                            //html.AppendLine("</tr>");

                            ////PDF 4 PRUEBA
                            //html.AppendLine("<tr " + columnStyleHtml + ">");

                            //html.AppendLine("<td style='border: 0.5px solid black;'>" + dr[0].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[1].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[2].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[3].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[4].ToString() + "</td>");
                            //html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + dr[5].ToString() + "</td>");

                            //html.AppendLine("</tr>");

                            j++;
                        }

                        html.AppendLine("</tbody>");
                        html.AppendLine("</table>");
                        html.AppendLine("</center>");

                        //html.AppendLine("</div>");
                        //html.AppendLine(ImgHtmlFooter);
                        html.AppendLine("</div>");
                        //html.AppendLine("<br />");
                        //html.AppendLine("<br />");
                        //html.AppendLine("<br />");
                        //html.AppendLine("<br />");

                        /////////////////////
                        //AGREGAR PRIMERA PAGINA ANALISIS CEO
                        ListHtmlCEO.Add(html.ToString());

                        #endregion

                        #region REPORTE CEO MECANICA

                        /////////////////////////////////
                        //PAGINA DOS REPORTE CEO MECANICA

                        html = new StringBuilder();
                        inicioHtml = new StringBuilder();
                        inicioDatosHtml = new StringBuilder();
                        finDatosHtml = new StringBuilder();
                        finHtml = new StringBuilder();
                        contRegCEO = 0;

                        if (ListReporteCEOMecanica.Count > 0)
                        {
                            dtReporteCEOMecanica = new DataTable();
                            dtReporteCEOMecanica.Columns.Add("Articulo");
                            dtReporteCEOMecanica.Columns.Add("Cantidad");

                            ListPeriodo = ListReporteCEOMecanica.GroupBy(n => n.Periodo).Select(m => m.Key).OrderBy(s => s).ToList();

                            foreach (int periodo in ListPeriodo)
                            {
                                ListAlcance = ListReporteCEOMecanica.Where(i => i.Periodo == periodo).GroupBy(n => n.Alcance).Select(m => m.Key.ToString()).ToList();

                                foreach (string alcance in ListAlcance)
                                {
                                    dtReporteCEOMecanica.Columns.Add(periodo.ToString() + "-" + alcance);
                                    //dtReporteCEOMecanica.Columns.Add(alcance);
                                }
                            }

                            /////////////////////////
                            //LLENAR DATOS A LA TABLA
                            ListArticuloCantidad = ListReporteCEOMecanica.GroupBy(n => new
                            {
                                n.Articulo,
                                n.Cantidad
                            }).Select(m => new ArticuloCantidad
                            {
                                Articulo = m.Key.Articulo,
                                Cantidad = m.Key.Cantidad
                            }).ToList();


                            ///////////////////////////////
                            //PAGINA 2 REPORTE CEO MECANICA                            
                            finDatosHtml.AppendLine("</tbody>");
                            finDatosHtml.AppendLine("</table>");
                            finDatosHtml.AppendLine("</center>");

                            finHtml.AppendLine("</div>");

                            inicioHtml.AppendLine("<div style='font-family: arial; position: relative;'>");

                            html.AppendLine(inicioHtml.ToString());

                            html.AppendLine("<center><h3>2) Mecanica</h3></center>");

                            inicioDatosHtml.AppendLine("<center>");
                            inicioDatosHtml.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                            inicioDatosHtml.AppendLine("<thead style='background-color:#3B89D8; font-size:13px; color:white;'>");
                            inicioDatosHtml.AppendLine("<tr>");

                            ListPeriodoAntAct = new List<PeriodoCEOMecanica>();


                            foreach (int periodo in ListPeriodo.OrderBy(n => n).ToList())
                            {
                                periodoMecanica = new PeriodoCEOMecanica();
                                periodoMecanica.Periodo = periodo.ToString();

                                if (periodoStyle)
                                {
                                    //COLOR VERDE
                                    periodoStyle = !periodoStyle;

                                    periodoMecanica.EColor = "background-color:#2E8B57;";
                                    periodoMecanica.Color = "background-color:#ABEBC6;";

                                    periodoMecanica.EAlpha = 1;
                                    periodoMecanica.ERed = 46;
                                    periodoMecanica.EGreen = 139;
                                    periodoMecanica.EBlue = 87;

                                    periodoMecanica.Alpha = 1;
                                    periodoMecanica.Red = 171;
                                    periodoMecanica.Green = 235;
                                    periodoMecanica.Blue = 198;
                                }
                                else
                                {
                                    //COLOR AZUL
                                    periodoStyle = !periodoStyle;

                                    periodoMecanica.EColor = "background-color:#3B89D8;";
                                    periodoMecanica.Color = "background-color:#CEE3F6";

                                    periodoMecanica.EAlpha = 1;
                                    periodoMecanica.ERed = 59;
                                    periodoMecanica.EGreen = 137;
                                    periodoMecanica.EBlue = 216;

                                    periodoMecanica.Alpha = 1;
                                    periodoMecanica.Red = 206;
                                    periodoMecanica.Green = 227;
                                    periodoMecanica.Blue = 246;
                                }

                                ListPeriodoAntAct.Add(periodoMecanica);
                            }

                            foreach (DataColumn dataColumn in dtReporteCEOMecanica.Columns)
                            {
                                periodoMecanica = new PeriodoCEOMecanica();
                                periodoMecanica = ListPeriodoAntAct.Where(n => dataColumn.ColumnName.ToUpper().Contains(n.Periodo.ToUpper())).FirstOrDefault();

                                if (periodoMecanica != null)
                                {
                                    inicioDatosHtml.AppendLine("<th style='text-align:left; border: 0.5px solid black; " + periodoMecanica.EColor + "'>" + dataColumn.ColumnName + "</th>");
                                }
                                else
                                {
                                    inicioDatosHtml.AppendLine("<th style='text-align:left; border: 0.5px solid black; background-color:#3B89D8;'>" + dataColumn.ColumnName + "</th>");
                                }
                            }


                            //FIN COLUMNAS
                            inicioDatosHtml.AppendLine("</tr>");
                            inicioDatosHtml.AppendLine("</thead>");

                            //LLENAR DATOS REPORTE CEO MECANICA
                            inicioDatosHtml.AppendLine("<tbody style='background-color:#CEE3F6; font-size:12px;'>");

                            html.AppendLine(inicioDatosHtml.ToString());

                            //columnStyle = true;

                            foreach (ArticuloCantidad articuloCantidad in ListArticuloCantidad)
                            {
                                ListDatosCeoMecanica = new List<string>();

                                ListDatosCeoMecanica.Add(articuloCantidad.Articulo);
                                ListDatosCeoMecanica.Add(articuloCantidad.Cantidad);

                                //if (columnStyle)
                                //{
                                //    columnStyleHtml = "style='background-color:#ffffff'";
                                //    columnStyle = !columnStyle;
                                //}
                                //else
                                //{
                                //    columnStyleHtml = "style='background-color:#CEE3F6'";
                                //    columnStyle = !columnStyle;
                                //}

                                //PDF
                                html.AppendLine("<tr>");

                                html.AppendLine("<td style='text-align:left; background-color:#CEE3F6; border: 0.5px solid black;'>" + ListDatosCeoMecanica[0].ToString() + "</td>");
                                html.AppendLine("<td style='text-align:left; background-color:#CEE3F6; border: 0.5px solid black;'>" + ListDatosCeoMecanica[1].ToString() + "</td>");


                                for (int i = 2; i < dtReporteCEOMecanica.Columns.Count; i++)
                                {
                                    var nombreColumnas = dtReporteCEOMecanica.Columns[i].ColumnName.Split('-');

                                    if (nombreColumnas != null && nombreColumnas.Count() > 1)
                                    {
                                        ReporteCEOMecanicaTemp = ListReporteCEOMecanica.Where(n => n.Articulo.ToUpper() == articuloCantidad.Articulo.ToUpper() &&
                                                                n.Cantidad.ToUpper() == articuloCantidad.Cantidad.ToUpper() &&
                                                                n.Periodo == Convert.ToInt32(nombreColumnas[0]) &&
                                                                n.Alcance.ToUpper() == nombreColumnas[1].ToUpper()).ToList().FirstOrDefault();

                                        if (ReporteCEOMecanicaTemp != null)
                                        {
                                            //ListDatosCeoMecanica.Add(ReporteCEOMecanicaTemp.Periodo.ToString());
                                            ListDatosCeoMecanica.Add(ReporteCEOMecanicaTemp.Mecanica);

                                            periodoMecanica = new PeriodoCEOMecanica();
                                            periodoMecanica = ListPeriodoAntAct.Where(n => dtReporteCEOMecanica.Columns[i].ColumnName.ToUpper().Contains(n.Periodo.ToUpper())).FirstOrDefault();

                                            if (periodoMecanica != null)
                                            {
                                                html.AppendLine("<td style='text-align:left; border: 0.5px solid black; " + periodoMecanica.Color + "'>" + ReporteCEOMecanicaTemp.Mecanica + "</td>");
                                            }
                                            else
                                            {
                                                html.AppendLine("<td style='text-align:left; border: 0.5px solid black; " + periodoMecanica.Color + "'>" + ReporteCEOMecanicaTemp.Mecanica + "</td>");
                                            }
                                        }
                                        else
                                        {
                                            //ListDatosCeoMecanica.Add(string.Empty);
                                            ListDatosCeoMecanica.Add(string.Empty);

                                            html.AppendLine("<td style='text-align:left; border: 0.5px solid black; " + periodoMecanica.Color + "'>" + string.Empty + "</td>");
                                        }
                                    }
                                    else
                                    {
                                        //ListDatosCeoMecanica.Add(string.Empty);
                                        ListDatosCeoMecanica.Add(string.Empty);

                                        html.AppendLine("<td style='text-align:left; border: 0.5px solid black; " + periodoMecanica.Color + "'>" + string.Empty + "</td>");
                                    }
                                }

                                dtReporteCEOMecanica.Rows.Add(ListDatosCeoMecanica.ToArray());

                                html.AppendLine("</tr>");


                                if (contRegCEO > RegPrimerPaginaCEOMerca)
                                {
                                    html.AppendLine(finDatosHtml.ToString());
                                    html.AppendLine(finHtml.ToString());

                                    ListHtmlCEO.Add(html.ToString());

                                    html = new StringBuilder();

                                    html.AppendLine(inicioHtml.ToString());
                                    html.AppendLine(inicioDatosHtml.ToString());

                                    contRegCEO = 0;
                                }

                                contRegCEO++;

                            }

                            //html.AppendLine("</tbody>");
                            //html.AppendLine("</table>");

                            html.AppendLine(finDatosHtml.ToString());

                            //html.AppendLine("</div>");
                            //html.AppendLine(ImgHtmlFooter);
                            //html.AppendLine("</div>");

                            html.AppendLine(finHtml.ToString());

                            //html.AppendLine("<br />");
                            //html.AppendLine("<br />");

                            ListHtmlCEO.Add(html.ToString());
                        }


                        #endregion

                        #region REPORTE CEO PUBLICIDAD

                        /////////////////////////////////
                        //CARGA INFORMACION DE PUBLICIDAD

                        if (ListReporteCEOPublicidad.Count > 0)
                        {
                            html = new StringBuilder();
                            inicioHtml = new StringBuilder();
                            inicioDatosHtml = new StringBuilder();
                            finDatosHtml = new StringBuilder();
                            finHtml = new StringBuilder();
                            contRegCEO = 0;

                            dtReporteCEOPublicidad.Columns.Add("Publicidad");
                            dtReporteCEOPublicidad.Columns.Add("Periodo Anterior");
                            dtReporteCEOPublicidad.Columns.Add("Periodo Actual");

                            /////////////////////////////////
                            //PAGINA 3 REPORTE CEO PUBLICIDAD
                            finDatosHtml.AppendLine("</tbody>");
                            finDatosHtml.AppendLine("</table>");
                            finDatosHtml.AppendLine("</center>");

                            finHtml.AppendLine("</div>");

                            inicioHtml.AppendLine("<div style='font-family: arial; position: relative;'>");

                            html.AppendLine(inicioHtml.ToString());
                            html.AppendLine("<center><h3>3) Inversion Publicitaria</h3></center>");

                            inicioDatosHtml.AppendLine("<center>");
                            inicioDatosHtml.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                            inicioDatosHtml.AppendLine("<thead style='background-color:#3B89D8; font-size:13px; color:white;'>");
                            inicioDatosHtml.AppendLine("<tr>");

                            inicioDatosHtml.AppendLine("<th style='text-align:left; background-color:#3B89D8; border: 0.5px solid black;'>Publicidad</th>");
                            inicioDatosHtml.AppendLine("<th style='text-align:left; background-color:#2E8B57; border: 0.5px solid black;'>Periodo Anterior</th>");
                            inicioDatosHtml.AppendLine("<th style='text-align:left; background-color:#3B89D8; border: 0.5px solid black;'>Periodo Actual</th>");

                            //FIN COLUMNAS
                            inicioDatosHtml.AppendLine("</tr>");
                            inicioDatosHtml.AppendLine("</thead>");

                            //LLENAR DATOS REPORTE CEO
                            inicioDatosHtml.AppendLine("<tbody style='background-color:#CEE3F6; font-size:12px;'>");

                            html.AppendLine(inicioDatosHtml.ToString());

                            //columnStyle = true;

                            foreach (ReporteCEOPublicidad reporteCEOPublicidad in ListReporteCEOPublicidad)
                            {
                                dtReporteCEOPublicidad.Rows.Add(reporteCEOPublicidad.Publicidad, reporteCEOPublicidad.PeriodoAnterior.ToString("#,##0.00"), reporteCEOPublicidad.PeriodoActual.ToString("#,##0.00"));

                                //if (columnStyle)
                                //{
                                //    columnStyleHtml = "style='background-color:#ffffff'";
                                //    columnStyle = !columnStyle;
                                //}
                                //else
                                //{
                                //    columnStyleHtml = "style='background-color:#CEE3F6'";
                                //    columnStyle = !columnStyle;
                                //}

                                //PDF
                                html.AppendLine("<tr>");

                                html.AppendLine("<td style='background-color:#CEE3F6; text-align:left; border: 0.5px solid black;'>" + reporteCEOPublicidad.Publicidad + "</td>");
                                html.AppendLine("<td style='background-color:#ABEBC6; text-align:right; border: 0.5px solid black;'>" + reporteCEOPublicidad.PeriodoAnterior.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='background-color:#CEE3F6; text-align:right; border: 0.5px solid black;'>" + reporteCEOPublicidad.PeriodoActual.ToString("#,##0.00") + "</td>");

                                html.AppendLine("</tr>");


                                if (contRegCEO > RegPrimerPaginaCEOPubli)
                                {
                                    html.AppendLine(finDatosHtml.ToString());
                                    html.AppendLine(finHtml.ToString());

                                    ListHtmlCEO.Add(html.ToString());

                                    html = new StringBuilder();

                                    html.AppendLine(inicioHtml.ToString());
                                    html.AppendLine(inicioDatosHtml.ToString());

                                    contRegCEO = 0;

                                }

                                contRegCEO++;

                            }

                            if (ListReporteCEOPublicidad.Count > 0)
                            {
                                totalPublicidadPAnt = ListReporteCEOPublicidad.Sum(n => n.PeriodoAnterior);
                                totalPublicidadPAct = ListReporteCEOPublicidad.Sum(n => n.PeriodoActual);

                                dtReporteCEOPublicidad.Rows.Add("Total Inversión Publicitaria", totalPublicidadPAnt.ToString("#,##0.00"), totalPublicidadPAct.ToString("#,##0.00"));

                                //AGREGAR TOTALES
                                html.AppendLine("<tr style='color:white;'>");

                                html.AppendLine("<td style='background-color:#3B89D8; text-align:left; border: 0.5px solid black;'>Total Inversión Publicitaria</td>");
                                html.AppendLine("<td style='background-color:#228B22; text-align:right; border: 0.5px solid black;'>" + totalPublicidadPAnt.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='background-color:#3B89D8; text-align:right; border: 0.5px solid black;'>" + totalPublicidadPAct.ToString("#,##0.00") + "</td>");

                                html.AppendLine("</tr>");
                            }

                            //html.AppendLine("</tbody>");
                            //html.AppendLine("</table>");
                            //html.AppendLine("</div>");
                            //html.AppendLine(ImgHtmlFooter);

                            html.AppendLine(finDatosHtml.ToString());

                            //html.AppendLine("</div>");
                            //html.AppendLine("<br />");
                            //html.AppendLine("<br />");

                            html.AppendLine(finHtml.ToString());

                            //FIN HTML
                            //html.AppendLine("</div>");

                            ListHtmlCEO.Add(html.ToString());
                        }

                        #endregion

                        //CREAR ARCHIVO PDF
                        //nombreArchivo = Guid.NewGuid().ToString() + ".pdf";

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                               ConfigurationManager.AppSettings["DirectorioReporte"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathArchivo = parametro.Valor;
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

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["WidthPDFVertical"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            widthPDFVertical = Convert.ToDouble(parametro.Valor);
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["HeigthPDFVertical"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            heigthPDFVertical = Convert.ToDouble(parametro.Valor);
                        }


                        //////////////////////
                        //GENERAR ARCHIVO PDF

                        //var Renderer = new IronPdf.HtmlToPdf();

                        //HtmlHeaderFooter htmlFooter = new HtmlHeaderFooter();
                        //htmlFooter.HtmlFragment = ImgHtmlFooter;
                        //Renderer.PrintOptions.Footer = htmlFooter;

                        //HtmlHeaderFooter htmlHeader = new HtmlHeaderFooter();
                        //htmlHeader.HtmlFragment = ImgHtmlHeader;
                        //Renderer.PrintOptions.Header = htmlHeader;

                        //Renderer.PrintOptions.SetCustomPaperSizeinMilimeters(widthPDFVertical, heigthPDFVertical);
                        //Renderer.PrintOptions.Header.DrawDividerLine = false;

                        //htmlReporteCEO = html.ToString();

                        //var PDF = Renderer.RenderHtmlAsPdf(uno);
                        ////PDF.WatermarkAllPages("<div><h1>UNO DOS TRES CUATRO</h1></div>", PdfDocument.WaterMarkLocation.MiddleCenter, 50, 50, "");
                        //PDF.SaveAs(pathArchivoCompleto);

                        //HtmlToPdf htmlToPdf = new HtmlToPdf();
                        //var PDF = htmlToPdf.RenderHtmlAsPdf(uno);
                        //PDF.SaveAs(pathArchivoCompleto);

                        //var PDF = new NReco.PdfGenerator.HtmlToPdfConverter();
                        //PDF.GeneratePdf(uno, null, pathArchivoCompleto);

                        //var PDF = new ExpertPdf.HtmlToPdf.PdfConverter().GetPdfDocumentObjectFromHtmlString(uno);
                        //PDF.Save(pathArchivoCompleto);


                        ////GENERAR PDF CON DLL SPIRE
                        //layoutCEO = new PdfHtmlLayoutFormat();
                        //layoutCEO.IsWaiting = true;

                        //pageCEO = new PdfPageSettings();
                        //pageCEO.Size = new SizeF(762, 986); //Spire.Pdf.PdfPageSize.Letter; //new SizeF(762, 986);

                        //int left = 15;
                        //int rigth = 15;
                        //int top = 35;
                        //int bottom = 35;

                        //pageCEO.Margins = new Spire.Pdf.Graphics.PdfMargins(left, top, rigth, bottom);


                        //docCEO = new PdfDocument();

                        //Thread thread = new Thread(() =>
                        //{
                        //    docCEO.LoadFromHTML(htmlReporteCEO, autoDetectPage, pageCEO, layoutCEO);
                        //});


                        //thread.SetApartmentState(ApartmentState.STA);
                        //thread.Start();
                        //thread.Join();

                        //docCEO.SaveToFile(pathArchivoCompleto);
                        //docCEO.Close();

                        /////////////////////
                        ////AGREGAR HEADER Y FOOTER
                        //docCEOImg = new PdfDocument();

                        //Thread thread2 = new Thread(() =>
                        //{
                        //    docCEOImg.LoadFromFile(pathArchivoCompleto);
                        //});

                        //headerImage = PdfImage.FromFile(pathImgHeader);
                        //footerImage = PdfImage.FromFile(pathImgFooter);

                        //thread2.SetApartmentState(ApartmentState.STA);
                        //thread2.Start();
                        //thread2.Join();

                        //int xHeader = 5;
                        //int yHeader = 5;
                        //int xFooter = 45;
                        //int yFooter = 961;

                        //foreach (PdfPageBase pdfPage in docCEOImg.Pages)
                        //{
                        //    pdfPage.Canvas.DrawImage(headerImage, xHeader, yHeader, 712, 20);
                        //    pdfPage.Canvas.DrawImage(footerImage, xFooter, yFooter, 712, 20);
                        //}

                        //docCEOImg.SaveToFile(pathArchivoCompleto);
                        //docCEOImg.Close();


                        //REPORTE CEO
                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                        ConfigurationManager.AppSettings["UrlReporteCEO"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            url = HttpUtility.HtmlEncode(parametro.Valor);
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                               ConfigurationManager.AppSettings["NombreReporteCEO"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            nombreReporte = parametro.Valor + rentabilidad.ClaveCampania;
                        }


                        //NOMBRE REPORTE CARPETA TEMPORAL
                        pathArchivoCompletoCEO = Path.Combine(pathArchivo, nombreReporte + ".pdf");

                        //NOMBRE REPORTE SHAREPOINT
                        urlCompletoCEO = Path.Combine(url, nombreReporte + ".pdf");

                        rentabilidadENTRes.UrlReporteCEOPDF = urlCompletoCEO;


                        var t = System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                ////////////////////
                                //GENERAR PDF CEO
                                //PDFTool.PDF.GenerarPDF(htmlReporteCEO, pathImgHeader, pathImgFooter, pathArchivoCompletoCEO, 792, 612);

                                PDFTool.PDF.GenerarPDFCEO(ListHtmlCEO, pathImgHeader, pathImgFooter, pathArchivoCompletoCEO, 792, 612);

                                return "OK";
                            }
                            catch(Exception ex)
                            {
                                ArchivoLog.EscribirLog(null, "ERROR: Service - PDF CEO, Source:" + ex.Source + ", Message:" + ex.Message);

                                return "ERROR: Service - PDF CEO, Source:" + ex.Source + ", Message:" + ex.Message;
                            }                         
                        });

                        var c = t.ContinueWith((antecedent) =>
                        {
                            if (antecedent.Result == "OK")
                            {
                                using (WebClient client = new WebClient())
                                {
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                                    client.UploadFile(urlCompletoCEO, "PUT", pathArchivoCompletoCEO);
                                   
                                }
                            }
                           
                        });


                        //////////////////////
                        ////GENERAR PDF NUEVO
                        //Thread threadCEO = new Thread(() =>
                        //{
                        //    PDFTool.PDF.GenerarPDF(htmlReporteCEO, pathImgHeader, pathImgFooter, pathArchivoCompleto, 792, 612);
                        //});


                        //threadCEO.SetApartmentState(ApartmentState.STA);            
                        //threadCEO.IsBackground = true;
                        //threadCEO.Start();

                        #region CREAR ARCHIVO EXCEL

                        /////////////////////
                        //CREAR ARCHIVO EXCEL

                        //dsReporteCEO.Tables.Add(dtReporteCEO);

                        //nombreArchivo = Guid.NewGuid().ToString() + ".xlsx

                        //NOMBRE REPORTE CARPETA TEMPORAL
                        pathArchivoCompletoCEOEx = Path.Combine(pathArchivo, nombreReporte + ".xlsx");

                        //ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoCompleto, dsReporteCEO);

                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            excel.Workbook.Worksheets.Add("Analisis");
                            excel.Workbook.Worksheets.Add("Mecanica");
                            excel.Workbook.Worksheets.Add("Publicidad");

                            #region HOJA ANALISIS

                            /////////////////////////////
                            //PESTAÑA ANALISIS
                            var worksheet = excel.Workbook.Worksheets["Analisis"];

                            worksheet.SetValue("B1", "Analisis");
                            worksheet.SetValue("B2", "Campaña:");
                            worksheet.SetValue("C2", nombreCampana);
                            worksheet.SetValue("B3", "Clave Campaña:");
                            worksheet.SetValue("C3", claveCampana);
                            worksheet.SetValue("B4", "Fecha Inicio SubCanal:");

                            if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal != "01/01/0001")
                            {
                                worksheet.SetValue("C4", fechaInicioSubCanal);
                            }
                            worksheet.SetValue("B5", "Fecha Fin SubCanal:");
                            if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal != "01/01/0001")
                            {
                                worksheet.SetValue("C5", fechaFinSubCanal);
                            }
                            worksheet.SetValue("B6", "Fecha Inicio Publico");
                            if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico != "01/01/0001")
                            {
                                worksheet.SetValue("C6", fechaInicioPublico);
                            }
                            worksheet.SetValue("B7", "Fecha Fin Publico");
                            if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico != "01/01/0001")
                            {
                                worksheet.SetValue("C7", fechaFinPublico);
                            }

                            worksheet.Cells["B1:C7"].Style.Font.Color.SetColor(1, 38, 130, 221);
                            worksheet.Cells["B1:C7"].Style.Font.Bold = true;
                            worksheet.Cells["B1:C7"].Style.Font.Size = 14;

                            worksheet.Cells["B9:" + Char.ConvertFromUtf32(dtReporteCEO.Columns.Count + 65) + "9"].LoadFromDataTable(dtReporteCEO, true);//, OfficeOpenXml.Table.TableStyles.Light2);

                            worksheet.Cells["B9:" + Char.ConvertFromUtf32(dtReporteCEO.Columns.Count + 65) + "9"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            worksheet.Cells["B9:" + Char.ConvertFromUtf32(dtReporteCEO.Columns.Count + 65) + "9"].Style.Fill.BackgroundColor.SetColor(1, 59, 137, 216);
                            worksheet.Cells["B9:" + Char.ConvertFromUtf32(dtReporteCEO.Columns.Count + 65) + "9"].Style.Font.Color.SetColor(Color.White);

                            worksheet.Cells["C9"].Style.Fill.BackgroundColor.SetColor(1, 46, 139, 87);

                            worksheet.Cells["B10:E35"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["B9:E35"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["B9:E35"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["B9:E35"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["B9:E35"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            worksheet.Cells["B10:E12"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B13:E13"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                            worksheet.Cells["B14:E15"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B16:E16"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                            worksheet.Cells["B17:E17"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B18:E18"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                            worksheet.Cells["B19:E19"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B20:E20"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                            worksheet.Cells["B21:E24"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B25:E26"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                            worksheet.Cells["B27:E27"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B28:E28"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                            worksheet.Cells["B29:E31"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B32:E32"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                            worksheet.Cells["B33:E34"].Style.Fill.BackgroundColor.SetColor(Color.White);
                            worksheet.Cells["B35:E35"].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);

                            worksheet.Cells["B1:E35"].AutoFitColumns();


                            //AGREGAR ENCABEZADO
                            worksheet.Cells["A12:A31"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A12:A31"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A12:A31"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A12:A31"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A12:A31"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //worksheet.Cells["A12:A31"].Style.Font.Color.SetColor(Color.White);

                            worksheet.Cells["A12:A31"].Style.Fill.BackgroundColor.SetColor(Color.White);

                            worksheet.Cells["A12:A31"].Merge = true;

                            worksheet.Cells["A12:A31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            worksheet.Cells["A12:A31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            worksheet.Cells["A12:A31"].Style.TextRotation = 90;

                            worksheet.Cells["A12:A31"].Value = "SELL IN";


                            worksheet.Cells["A32:A35"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A32:A35"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A32:A35"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A32:A35"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A32:A35"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            //worksheet.Cells["A12:A31"].Style.Font.Color.SetColor(Color.White);

                            worksheet.Cells["A32:A35"].Style.Fill.BackgroundColor.SetColor(Color.White);

                            worksheet.Cells["A32:A35"].Merge = true;

                            worksheet.Cells["A32:A35"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            worksheet.Cells["A32:A35"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            worksheet.Cells["A32:A35"].Style.TextRotation = 90;

                            worksheet.Cells["A32:A35"].Value = "SELL OUT";


                            #endregion

                            #region HOJA MECANICA

                            if (dtReporteCEOMecanica.Rows.Count > 0)
                            {
                                ///////////////////////////////
                                //PESTAÑA MECANICA
                                worksheet = excel.Workbook.Worksheets["Mecanica"];

                                worksheet.SetValue("A1", "Mecanica");
                                worksheet.SetValue("A2", "Campaña:");
                                worksheet.SetValue("B2", nombreCampana);
                                worksheet.SetValue("A3", "Clave Campaña:");
                                worksheet.SetValue("B3", claveCampana);
                                worksheet.SetValue("A4", "Fecha Inicio SubCanal:");

                                if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal != "01/01/0001")
                                {
                                    worksheet.SetValue("B4", fechaInicioSubCanal);
                                }
                                worksheet.SetValue("A5", "Fecha Fin SubCanal:");
                                if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal != "01/01/0001")
                                {
                                    worksheet.SetValue("B5", fechaFinSubCanal);
                                }
                                worksheet.SetValue("A6", "Fecha Inicio Publico");
                                if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico != "01/01/0001")
                                {
                                    worksheet.SetValue("B6", fechaInicioPublico);
                                }
                                worksheet.SetValue("A7", "Fecha Fin Publico");
                                if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico != "01/01/0001")
                                {
                                    worksheet.SetValue("B7", fechaFinPublico);
                                }

                                worksheet.Cells["A1:B7"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                worksheet.Cells["A1:B7"].Style.Font.Bold = true;
                                worksheet.Cells["A1:B7"].Style.Font.Size = 14;

                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + "9"].LoadFromDataTable(dtReporteCEOMecanica, true);//, OfficeOpenXml.Table.TableStyles.Light2);

                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + "9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + "9"].Style.Fill.BackgroundColor.SetColor(1, 59, 137, 216);
                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + "9"].Style.Font.Color.SetColor(Color.White);

                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + "9"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Fill.PatternType = ExcelFillStyle.Solid;

                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                worksheet.Cells["A1:" + Char.ConvertFromUtf32(dtReporteCEOMecanica.Columns.Count + 64) + (dtReporteCEOMecanica.Rows.Count + 9)].AutoFitColumns();


                                for (int i = 2; i < dtReporteCEOMecanica.Columns.Count; i++)
                                {
                                    periodoMecanica = new PeriodoCEOMecanica();
                                    periodoMecanica = ListPeriodoAntAct.Where(n => dtReporteCEOMecanica.Columns[i].ColumnName.ToUpper().Contains(n.Periodo.ToUpper())).FirstOrDefault();

                                    if (periodoMecanica != null)
                                    {
                                        worksheet.Cells[Char.ConvertFromUtf32(i + 1 + 64) + "9"].Style.Fill.BackgroundColor.SetColor(periodoMecanica.EAlpha, periodoMecanica.ERed, periodoMecanica.EGreen, periodoMecanica.EBlue);

                                        worksheet.Cells[Char.ConvertFromUtf32(i + 1 + 64) + "10:" + Char.ConvertFromUtf32(i + 1 + 64) + (dtReporteCEOMecanica.Rows.Count + 9).ToString()].Style.Fill.BackgroundColor.SetColor(periodoMecanica.Alpha, periodoMecanica.Red, periodoMecanica.Green, periodoMecanica.Blue);
                                    }
                                }

                                if (dtReporteCEOMecanica.Rows.Count > 0)
                                {
                                    worksheet.Cells["A10:A" + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 171, 235, 198);
                                    worksheet.Cells["B10:B" + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 171, 235, 198);
                                }

                                //worksheet.Cells["C10:C" + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 77, 178, 30);
                                //worksheet.Cells["D10:D" + (dtReporteCEOMecanica.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 77, 178, 30);

                                //worksheet.Cells["C9"].Style.Fill.BackgroundColor.SetColor(1, 76, 127, 15);
                                //worksheet.Cells["D9"].Style.Fill.BackgroundColor.SetColor(1, 76, 127, 15);
                            }

                            #endregion

                            #region HOJA PUBLICIDAD

                            if (dtReporteCEOPublicidad.Rows.Count > 0)
                            {
                                //////////////////////////////
                                //PESTAÑA PUBLICIDAD
                                worksheet = excel.Workbook.Worksheets["Publicidad"];

                                worksheet.SetValue("A1", "Mecanica");
                                worksheet.SetValue("A2", "Campaña:");
                                worksheet.SetValue("B2", nombreCampana);
                                worksheet.SetValue("A3", "Clave Campaña:");
                                worksheet.SetValue("B3", claveCampana);
                                worksheet.SetValue("A4", "Fecha Inicio SubCanal:");
                                if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal != "01/01/0001")
                                {
                                    worksheet.SetValue("B4", fechaInicioSubCanal);
                                }
                                worksheet.SetValue("A5", "Fecha Fin SubCanal:");
                                if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal != "01/01/0001")
                                {
                                    worksheet.SetValue("B5", fechaFinSubCanal);
                                }
                                worksheet.SetValue("A6", "Fecha Inicio Publico");
                                if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico != "01/01/0001")
                                {
                                    worksheet.SetValue("B6", fechaInicioPublico);
                                }
                                worksheet.SetValue("A7", "Fecha Fin Publico");
                                if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico != "01/01/0001")
                                {
                                    worksheet.SetValue("B7", fechaFinPublico);
                                }

                                worksheet.Cells["A1:B7"].Style.Font.Color.SetColor(1, 38, 130, 221);
                                worksheet.Cells["A1:B7"].Style.Font.Bold = true;
                                worksheet.Cells["A1:B7"].Style.Font.Size = 14;

                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOPublicidad.Columns.Count + 64) + "9"].LoadFromDataTable(dtReporteCEOPublicidad, true);//, OfficeOpenXml.Table.TableStyles.Light2);

                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOPublicidad.Columns.Count + 64) + "9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOPublicidad.Columns.Count + 64) + "9"].Style.Fill.BackgroundColor.SetColor(1, 59, 137, 216);
                                worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteCEOPublicidad.Columns.Count + 64) + "9"].Style.Font.Color.SetColor(Color.White);

                                worksheet.Cells["B9"].Style.Fill.BackgroundColor.SetColor(1, 46, 139, 87);

                                if (dtReporteCEOPublicidad.Rows.Count > 0)
                                {
                                    worksheet.Cells["A10:C" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells["A10:A" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                                    worksheet.Cells["B10:B" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 171, 235, 198);
                                    worksheet.Cells["C10:C" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 206, 227, 246);
                                }

                                worksheet.Cells["A9:C" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells["A9:C" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells["A9:C" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells["A9:C" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                                worksheet.Cells["A" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 59, 137, 216);
                                worksheet.Cells["B" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 46, 139, 87);
                                worksheet.Cells["C" + (dtReporteCEOPublicidad.Rows.Count + 9)].Style.Fill.BackgroundColor.SetColor(1, 59, 137, 216);

                                worksheet.Cells["A1:C" + (dtReporteCEOPublicidad.Rows.Count + 9)].AutoFitColumns();
                            }

                            #endregion

                            FileInfo excelFile = new FileInfo(pathArchivoCompletoCEOEx);
                            excel.SaveAs(excelFile);
                        }

                        //NOMBRE REPORTE SHARE POINT
                        urlCompletoCEOEx = Path.Combine(url, nombreReporte + ".xlsx");

                        using (WebClient client = new WebClient())
                        {
                            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                            client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                            client.UploadFile(urlCompletoCEOEx, "PUT", pathArchivoCompletoCEOEx);

                            rentabilidadENTRes.UrlReporteCEOExcel = urlCompletoCEOEx;
                        }

                        #endregion

                    }

                    #endregion


                    #region REPORTE MKT

                    //REPORTE MKT
                    if (ListReporteMKT.Count > 0)
                    {
                        html = new StringBuilder();
                        inicioHtml = new StringBuilder();
                        inicioDatosHtml = new StringBuilder();
                        finDatosHtml = new StringBuilder();
                        finHtml = new StringBuilder();
                        tablaInicioHtml = new StringBuilder();
                        ListHtmlMercadotecnia = new List<string>();

                        //PDF
                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgComex"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgComex = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgPPG"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgPPG = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgHeader"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgHeader = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgFooter"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgFooter = parametro.Valor;
                        }

                        //var pngBinaryDataComex = File.ReadAllBytes(pathImgComex);
                        //var ImgDataURIComex = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataComex);
                        //var ImgHtmlComex = String.Format("<img src='{0}' style='float:right;vertical-align:top;' width='120' height='50'>", ImgDataURIComex);

                        //var pngBinaryDataPPG = File.ReadAllBytes(pathImgPPG);
                        //var ImgDataURIPPG = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataPPG);
                        //var ImgHtmlPPG = String.Format("<img src='{0}' style='float:left;vertical-align:top;' width='70' height='50'>", ImgDataURIPPG);

                        var pngBinaryDataHeaer = File.ReadAllBytes(pathImgHeader);
                        var ImgDataURIHeader = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataHeaer);
                        var pngBinaryDataFooter = File.ReadAllBytes(pathImgFooter);
                        var ImgDataURIFooter = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataFooter);

                        //var ImgHtmlHeader = String.Format("<div style='width: 100%; height: 50px; position: absolute; top: 1px;'><img src='{0}' width='100%' height='100%'></div>", ImgDataURIHeader);
                        //var ImgHtmlFooter = String.Format("<footer style='width: 100%; height: 50px; flow: static(footer);'><img src='{0}' width='100%' height='100%'></footer>", ImgDataURIFooter);
                        //var ImgHtmlFooter = String.Format("<div style='width: 100%; height: 50px; position: absolute; bottom: 1px;'><img src='{0}' width='100%' height='100%'></div>", ImgDataURIFooter);


                        //html.AppendLine(ImgHtmlComex);
                        inicioHtml.AppendLine("<div style='font-family: arial; position: relative; height: 100%;'>");
                        //html.AppendLine(ImgHtmlHeader);
                        //html.AppendLine(ImgHtmlPPG);
                        inicioHtml.AppendLine("<div style='font-family:arial; position: absolute;'>");
                        html.AppendLine(inicioHtml.ToString());

                        tablaInicioHtml.AppendLine("<center><h3>Reporte Mercadotecnia</h3></center>");
                        //TABLA
                        tablaInicioHtml.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                        tablaInicioHtml.AppendLine("<tbody>");
                        //REGISTRO 1
                        tablaInicioHtml.AppendLine("<tr style='background-color:#ffffff;'>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Campaña</td>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + nombreCampana + "</td>");
                        tablaInicioHtml.AppendLine("</tr>");
                        //REGISTRO 2
                        tablaInicioHtml.AppendLine("<tr style='background-color:#CEE3F6;'>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Clave Campaña</td>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + claveCampana + "</td>");
                        tablaInicioHtml.AppendLine("</tr>");
                        //REGISTRO 3
                        if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#ffffff;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Inicio SubCanal</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + fechaInicioSubCanal + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }
                        //REGISTRO 4
                        if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#CEE3F6;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Fin SubCanal</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + fechaFinSubCanal + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }
                        //REGISTRO 5
                        if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#ffffff;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Inicio Publico</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + fechaInicioPublico + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }
                        //REGISTRO 6
                        if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#CEE3F6;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Fin Publico</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + fechaFinPublico + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }

                        tablaInicioHtml.AppendLine("</tbody>");
                        tablaInicioHtml.AppendLine("</table>");
                        tablaInicioHtml.AppendLine("<br />");

                        html.AppendLine(tablaInicioHtml.ToString());

                        inicioDatosHtml.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                        inicioDatosHtml.AppendLine("<thead style='background-color:#3B89D8; font-size:12px; color:white;'>");
                        inicioDatosHtml.AppendLine("<tr>");
                        //AGREGAR COLUMNAS
                        inicioDatosHtml.AppendLine("<th style='width:70px; background-color:#3B89D8; border: 0.5px solid black;'>SEGMENTO</th>");
                        inicioDatosHtml.AppendLine("<th style='background-color:#3B89D8; border: 0.5px solid black;'>ARTICULO</th>");
                        inicioDatosHtml.AppendLine("<th style='width:130px; background-color:#3B89D8; border: 0.5px solid black;'>DESCRIPCION</th>");
                        inicioDatosHtml.AppendLine("<th style='width:50px; background-color:#3B89D8; border: 0.5px solid black;'>LITROS ACTUAL CC</th>");
                        inicioDatosHtml.AppendLine("<th style='width:50px; background-color:#3B89D8; border: 0.5px solid black;'>PIEZAS ACTUAL CC</th>");
                        inicioDatosHtml.AppendLine("<th style='width:60px; background-color:#3B89D8; border: 0.5px solid black;'>IMPORTE C/PRECIO CONC. LITROS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:60px; background-color:#3B89D8; border: 0.5px solid black;'>IMPORTE C/PRECIO PIEZAS</th>");
                        inicioDatosHtml.AppendLine("<th style='background-color:#3B89D8; border: 0.5px solid black;'>RENTABILIDAD</th>");
                        inicioDatosHtml.AppendLine("<th style='width:50px; background-color:#2E8B57; border: 0.5px solid black;'>LITROS NECESARIO</th>");
                        inicioDatosHtml.AppendLine("<th style='width:50px;background-color:#2E8B57; border: 0.5px solid black;'>PIEZAS NECESARIO</th>");
                        //FIN COLUMNAS
                        inicioDatosHtml.AppendLine("</tr>");
                        inicioDatosHtml.AppendLine("</thead>");
                        inicioDatosHtml.AppendLine("<tbody style='font-size:11px;'>");

                        html.AppendLine(inicioDatosHtml.ToString());

                        finDatosHtml.AppendLine("</tbody>");
                        finDatosHtml.AppendLine("</table>");
                        finHtml.AppendLine("</div>");
                        finHtml.AppendLine("</div>");


                        //EXCEL
                        dtReporteMKT.Columns.Add("SEGMENTO");
                        dtReporteMKT.Columns.Add("ARTICULO");
                        dtReporteMKT.Columns.Add("DESCRIPCION");
                        dtReporteMKT.Columns.Add("LITROS ACTUAL CC");
                        dtReporteMKT.Columns.Add("PIEZAS ACTUAL CC");
                        dtReporteMKT.Columns.Add("IMPORTE C/PRECIO CONC. LITROS");
                        dtReporteMKT.Columns.Add("IMPORTE C/PRECIO PIEZAS");
                        dtReporteMKT.Columns.Add("RENTABILIDAD");
                        dtReporteMKT.Columns.Add("LITROS NECESARIO");
                        dtReporteMKT.Columns.Add("PIEZAS NECESARIO");

                        //columnStyle = true;
                        columnStyleHtml = string.Empty;

                        contRegPDF = 0;
                        primerPaginaMercadotecnia = true;
                        contRegMercadotecnia = 0;
                        ListHtmlMercadotecnia = new List<string>();

                        ListReporteMKTOrden = new List<ReporteMKT>();

                        ListReporteMKTOrden.AddRange(ListReporteMKT.Where(m => m.Descripcion.ToUpper().Trim() != "TOTAL CAMPAÑA").
                            OrderBy(n => n.Alcance).ThenBy(n => n.Descripcion).ToList());

                        ReporteMKT totalReporteMKT = ListReporteMKT.Where(m => m.Descripcion.ToUpper().Trim() == "TOTAL CAMPAÑA").FirstOrDefault();

                        if (totalReporteMKT != null)
                        {
                            totalReporteMKT.Alcance = string.Empty;
                            //ListReporteMKTOrden.Add(new ReporteMKT());
                            ListReporteMKTOrden.Add(totalReporteMKT);
                        }

                        ReporteMKT reporteMKTAnterior = new ReporteMKT();

                        foreach (ReporteMKT reporteMKT in ListReporteMKTOrden)
                        {
                            //if (columnStyle)
                            //{
                            //    columnStyleHtml = "style='background-color:#ffffff'";
                            //    columnStyle = !columnStyle;
                            //}
                            //else
                            //{
                            //    columnStyleHtml = "style='background-color:#CEE3F6'";
                            //    columnStyle = !columnStyle;
                            //}

                            /////////////////////////////
                            //IDENTIFICAR SI ES EL ULTIMO REGISTRO
                            if((reporteMKTAnterior.Alcance != null && reporteMKT.Alcance != reporteMKTAnterior.Alcance) || ListReporteMKTOrden.Last() == reporteMKT)
                            {
                                //INSERTAR ESPACIO EN BLANCO HTML
                                html.AppendLine("<tr style='background-color:#E5E8E8;'" + columnStyleHtml + ">");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("<td style='background-color:#E5E8E8; border: 0.5px solid black;'>&nbsp;&nbsp;</td>");
                                html.AppendLine("</tr>");

                                //INSERTAR ESPACION EN BLANCO DATATABLE
                                dtReporteMKT.Rows.Add(string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty);
                            }

                            //PDF
                            html.AppendLine("<tr " + columnStyleHtml + ">");
                            //AGREGAR CONTENIDO DE LAS CELDAS
                            html.AppendLine("<td style='border: 0.5px solid black;'>" + reporteMKT.Alcance + "</td>");
                            html.AppendLine("<td style='border: 0.5px solid black;'>" + reporteMKT.Articulo + "</td>");
                            html.AppendLine("<td style='border: 0.5px solid black;'>" + reporteMKT.Descripcion + "</td>");
                            html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteMKT.LitrosActualCC.ToString("#,##0") + "</td>");
                            html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteMKT.PiezasActualCC.ToString("#,##0") + "</td>");
                            html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteMKT.ImporteCPRecioLitros.ToString("#,##0.00") + "</td>");
                            html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteMKT.ImporteCPrecioPiezas.ToString("#,##0.00") + "</td>");

                            if (string.IsNullOrEmpty(reporteMKT.Rentabilidad))
                            {
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'></td>");
                            }
                            else if (reporteMKT.Rentabilidad.ToUpper() == etiquetaRentable)
                            {
                                html.AppendLine("<td style='background-color:#ABEBC6; color:#0E6655; text-align: right; border: 0.5px solid black;'>" + reporteMKT.Rentabilidad.ToString() + "</td>");
                            }
                            else if (reporteMKT.Rentabilidad.ToUpper() == etiqueteNoRentable)
                            {
                                html.AppendLine("<td style='background-color:#F5B7B1; color:#922B21; text-align: right; border: 0.5px solid black;'>" + reporteMKT.Rentabilidad.ToString() + "</td>");
                            }
                            else if (reporteMKT.Rentabilidad.ToUpper() == etiquetaRevisar)
                            {
                                html.AppendLine("<td style='background-color:#F9E79F; color:#B7950B; text-align: right; border: 0.5px solid black;'>" + reporteMKT.Rentabilidad.ToString() + "</td>");
                            }
                            else
                            {
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteMKT.Rentabilidad + "</td>");
                            }

                            html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteMKT.Litros.ToString("#,##0") + "</td>");
                            html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteMKT.Piezas.ToString("#,##0") + "</td>");
                            //FIN CONTENIDO CELDAS
                            html.AppendLine("</tr>");

                            if (primerPaginaMercadotecnia && contRegMercadotecnia > RegPrimerPaginaMercadotecnia)
                            {
                                html.AppendLine(finDatosHtml.ToString());
                                html.AppendLine(finHtml.ToString());

                                ListHtmlMercadotecnia.Add(html.ToString());

                                html = new StringBuilder();

                                html.AppendLine(inicioHtml.ToString());
                                html.AppendLine(inicioDatosHtml.ToString());

                                primerPaginaMercadotecnia = false;

                                contRegMercadotecnia = 0;
                            }
                            else if(contRegMercadotecnia > RegSegundoPaginaMercadotecnia)
                            {
                                html.AppendLine(finDatosHtml.ToString());
                                html.AppendLine(finHtml.ToString());

                                ListHtmlMercadotecnia.Add(html.ToString());

                                html = new StringBuilder();

                                html.AppendLine(inicioHtml.ToString());
                                html.AppendLine(inicioDatosHtml.ToString());

                                primerPaginaMercadotecnia = false;

                                contRegMercadotecnia = 0;
                            }

                            contRegMercadotecnia++;

                            //contRegPDF = contRegPDF + 1;

                            //EXCEL
                            dtReporteMKT.Rows.Add(reporteMKT.Alcance == null ? string.Empty : reporteMKT.Alcance,
                                                reporteMKT.Articulo == null ? string.Empty : reporteMKT.Articulo,
                                                reporteMKT.Descripcion == null ? string.Empty : reporteMKT.Descripcion,
                                                reporteMKT.LitrosActualCC.ToString("#,##0"),
                                                reporteMKT.PiezasActualCC.ToString("#,##0"),
                                                reporteMKT.ImporteCPRecioLitros.ToString("#,##0.00"),
                                                reporteMKT.ImporteCPrecioPiezas.ToString("#,##0.00"),
                                                reporteMKT.Rentabilidad == null ? string.Empty : reporteMKT.Rentabilidad,
                                                reporteMKT.Litros.ToString("#,##0.00"),
                                                reporteMKT.Piezas.ToString("#,##0.00"));

                            reporteMKTAnterior = reporteMKT;

                        }

                        //html.AppendLine("</tbody>");
                        //html.AppendLine("</table>");

                        html.AppendLine(finDatosHtml.ToString());
                        html.AppendLine(finHtml.ToString());

                        ListHtmlMercadotecnia.Add(html.ToString());

                        //html.AppendLine("</div>");
                        //html.AppendLine(ImgHtmlFooter);
                        //html.AppendLine("</div>");


                        //CREAR ARCHIVO PDF
                        //nombreArchivo = Guid.NewGuid().ToString() + ".pdf";

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                               ConfigurationManager.AppSettings["DirectorioReporte"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathArchivo = parametro.Valor;
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


                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["WidthPDF"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            widthPDF = Convert.ToDouble(parametro.Valor);
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["HeigthPDF"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            heigthPDF = Convert.ToDouble(parametro.Valor);
                        }


                        //var Renderer = new IronPdf.HtmlToPdf();

                        //HtmlHeaderFooter htmlFooter = new HtmlHeaderFooter();
                        //htmlFooter.HtmlFragment = ImgHtmlFooter;
                        //Renderer.PrintOptions.Footer = htmlFooter;

                        //HtmlHeaderFooter htmlHeader = new HtmlHeaderFooter();
                        //htmlHeader.HtmlFragment = ImgHtmlHeader;
                        //Renderer.PrintOptions.Header = htmlHeader;

                        //Renderer.PrintOptions.SetCustomPaperSizeinMilimeters(widthPDF, heigthPDF);
                        //Renderer.PrintOptions.Header.DrawDividerLine = true;
                        //var PDF = Renderer.RenderHtmlAsPdf(html.ToString());
                        //PDF.SaveAs(pathArchivoCompleto);

                        //htmlReporteMercadotecnia = html.ToString();


                        ////GENERAR PDF CON DLL SPIRE
                        //layoutMercadotecnia = new PdfHtmlLayoutFormat();
                        //layoutMercadotecnia.IsWaiting = true;

                        //pageMercadotecnia = new PdfPageSettings();
                        //pageMercadotecnia.Size = new SizeF(762, 986);  //Spire.Pdf.PdfPageSize.Letter;

                        //int left = 15;
                        //int rigth = 15;
                        //int top = 35;
                        //int bottom = 35;

                        //pageMercadotecnia.Margins = new Spire.Pdf.Graphics.PdfMargins(left, top, rigth, bottom);
                        ////pageMercadotecnia.

                        //docMercadotecnia = new PdfDocument();

                        //Thread thread = new Thread(() =>
                        //{
                        //    docMercadotecnia.LoadFromHTML(htmlReporteMercadotecnia, autoDetectPage, pageMercadotecnia, layoutMercadotecnia);
                        //});

                        //thread.SetApartmentState(ApartmentState.STA);

                        //thread.Start();

                        //thread.Join();

                        //docMercadotecnia.SaveToFile(pathArchivoCompleto);
                        //docMercadotecnia.Clone();

                        ///////////////////
                        ////AGREGAR HEADER Y FOOTER
                        //docMercadotecniaImg = new PdfDocument();

                        //Thread thread2 = new Thread(() =>
                        //{
                        //    docMercadotecniaImg.LoadFromFile(pathArchivoCompleto);
                        //});

                        //headerImage = PdfImage.FromFile(pathImgHeader);
                        //footerImage = PdfImage.FromFile(pathImgFooter);

                        //thread2.SetApartmentState(ApartmentState.STA);
                        //thread2.Start();
                        //thread2.Join();

                        //int xHeader = 5;
                        //int yHeader = 5;
                        //int xFooter = 45;
                        //int yFooter = 961;

                        //foreach (PdfPageBase pdfPage in docMercadotecniaImg.Pages)
                        //{
                        //    pdfPage.Canvas.DrawImage(headerImage, xHeader, yHeader, 712, 20);
                        //    pdfPage.Canvas.DrawImage(footerImage, xFooter, yFooter, 712, 20);
                        //}

                        //docMercadotecniaImg.SaveToFile(pathArchivoCompleto);
                        //docMercadotecniaImg.Close();


                        //REPORTE MKT
                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["UrlReporteMKT"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            url = HttpUtility.HtmlEncode(parametro.Valor);
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                               ConfigurationManager.AppSettings["NombreReporteMKT"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            nombreReporte = parametro.Valor + rentabilidad.ClaveCampania;
                        }


                        //NOMBRE REPORTE CARPETA TEMPORAL
                        pathArchivoCompletoMercadotecnia = Path.Combine(pathArchivo, nombreReporte + ".pdf");


                        //NOMBRE REPORTE SHARE POINT
                        urlCompletoMercadotecnia = Path.Combine(url, nombreReporte + ".pdf");

                        rentabilidadENTRes.UrlReporteMKTPDF = urlCompletoMercadotecnia;


                        var t = System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                ////////////////////
                                //GENERAR PDF MERCADOTECNIA
                                //PDFTool.PDF.GenerarPDF(htmlReporteMercadotecnia, pathImgHeader, pathImgFooter, pathArchivoCompletoMercadotecnia, 792, 670);

                                PDFTool.PDF.GenerarPDFMercadotecnia(ListHtmlMercadotecnia, pathImgHeader, pathImgFooter, pathArchivoCompletoMercadotecnia, 792, 670);

                                return "OK";
                            }
                            catch (Exception ex)
                            {
                                ArchivoLog.EscribirLog(null, "ERROR: Service - PDF Mercadotecnia, Source:" + ex.Source + ", Message:" + ex.Message);

                                return "ERROR: Service - PDF Mercadotecnia, Source:" + ex.Source + ", Message:" + ex.Message;
                            }
                        });

                        var c = t.ContinueWith((antecedent) =>
                        {
                            if (antecedent.Result == "OK")
                            {
                                using (WebClient client = new WebClient())
                                {
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                                    client.UploadFile(urlCompletoMercadotecnia, "PUT", pathArchivoCompletoMercadotecnia);

                                }
                            }

                        });


                        //CREAR ARCHIVO EXCEL
                        //dsReporteMKT.Tables.Add(dtReporteMKT);

                        //nombreArchivo = Guid.NewGuid().ToString() + ".xlsx";

                        //NOMBRE REPORTE CARPETA TEMPORAL
                        pathArchivoCompletoMercadotecniaEx = Path.Combine(pathArchivo, nombreReporte + ".xlsx");

                        //ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoCompleto, dsReporteMKT);


                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            excel.Workbook.Worksheets.Add("ReporteMKT");

                            var worksheet = excel.Workbook.Worksheets["ReporteMKT"];

                            worksheet.SetValue("A1", "Reporte Marketing");
                            worksheet.SetValue("A2", "Campaña:");
                            worksheet.SetValue("B2", nombreCampana);
                            worksheet.SetValue("A3", "Clave Campaña:");
                            worksheet.SetValue("B3", claveCampana);
                            worksheet.SetValue("A4", "Fecha Inicio SubCanal:");

                            if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal != "01/01/0001")
                            {
                                worksheet.SetValue("B4", fechaInicioSubCanal);
                            }
                            worksheet.SetValue("A5", "Fecha Fin SubCanal:");
                            if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal != "01/01/0001")
                            {
                                worksheet.SetValue("B5", fechaFinSubCanal);
                            }
                            worksheet.SetValue("A6", "Fecha Inicio Publico");
                            if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico != "01/01/0001")
                            {
                                worksheet.SetValue("B6", fechaInicioPublico);
                            }
                            worksheet.SetValue("A7", "Fecha Fin Publico");
                            if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico != "01/01/0001")
                            {
                                worksheet.SetValue("B7", fechaFinPublico);
                            }

                            worksheet.Cells["A1:B7"].Style.Font.Color.SetColor(1, 38, 130, 221);
                            worksheet.Cells["A1:B7"].Style.Font.Bold = true;
                            worksheet.Cells["A1:B7"].Style.Font.Size = 14;

                            worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteMKT.Columns.Count + 64) + "9"].LoadFromDataTable(dtReporteMKT, true);//, OfficeOpenXml.Table.TableStyles.Light2);

                            worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteMKT.Columns.Count + 64) + "9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteMKT.Columns.Count + 64) + "9"].Style.Fill.BackgroundColor.SetColor(1, 59, 137, 216);
                            worksheet.Cells["A9:" + Char.ConvertFromUtf32(dtReporteMKT.Columns.Count + 64) + "9"].Style.Font.Color.SetColor(Color.White);

                            worksheet.Cells["A9:J" + (dtReporteMKT.Rows.Count + 9)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A9:J" + (dtReporteMKT.Rows.Count + 9)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A9:J" + (dtReporteMKT.Rows.Count + 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A9:J" + (dtReporteMKT.Rows.Count + 9)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            worksheet.Cells["I9:J9"].Style.Fill.BackgroundColor.SetColor(1, 46, 139, 87);

                            if (dtReporteMKT.Rows.Count > 0)
                            {
                                worksheet.Cells["H10:H" + (dtReporteMKT.Rows.Count + 9)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            }

                            for (int m=0; m < dtReporteMKT.Rows.Count; m++)
                            {
                                DataRow dr = dtReporteMKT.Rows[m];

                                if(string.IsNullOrEmpty(dr["RENTABILIDAD"].ToString()))
                                {
                                    worksheet.Cells["H" + (m + 10)].Style.Fill.BackgroundColor.SetColor(Color.Transparent);
                                    //worksheet.Cells["G" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else if (dr["RENTABILIDAD"].ToString().ToUpper() == etiquetaRentable)
                                {
                                    worksheet.Cells["H" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 171, 235, 198);
                                    worksheet.Cells["H" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else if(dr["RENTABILIDAD"].ToString().ToUpper() == etiqueteNoRentable)
                                {
                                    worksheet.Cells["H" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 245, 183, 177);
                                    worksheet.Cells["H" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else if(dr["RENTABILIDAD"].ToString().ToUpper() == etiquetaRevisar)
                                {
                                    worksheet.Cells["H" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 249, 231, 159);
                                    worksheet.Cells["H" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else
                                {
                                    worksheet.Cells["H" + (m + 10)].Style.Fill.BackgroundColor.SetColor(Color.Transparent);
                                    //worksheet.Cells["G" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                            }

                            worksheet.Cells["A1:J" + (dtReporteMKT.Rows.Count + 9)].AutoFitColumns();

                            FileInfo excelFile = new FileInfo(pathArchivoCompletoMercadotecniaEx);
                            excel.SaveAs(excelFile);
                        }

                        //NOMBRE REPORTE SHARE POINT
                        urlCompletoMercadotecniaEx = Path.Combine(url, nombreReporte + ".xlsx");

                        using (WebClient client = new WebClient())
                        {
                            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                            client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                            client.UploadFile(urlCompletoMercadotecniaEx, "PUT", pathArchivoCompletoMercadotecniaEx);

                            rentabilidadENTRes.UrlReporteMKTExcel = urlCompletoMercadotecniaEx;
                        }

                    }

                    #endregion


                    #region REPORTE SKU

                    //REPORTE SKU
                    if (ListReporteSKU.Count > 0)
                    {
                        html = new StringBuilder();
                        inicioHtml = new StringBuilder();
                        inicioDatosHtml = new StringBuilder();
                        finDatosHtml = new StringBuilder();
                        finHtml = new StringBuilder();
                        tablaInicioHtml = new StringBuilder();
                        ListHtmlSKU = new List<string>();

                        //PDF
                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgComex"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgComex = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgPPG"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgPPG = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgHeader"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgHeader = parametro.Valor;
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["PathImgFooter"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathImgFooter = parametro.Valor;
                        }


                        //var pngBinaryDataComex = File.ReadAllBytes(pathImgComex);
                        //var ImgDataURIComex = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataComex);
                        //var ImgHtmlComex = String.Format("<img src='{0}' style='float:right;vertical-align:top;' width='120' height='50'>", ImgDataURIComex);

                        var pngBinaryDataHeaer = File.ReadAllBytes(pathImgHeader);
                        var ImgDataURIHeader = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataHeaer);
                        var pngBinaryDataFooter = File.ReadAllBytes(pathImgFooter);
                        var ImgDataURIFooter = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataFooter);

                        //var pngBinaryDataPPG = File.ReadAllBytes(pathImgPPG);
                        //var ImgDataURIPPG = @"data:image/png;base64," + Convert.ToBase64String(pngBinaryDataPPG);
                        //var ImgHtmlPPG = String.Format("<img src='{0}' style='float:left;vertical-align:top;' width='70' height='50'>", ImgDataURIPPG);

                        //var ImgHtmlHeader = String.Format("<div style='width: 100%; height: 50px; position: absolute; top: 1px;'><img src='{0}' width='100%' height='100%'></div>", ImgDataURIHeader);
                        ////var ImgHtmlFooter = String.Format("<div style='width: 100%; height: 50px; position: absolute; bottom: 1px;'><img src='{0}' width='100%' height='100%'></div>", ImgDataURIFooter);
                        //var ImgHtmlFooter = String.Format("<footer style='width: 100%; height: 50px; flow: static(footer);'><img src='{0}' width='100%' height='100%'></footer>", ImgDataURIFooter);

                        //html.AppendLine(ImgHtmlComex);
                        //inicioHtml.AppendLine("<div style='font-family: arial; position: relative; height: 100%;'>");
                        //html.AppendLine(ImgHtmlHeader);
                        //html.AppendLine(ImgHtmlPPG);
                        inicioHtml.AppendLine("<div style='font-family: arial; position: absolute;'>");
                        html.AppendLine(inicioHtml.ToString());

                        tablaInicioHtml.AppendLine("<center><h3>Reporte SKU</h3></center>");
                        //TABLA
                        tablaInicioHtml.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                        tablaInicioHtml.AppendLine("<tbody>");
                        //REGISTRO 1
                        tablaInicioHtml.AppendLine("<tr style='background-color:#ffffff;'>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Campaña</td>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + nombreCampana + "</td>");
                        tablaInicioHtml.AppendLine("</tr>");
                        //REGISTRO 2
                        tablaInicioHtml.AppendLine("<tr style='background-color:#CEE3F6;'>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Clave Campaña</td>");
                        tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + claveCampana + "</td>");
                        tablaInicioHtml.AppendLine("</tr>");
                        //REGISTRO 3
                        if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#ffffff;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Inicio SubCanal</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + fechaInicioSubCanal + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }
                        //REGISTRO 4
                        if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#CEE3F6;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Fin SubCanal</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + fechaFinSubCanal + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }
                        //REGISTRO 5
                        if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#ffffff;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Inicio Publico</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; font-weight:bold;'>" + fechaInicioPublico + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }
                        //REGISTRO 6
                        if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico.ToString() != "01/01/0001")
                        {
                            tablaInicioHtml.AppendLine("<tr style='background-color:#CEE3F6;'>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#3B89D8; font-size:13px; color:white; font-weight:bold;'>Fecha Fin Publico</td>");
                            tablaInicioHtml.AppendLine("<td style='border: 0.5px solid black; background-color:#CEE3F6; font-weight:bold;'>" + fechaFinPublico + "</td>");
                            tablaInicioHtml.AppendLine("</tr>");
                        }

                        tablaInicioHtml.AppendLine("</tbody>");
                        tablaInicioHtml.AppendLine("</table>");
                        tablaInicioHtml.AppendLine("<br />");

                        html.AppendLine(tablaInicioHtml.ToString());

                        inicioDatosHtml.AppendLine("<table border='0' cellspacing='0' cellpadding='0' style='border-collapse: collapse; border-spacing: 0px; border: 0.5px solid black; font-family:arial; font-size:15px;'>");
                        inicioDatosHtml.AppendLine("<thead style='background-color:#3B89D8; font-size:12px; color:white;'>");

                        //ENCABEZADO DE GRUPOS
                        inicioDatosHtml.AppendLine("<tr>");

                        inicioDatosHtml.AppendLine("<th colspan='4' style='background-color:#3B89D8; border: 0.5px solid black;'></th>");
                        inicioDatosHtml.AppendLine("<th colspan='2' style='background-color:#3B89D8; border: 0.5px solid black;'>ACTUAL</th>");
                        inicioDatosHtml.AppendLine("<th colspan='2' style='background-color:#3B89D8; border: 0.5px solid black;'>PROMOCION</th>");
                        inicioDatosHtml.AppendLine("<th colspan='8' style='background-color:#2E8B57; border: 0.5px solid black;'>VENTA REAL AÑO ANTERIOR</th>");
                        inicioDatosHtml.AppendLine("<th colspan='8' style='background-color:#3B89D8; border: 0.5px solid black;'>PRESUPUESTO SIN PROMOCIÓN</th>");
                        inicioDatosHtml.AppendLine("<th colspan='8' style='background-color:#3B89D8; border: 0.5px solid black;'>PRESUPUESTO CON PROMOCION</th>");
                        inicioDatosHtml.AppendLine("<th colspan='1' style='background-color:#3B89D8; border: 0.5px solid black;'></th>");

                        inicioDatosHtml.AppendLine("</tr>");


                        inicioDatosHtml.AppendLine("<tr>");

                        //AGREGAR COLUMNAS
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>ALCANCE</th>");
                        inicioDatosHtml.AppendLine("<th style='width:60px;background-color:#3B89D8; border: 0.5px solid black;'>SKU</th>");
                        inicioDatosHtml.AppendLine("<th style='width:140px;background-color:#3B89D8; border: 0.5px solid black;'>PRODUCTO</th>");
                        inicioDatosHtml.AppendLine("<th style='background-color:#3B89D8; border: 0.5px solid black;'>CAPACIDAD</th>");                       

                        //ACTUAL
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>PRECIO CONC.</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>FACTOR UTILIDAD BRUTA</th>");

                        //PROMOCION
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>PRECIO CONC.</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>FACTOR UTILIDAD</th>");

                        //VENTA REAL AÑO ANTERIOR
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>LITROS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>PIEZAS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>COSTO MP + ENV + FAB</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>GASTO PLANTA</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>GASTO KROMA</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>UTILIDAD BRUTA COMEX</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>IMPORTE C/PRECIO CONC. LITROS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#2E8B57; border: 0.5px solid black;'>IMPORTE C/PRECIO CONC. PIEZAS</th>");

                        //PRESUPUESTO SIN PROMOCION
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>LITROS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>PIEZAS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>COSTO MP + ENV + FAB</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>GASTO PLANTA</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>GASTO KROMA</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>UTILIDAD BRUTA COMEX</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>IMPORTE C/PRECIO CONC. LITROS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>IMPORTE C/PRECIO CONC. PIEZAS</th>");

                        //PRESUPUESTO CON PROMOCION
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>LITROS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>PIEZAS</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>COSTO MP + ENV + FAB</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>GASTO PLANTA</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>GASTO KROMA</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>UTILIDAD BRUTA COMEX</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>IMPORTE C/PRECIO CONC. LITRO</th>");
                        inicioDatosHtml.AppendLine("<th style='width:80px; background-color:#3B89D8; border: 0.5px solid black;'>IMPORTE C/PRECIO CONC. PIEZAS</th>");

                        //COMENTARIO
                        inicioDatosHtml.AppendLine("<th style='background-color:#3B89D8; border: 0.5px solid black;'>COMENTARIO</th>");


                        //FIN COLUMNAS
                        inicioDatosHtml.AppendLine("</tr>");
                        inicioDatosHtml.AppendLine("</thead>");

                        inicioDatosHtml.AppendLine("<tbody style='font-size:11px;'>");

                        html.AppendLine(inicioDatosHtml.ToString());

                        finDatosHtml.AppendLine("</tbody>");
                        finDatosHtml.AppendLine("</table>");
                        finHtml.AppendLine("</div>");


                        //EXCEL
                        //dtReporteSKU.Columns.Add("CLAVE");
                        dtReporteSKU.Columns.Add("SEGMENTO");
                        dtReporteSKU.Columns.Add("SKU");
                        dtReporteSKU.Columns.Add("PRODUCTO");
                        dtReporteSKU.Columns.Add("CAPACIDAD");      

                        //COLUMNA VERDE
                        dtReporteSKU.Columns.Add("COSTO M.P + ENV + CONV AÑO ANTERIOR");
                        dtReporteSKU.Columns.Add("GASTO PLANTA AÑO ANTERIOR");
                        dtReporteSKU.Columns.Add("GASTO KROMA AÑO ANTERIOR");

                        //COLUMNA AZUL
                        dtReporteSKU.Columns.Add("COSTO M.P + ENV + CONV ACTUAL");
                        dtReporteSKU.Columns.Add("GASTO PLANTA ACTUAL");
                        dtReporteSKU.Columns.Add("GASTO KROMA ACTUAL");

                        //COLUMNA VERDE
                        //9
                        dtReporteSKU.Columns.Add("PRECIO CONC. ANTERIOR");
                        dtReporteSKU.Columns.Add("PRECIO ANT CON PROMOCIÓN");
                        dtReporteSKU.Columns.Add("INVERSIÓN TINTA");                  
                        dtReporteSKU.Columns.Add("PRECIO PÚBLICO ANTERIOR CON PROMOCIÓN");

                        //COLUMNA AZUL
                        dtReporteSKU.Columns.Add("PRECIO CONC. A").Caption = "PRECIO CONC.";
                        dtReporteSKU.Columns.Add("FACTOR UTILIDAD A").Caption = "FACTOR UTILIDAD";
                        dtReporteSKU.Columns.Add("PRECIO PUBLICO A").Caption = "PRECIO PUBLICO";
                        dtReporteSKU.Columns.Add("INVERSIÓN TINTA A").Caption = "INVERSIÓN TINTA";
                        dtReporteSKU.Columns.Add("% MARGEN CONC. A").Caption = "% MARGEN CONC.";

                        dtReporteSKU.Columns.Add("PRECIO CONC. P").Caption = "PRECIO CONC.";
                        dtReporteSKU.Columns.Add("FACTOR UTILIDAD P").Caption = "FACTOR UTILIDAD";
                        dtReporteSKU.Columns.Add("PRECIO PUBLICO P").Caption = "PRECIO PUBLICO";
                        dtReporteSKU.Columns.Add("INVERSION TINTA P").Caption = "INVERSION TINTA";
                        dtReporteSKU.Columns.Add("% MARGEN CONC. P").Caption = "% MARGEN CONC.";

                        //COLUMNA VERDE
                        //24
                        dtReporteSKU.Columns.Add("LITROS VR").Caption = "LITROS";
                        dtReporteSKU.Columns.Add("PIEZAS 1 VR").Caption = "PIEZAS";
                        dtReporteSKU.Columns.Add("PIEZAS 2 VR").Caption = "PIEZAS";
                        dtReporteSKU.Columns.Add("COSTO MP + ENV + FAB VR").Caption = "COSTO MP + ENV + FAB";
                        dtReporteSKU.Columns.Add("GASTO PLANTA VR").Caption = "GASTO PLANTA";
                        dtReporteSKU.Columns.Add("GASTO KROMA VR").Caption = "GASTO KROMA";
                        dtReporteSKU.Columns.Add("UTILIDAD BRUTA COMEX VR").Caption = "UTILIDAD BRUTA COMEX";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC. LITROS VR").Caption = "IMPORTE C/PRECIO CONC. LITROS";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC. PIEZAS VR").Caption = "IMPORTE C/PRECIO CONC. PIEZAS";
                        dtReporteSKU.Columns.Add("UTILIDAD CONC. VR").Caption = "UTILIDAD CONC.";
                        dtReporteSKU.Columns.Add("INVERSIÓN TINTA VR").Caption = "INVERSIÓN TINTA";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO PUBLICO VR").Caption = "IMPORTE C/PRECIO PUBLICO";

                        //COLUMNA AZUL
                        //SIN PRESUPUESTO
                        dtReporteSKU.Columns.Add("LITROS PS").Caption = "LITROS";
                        dtReporteSKU.Columns.Add("PIEZAS 1 PS").Caption = "PIEZAS";
                        dtReporteSKU.Columns.Add("PIEZAS 2 PS").Caption = "PIEZAS";
                        dtReporteSKU.Columns.Add("COSTO MP + ENV + FAB PS").Caption = "COSTO MP + ENV + FAB";
                        dtReporteSKU.Columns.Add("GASTO PLANTA PS").Caption = "GASTO PLANTA";
                        dtReporteSKU.Columns.Add("GASTO KROMA PS").Caption = "GASTO KROMA";
                        dtReporteSKU.Columns.Add("UTILIDAD BRUTA COMEX PS").Caption = "UTILIDAD BRUTA COMEX";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC. LITROS PS").Caption = "IMPORTE C/PRECIO CONC. LITROS";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC. PIEZAS PS").Caption = "IMPORTE C/PRECIO CONC. PIEZAS";
                        dtReporteSKU.Columns.Add("UTILIDAD CONC. PS").Caption = "UTILIDAD CONC.";
                        dtReporteSKU.Columns.Add("INVERSIÓN TINTA PS").Caption = "INVERSIÓN TINTA";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO PUBLICO PS").Caption = "IMPORTE C/PRECIO PUBLICO";

                        //CON PRESUPUESTO
                        //48
                        dtReporteSKU.Columns.Add("LITROS PC").Caption = "LITROS";
                        dtReporteSKU.Columns.Add("PIEZAS 1 PC").Caption = "PIEZAS";
                        dtReporteSKU.Columns.Add("PIEZAS 2 PC").Caption = "PIEZAS ";
                        dtReporteSKU.Columns.Add("COSTO MP + ENV + FAB PC").Caption = "COSTO MP + ENV + FAB";
                        dtReporteSKU.Columns.Add("GASTO PLANTA PC").Caption = "GASTO PLANTA";
                        dtReporteSKU.Columns.Add("GASTO KROMA PC").Caption = "GASTO KROMA";
                        dtReporteSKU.Columns.Add("UTILIDAD BRUTA COMEX PC").Caption = "UTILIDAD BRUTA COMEX";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC. LITROS PC").Caption = "IMPORTE C/PRECIO CONC. LITROS";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC. PIEZAS PC").Caption = "IMPORTE C/PRECIO CONC. PIEZAS";
                        dtReporteSKU.Columns.Add("UTILIDAD CONC. PC").Caption = "UTILIDAD CONC.";
                        dtReporteSKU.Columns.Add("INVERSIÓN TINTA PC").Caption = "INVERSIÓN TINTA";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO PUBLICO PC").Caption = "IMPORTE C/PRECIO PUBLICO";

                        dtReporteSKU.Columns.Add("LITROS NECESARIO").Caption = "LITROS";
                        dtReporteSKU.Columns.Add("PIEZAS NECESARIO").Caption = "PIEZAS";
                        dtReporteSKU.Columns.Add("PIEZAS").Caption = "PIEZAS";
                        dtReporteSKU.Columns.Add("COSTO MP + ENV NECESARIO").Caption = "COSTO MP + ENV";
                        dtReporteSKU.Columns.Add("COSTO PLANTA NECESARIO").Caption = "COSTO PLANTA";
                        dtReporteSKU.Columns.Add("GASTO KROMA NECESARIO").Caption = "GASTO KROMA";
                        dtReporteSKU.Columns.Add("UTILIDAD BRUTA");
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC LITROS NECESARIO").Caption = "IMPORTE C/PRECIO CONC LITROS";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO CONC PIEZAS NECEARIO").Caption = "IMPORTE C/PRECIO CONC PIEZAS";
                        dtReporteSKU.Columns.Add("UTILIDAD CONC");
                        dtReporteSKU.Columns.Add("TINTA NECESARIO").Caption = "TINTA";
                        dtReporteSKU.Columns.Add("IMPORTE C/PRECIO PUBLICO NECESARIO").Caption = "IMPORTE C/PRECIO PUBLICO";

                        dtReporteSKU.Columns.Add("ANTERIOR");
                        dtReporteSKU.Columns.Add("SIN CAMPAÑA");
                        dtReporteSKU.Columns.Add("CON PROMOCION");
                        dtReporteSKU.Columns.Add("RENTABILIDAD");

                        dtReporteSKU.Columns.Add("COMENTARIO");

                        //64

                        //columnStyle = true;
                        columnStyleHtml = string.Empty;

                        contRegPDF = 0;

                        ///////////////////////////////
                        //////////////////////////////
                        //ORDENAR SKU

                        //var uno = ListReporteSKU.GroupBy(n => new { n.Alcance, n.Articulo }).Select(k => new { articulo = k.Key, cantidad = k.Count() }).Where(m => m.cantidad == 1).ToList();

                        //List<ReporteSKU> ListReporteSKUSinGrupo = (from skuGroup in ListReporteSKU.GroupBy(n => n.Articulo).Select(k => new { articulo = k.Key, cantidad = k.Count() }).Where(m => m.cantidad == 1).ToList()
                        //                                           from sku in ListReporteSKU.ToList()
                        //                                           where skuGroup.articulo == sku.Articulo
                        //                                           select sku).ToList();

                        //ListReporteSKUSinGrupo.ForEach(n =>
                        //{
                        //    n.Agrupador = 1;
                        //    n.SubArticulo = n.Articulo;
                        //    n.SubDescripcion = n.Descripcion;
                        //});

                        //List<ReporteSKU> ListReporteSKUConGrupo = new List<ReporteSKU>();

                        //ListReporteSKUConGrupo.AddRange(ListReporteSKU);

                        //ListReporteSKUConGrupo = ListReporteSKUConGrupo.Except(ListReporteSKUSinGrupo).ToList();

                        int contGrupoSKU = 0;
                        int subContGrupoSKU = 0;
                        int contOrdenSKU = 0;
                        string antSKU = string.Empty;
                        string antDescripcion = string.Empty;
                        string antAlcance = string.Empty;

                        int indiceMaxAgrupador = -1;
                        bool esAgrupador = true;

                        ListReporteSKU.ForEach(n =>
                        {
                            if(esAgrupador && ListReporteSKU.Count >= contOrdenSKU + 2 && ListReporteSKU[contOrdenSKU + 2].Descripcion.ToUpper().Trim() == etiquetaSubtotal)
                            {
                                subContGrupoSKU = 0;
                                contGrupoSKU++;

                                antSKU = n.Articulo;                          
                                antAlcance = n.Alcance;
                                antDescripcion = n.Descripcion;

                                n.SubArticulo = antSKU;
                                n.SubDescripcion = antDescripcion;
                                n.SubAlcance = antAlcance;

                                n.Agrupador = contGrupoSKU;
                                n.SubAgrupador = subContGrupoSKU;

                                indiceMaxAgrupador = contOrdenSKU + 2;

                                esAgrupador = false;                         
                            }
                            else
                            {
                                if (indiceMaxAgrupador >= contOrdenSKU)
                                {
                                    subContGrupoSKU++;

                                    n.SubArticulo = antSKU;
                                    n.SubDescripcion = antDescripcion;
                                    n.SubAlcance = antAlcance;

                                    n.Agrupador = contGrupoSKU;
                                    n.SubAgrupador = subContGrupoSKU;

                                    if (indiceMaxAgrupador == contOrdenSKU)
                                    {
                                        esAgrupador = true;
                                    }
                                    else
                                    {
                                        esAgrupador = false;
                                    }
                                }
                                else
                                {
                                    subContGrupoSKU = 0;
                                    contGrupoSKU++;
                                    indiceMaxAgrupador = -1;

                                    antSKU = n.Articulo;
                                    antAlcance = n.Alcance;
                                    antDescripcion = n.Descripcion;

                                    n.SubArticulo = antSKU;
                                    n.SubDescripcion = antDescripcion;
                                    n.SubAlcance = antAlcance;

                                    n.Agrupador = contGrupoSKU;
                                    n.SubAgrupador = subContGrupoSKU;

                                    esAgrupador = true;
                                }
                            }

                            contOrdenSKU++;
                        });


                        //ListReporteSKU.ForEach(n =>
                        //{
                        //    if (n.Descripcion.Trim().ToUpper() != "SUBTOTAL" 
                        //        && ((string.IsNullOrEmpty(n.Articulo) || string.IsNullOrEmpty(antSKU))
                        //        || (n.Articulo.Substring(5) != antSKU.Substring(5)))
                        //        )
                        //    {
                        //        contGrupoSKU = 0;
                        //    }

                        //    contGrupoSKU++;

                        //    if (string.IsNullOrEmpty(n.Articulo))
                        //    {
                        //        n.SubArticulo = antSKU;
                        //        n.SubDescripcion = antDescripcion;
                        //        n.SubAlcance = antAlcance;
                        //    }
                        //    else
                        //    {
                        //        n.SubArticulo = n.Articulo;
                        //        n.SubDescripcion = n.Descripcion;
                        //        n.SubAlcance = n.Alcance;
                        //    }

                        //    n.Agrupador = contGrupoSKU;

                        //    antSKU = n.Articulo;
                        //    antDescripcion = n.Descripcion;
                        //    antAlcance = n.Alcance;

                        //    contOrdenSKU++;

                        //    n.Orden = contOrdenSKU;
                        //});

                        ListReporteSKU = ListReporteSKU.OrderBy(n => n.SubAlcance).ThenBy(n => n.SubDescripcion).ThenBy(n => n.SubArticulo)
                                                            .ThenBy(n => n.Agrupador).ThenBy(n => n.SubAgrupador).ToList();

                        ReporteSKU reporteSKUAnterior = new ReporteSKU();

                        foreach (ReporteSKU reporteSKU in ListReporteSKU)//.OrderBy(n => n.Descripcion))
                        {
                            //if (columnStyle)
                            //{
                            //    columnStyleHtml = "style='background-color:#ffffff'";
                            //    columnStyle = !columnStyle;
                            //}
                            //else
                            //{
                            //    columnStyleHtml = "style='background-color:#CEE3F6'";
                            //    columnStyle = !columnStyle;
                            //}

                            #region AGREGAR SALTO ADICIONAL

                            if (reporteSKUAnterior.SubAlcance != null && reporteSKUAnterior.SubAlcance.ToUpper().Trim() != reporteSKU.SubAlcance.ToUpper().Trim())
                            {
                                //REGISTRO
                                //PDF
                                html.AppendLine("<tr style='background-color:#E5E8E8;'" + columnStyleHtml + ">");

                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //ACTUAL
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //VENTA REAL AÑO ANTERIOR
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PRESUPUESTO SIN PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PRESUPUESTO CON PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //COMENTARIO
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //FIN CONTENIDO CELDAS
                                html.AppendLine("</tr>");


                                //contRegPDF = contRegPDF + 1;


                                //REGISTRO
                                //EXCEL
                                dtReporteSKU.Rows.Add(
                                                //1    
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                //11
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //21
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //31
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //41
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //NUEVOS
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty);
                            }

                            #endregion

                            if (reporteSKU.Descripcion.ToUpper().Trim() != etiquetaSubtotal)
                            {
                                //PDF
                                html.AppendLine("<tr " + columnStyleHtml + ">");

                                html.AppendLine("<td style='border: 0.5px solid black;'>" + reporteSKU.Alcance + "</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>" + reporteSKU.Articulo + "</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>" + reporteSKU.Descripcion + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.Capacidad + "</td>");
                                
                                //ACTUAL
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.PrecioConc.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.FactorUtilidadBruto.ToString("#,##0.00") + "</td>");

                                //PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.PrecioConce.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.FactorUtilidadCC.ToString("#,##0.00") + "</td>");

                                //VENTA REAL AÑO ANTERIOR
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.LitrosAnterior.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.PiezasVacioAnt.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.CostoMPEnvFab.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.GastoPlantaAnt.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.GastoKromaAnter.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.UtilidadBrutaComexAnt.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.ImporteCPrecioConcLitro.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.ImporteCPrecioConcPiezas.ToString("#,##0.00") + "</td>");

                                //PRESUPUESTO SIN PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.LitrosActualSC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.LitrosVacioSC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.CostoMPEnvFabrica.ToString("#,##0.00%") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.GastoPlantaActualSC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.GastoKromaActualSC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.UtilidadBrutaComexSC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.Importe.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.ImporteCPrecioConcPiezasSC.ToString("#,##0.00%") + "</td>");

                                //PRESUPUESTO CON PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.LitrosActualCC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.LitrosVacioCC.ToString("#,##0") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.CostoMPEnvFabCC.ToString("#,##0") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.GastoPlantaActualCC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.GastoKromaActualCC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.UtilidadBrutaComexCC.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.ImporteCPrecioConcLitros.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.ImporteCPrecioConcCCPiezas.ToString("#,##0.00") + "</td>");

                                //COMENTARIO
                                if (string.IsNullOrEmpty(reporteSKU.Comentario))
                                {
                                    html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'></td>");
                                }
                                else if (reporteSKU.Comentario.ToUpper() == etiquetaRentable)
                                {
                                    html.AppendLine("<td style='background-color:#ABEBC6; color:#0E6655; text-align: right; border: 0.5px solid black;'>" + reporteSKU.Comentario + "</td>");
                                }
                                else if (reporteSKU.Comentario.ToUpper() == etiqueteNoRentable)
                                {
                                    html.AppendLine("<td style='background-color:#F5B7B1; color:#922B21; text-align: right; border: 0.5px solid black;'>" + reporteSKU.Comentario + "</td>");
                                }
                                else if (reporteSKU.Comentario.ToUpper() == etiquetaRevisar)
                                {
                                    html.AppendLine("<td style='background-color:#F9E79F; color:#B7950B; text-align: right; border: 0.5px solid black;'>" + reporteSKU.Comentario + "</td>");
                                }
                                else
                                {
                                    html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.Comentario + "</td>");
                                }

                                //FIN CONTENIDO CELDAS
                                html.AppendLine("</tr>");

                                //contRegPDF = contRegPDF + 1;


                                //EXCEL
                                dtReporteSKU.Rows.Add(//string.IsNullOrEmpty(reporteSKU.ClaveCampania) ? string.Empty : reporteSKU.ClaveCampania,
                                                      //1     
                                                    string.IsNullOrEmpty(reporteSKU.Alcance) ? string.Empty : reporteSKU.Alcance,
                                                    string.IsNullOrEmpty(reporteSKU.Articulo) ? string.Empty : reporteSKU.Articulo,
                                                    string.IsNullOrEmpty(reporteSKU.Descripcion) ? string.Empty : reporteSKU.Descripcion,
                                                    string.IsNullOrEmpty(reporteSKU.Capacidad) ? string.Empty : reporteSKU.Capacidad,
                                                    
                                                    reporteSKU.CostoMPAnterior.ToString("#,##0.00"),
                                                    reporteSKU.GastoPlantaAnterior.ToString("#,##0.00"),
                                                    reporteSKU.GastoKromaAñoAnterior.ToString("#,##0.00"),
                                                    reporteSKU.CostoMPEnvConvAñoActual.ToString("#,##0.00"),
                                                    reporteSKU.CostoPlantaAñoActual.ToString("#,##0.00"),
                                                    reporteSKU.GastoKromaActual.ToString("#,##0.00"),
                                                    reporteSKU.PrecioConcAnterior.ToString("#,##0.00"),

                                                    //11
                                                    reporteSKU.PrecioAntConPromocion.ToString("#,##0.00"),
                                                    reporteSKU.InversionTintaAnt.ToString("#,##0.00"),
                                                    reporteSKU.PrecioPublicoAntCP.ToString("#,##0.00"),
                                                    reporteSKU.PrecioConc.ToString("#,##0.00"),
                                                    reporteSKU.FactorUtilidadBruto.ToString("#,##0.00"),
                                                    reporteSKU.PrecioPublicoSC.ToString("#,##0.00"),
                                                    reporteSKU.InversionTintaActSC.ToString("#,##0.00"),
                                                    reporteSKU.PorcMargenConcSC.ToString("#,##0.00"),
                                                    reporteSKU.PrecioConce.ToString("#,##0.00"),
                                                    reporteSKU.FactorUtilidadCC.ToString("#,##0.00"),

                                                    //21
                                                    reporteSKU.PrecioPublicoCC.ToString("#,##0.00"),
                                                    reporteSKU.InversionTintaActCC.ToString("#,##0.00"),
                                                    reporteSKU.PorcMargenConcCC.ToString("#,##0.00"),
                                                    reporteSKU.LitrosAnterior.ToString("#,##0.00"),
                                                    reporteSKU.PiezasVacioAnt.ToString("#,##0.00"),
                                                    reporteSKU.PiezasAnt.ToString("#,##0.00"),
                                                    reporteSKU.CostoMPEnvFab.ToString("#,##0.00"),
                                                    reporteSKU.GastoPlantaAnt.ToString("#,##0.00"),
                                                    reporteSKU.GastoKromaAnter.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadBrutaComexAnt.ToString("#,##0.00"),

                                                    //31
                                                    reporteSKU.ImporteCPrecioConcLitro.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioConcPiezas.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadConcAnterior.ToString("#,##0.00"),
                                                    reporteSKU.InversionTintaAnterior.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioPublicoAnt.ToString("#,##0.00"),
                                                    reporteSKU.LitrosActualSC.ToString("#,##0.00"),
                                                    reporteSKU.LitrosVacioSC.ToString("#,##0.00"),
                                                    reporteSKU.PiezasActualSC.ToString("#,##0.00"),
                                                    reporteSKU.CostoMPEnvFabrica.ToString("#,##0.00"),
                                                    reporteSKU.GastoPlantaActualSC.ToString("#,##0.00"),

                                                    //41
                                                    reporteSKU.GastoKromaActualSC.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadBrutaComexSC.ToString("#,##0.00"),
                                                    reporteSKU.Importe.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioConcPiezasSC.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadConcSC.ToString("#,##0.00"),
                                                    reporteSKU.InversionTintaSC.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioPublicoSC.ToString("#,##0.00"),
                                                    reporteSKU.LitrosActualCC.ToString("#,##0.00"),
                                                    reporteSKU.LitrosVacioCC.ToString("#,##0.00"),
                                                    reporteSKU.PiezasActualCC.ToString("#,##0.00"),
                                                    reporteSKU.CostoMPEnvFabCC.ToString("#,##0.00"),
                                                    reporteSKU.GastoPlantaActualCC.ToString("#,##0.00"),
                                                    reporteSKU.GastoKromaActualCC.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadBrutaComexCC.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioConcLitros.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioConcCCPiezas.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadConcCC.ToString("#,##0.00"),
                                                    reporteSKU.InversionTintaCC.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioPublicoCC.ToString("#,##0.00"),

                                                    //NUEVOS
                                                    reporteSKU.LitrosNecesarios.ToString("#,##0.00"),
                                                    reporteSKU.PiezasNecesario.ToString("#,##0.00"),
                                                    reporteSKU.Piezas.ToString("#,##0.00"),
                                                    reporteSKU.CostoMPEnvNecesario.ToString("#,##0.00"),
                                                    reporteSKU.CostoPlantaNecesario.ToString("#,##0.00"),
                                                    reporteSKU.GastoKromaNecesario.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadBruta.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioConcLitrosNecesario.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioConcPiezasNecesario.ToString("#,##0.00"),
                                                    reporteSKU.UtilidadConc.ToString("#,##0.00"),
                                                    reporteSKU.TintaNecesario.ToString("#,##0.00"),
                                                    reporteSKU.ImporteCPrecioPublicoNecesario.ToString("#,##0.00"),

                                                    reporteSKU.Anterior.ToString("#,##0.00"),
                                                    reporteSKU.SinCampania.ToString("#,##0.00"),
                                                    reporteSKU.ConPromocion.ToString("#,##0.00"),
                                                    reporteSKU.Rentabilidad.ToString("#,##0.00"),

                                                    string.IsNullOrEmpty(reporteSKU.Comentario) ? string.Empty : reporteSKU.Comentario);
                            }
                            else
                            {

                                //REGISTRO 1
                                //PDF
                                html.AppendLine("<tr " + columnStyleHtml + ">");

                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>" + reporteSKU.Descripcion + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //ACTUAL
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.PrecioConc.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.FactorUtilidadBruto.ToString("#,##0.00") + "</td>");

                                //PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.PrecioConce.ToString("#,##0.00") + "</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>" + reporteSKU.FactorUtilidadCC.ToString("#,##0.00") + "</td>");

                                //VENTA REAL AÑO ANTERIOR
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PRESUPUESTO SIN PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PRESUPUESTO CON PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //COMENTARIO
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //FIN CONTENIDO CELDAS
                                html.AppendLine("</tr>");

                                //REGISTRO 2
                                //PDF
                                html.AppendLine("<tr " + columnStyleHtml + ">");

                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //ACTUAL
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //VENTA REAL AÑO ANTERIOR
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PRESUPUESTO SIN PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //PRESUPUESTO CON PROMOCION
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //COMENTARIO
                                html.AppendLine("<td style='text-align: right; border: 0.5px solid black;'>&nbsp;</td>");

                                //FIN CONTENIDO CELDAS
                                html.AppendLine("</tr>");


                                contRegPDF = contRegPDF + 1;


                                //REGISTRO 1
                                //EXCEL
                                dtReporteSKU.Rows.Add(
                                                //1      
                                                string.Empty,
                                                string.Empty,
                                                string.IsNullOrEmpty(reporteSKU.Descripcion) ? string.Empty : reporteSKU.Descripcion,
                                                string.IsNullOrEmpty(reporteSKU.Capacidad) ? string.Empty : reporteSKU.Capacidad,
                                                
                                                reporteSKU.CostoMPAnterior.ToString("#,##0.00"),
                                                reporteSKU.GastoPlantaAnterior.ToString("#,##0.00"),
                                                reporteSKU.GastoKromaAñoAnterior.ToString("#,##0.00"),
                                                reporteSKU.CostoMPEnvConvAñoActual.ToString("#,##0.00"),
                                                reporteSKU.CostoPlantaAñoActual.ToString("#,##0.00"),
                                                reporteSKU.GastoKromaActual.ToString("#,##0.00"),
                                                reporteSKU.PrecioConcAnterior.ToString("#,##0.00"),
                                                //11
                                                reporteSKU.PrecioAntConPromocion.ToString("#,##0.00"),
                                                reporteSKU.InversionTintaAnt.ToString("#,##0.00"),
                                                reporteSKU.PrecioPublicoAntCP.ToString("#,##0.00"),
                                                reporteSKU.PrecioConc.ToString("#,##0.00"),
                                                reporteSKU.FactorUtilidadBruto.ToString("#,##0.00"),
                                                reporteSKU.PrecioPublicoSC.ToString("#,##0.00"),
                                                reporteSKU.InversionTintaActSC.ToString("#,##0.00"),
                                                reporteSKU.PorcMargenConcSC.ToString("#,##0.00"),
                                                reporteSKU.PrecioConce.ToString("#,##0.00"),
                                                reporteSKU.FactorUtilidadCC.ToString("#,##0.00"),

                                                //21
                                                reporteSKU.PrecioPublicoCC.ToString("#,##0.00"),
                                                reporteSKU.InversionTintaActCC.ToString("#,##0.00"),
                                                reporteSKU.PorcMargenConcCC.ToString("#,##0.00"),
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //31
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //41
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //NUEVOS
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty);

                                //REGISTRO 2
                                //EXCEL
                                dtReporteSKU.Rows.Add(
                                                //1   
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //11
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //21
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //31
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //41
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                //NUEVOS
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty,

                                                string.Empty,
                                                string.Empty,
                                                string.Empty,
                                                string.Empty);

                            }

                            if (primerPaginaSKU && contRegSKU > RegPrimerPaginaSKU)
                            {
                                html.AppendLine(finDatosHtml.ToString());
                                html.AppendLine(finHtml.ToString());

                                ListHtmlSKU.Add(html.ToString());

                                html = new StringBuilder();

                                html.AppendLine(inicioHtml.ToString());
                                html.AppendLine(inicioDatosHtml.ToString());

                                primerPaginaSKU = false;

                                contRegSKU = 0;
                            }
                            else if (contRegSKU > RegSegundoPaginaSKU)
                            {
                                html.AppendLine(finDatosHtml.ToString());
                                html.AppendLine(finHtml.ToString());

                                ListHtmlSKU.Add(html.ToString());

                                html = new StringBuilder();

                                html.AppendLine(inicioHtml.ToString());
                                html.AppendLine(inicioDatosHtml.ToString());

                                primerPaginaSKU = false;

                                contRegSKU = 0;
                            }


                            contRegSKU++;

                            reporteSKUAnterior = reporteSKU;

                        }

                        html.AppendLine(finDatosHtml.ToString());
                        html.AppendLine(finHtml.ToString());

                        //html.AppendLine(ImgHtmlFooter);
                        //html.AppendLine("</div>");

                        ListHtmlSKU.Add(html.ToString());


                        //CREAR ARCHIVO PDF
                        //nombreArchivo = Guid.NewGuid().ToString() + ".pdf";

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                               ConfigurationManager.AppSettings["DirectorioReporte"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            pathArchivo = parametro.Valor;
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

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["WidthPDFHorizontal"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            widthPDF = Convert.ToDouble(parametro.Valor);
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["HeigthPDFHorizontal"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            heigthPDF = Convert.ToDouble(parametro.Valor);
                        }
                        

                        //var Renderer = new IronPdf.HtmlToPdf();

                        //HtmlHeaderFooter htmlFooter = new HtmlHeaderFooter();
                        //htmlFooter.HtmlFragment = ImgHtmlFooter;
                        //Renderer.PrintOptions.Footer = htmlFooter;

                        //HtmlHeaderFooter htmlHeader = new HtmlHeaderFooter();
                        //htmlHeader.HtmlFragment = ImgHtmlHeader;
                        //Renderer.PrintOptions.Header = htmlHeader;

                        //Renderer.PrintOptions.SetCustomPaperSizeinMilimeters(widthPDF, heigthPDF);
                        //Renderer.PrintOptions.Header.DrawDividerLine = false;
                        //var PDF = Renderer.RenderHtmlAsPdf(html.ToString());
                        //PDF.SaveAs(pathArchivoCompleto);

                        //ANTERIOR HTML
                        //htmlReporteSKU = html.ToString();


                        ////GENERAR PDF CON DLL SPIRE
                        //layoutSKU = new PdfHtmlLayoutFormat();
                        //layoutSKU.IsWaiting = true;

                        //pageSKU = new PdfPageSettings();
                        //pageSKU.Size = new SizeF(1200, 2500);
                        ////pageSKU.Rotate = PdfPageRotateAngle.RotateAngle180;
                        //pageSKU.Orientation = PdfPageOrientation.Landscape;

                        //int left = 15;
                        //int rigth = 15;
                        //int top = 35;
                        //int bottom = 35;

                        //pageSKU.Margins = new Spire.Pdf.Graphics.PdfMargins(left, top, rigth, bottom);

                        //docSKU = new PdfDocument();

                        //Thread thread = new Thread(() =>
                        //{
                        //    docSKU.LoadFromHTML(htmlReporteSKU, autoDetectPage, pageSKU, layoutSKU);
                        //});

                        //thread.SetApartmentState(ApartmentState.STA);

                        //thread.Start();

                        //thread.Join();

                        //docSKU.SaveToFile(pathArchivoCompleto);
                        //docSKU.Clone();


                        ///////////////////
                        ////AGREGAR HEADER Y FOOTER
                        //docSKUImg = new PdfDocument();

                        //Thread thread2 = new Thread(() =>
                        //{
                        //    docSKUImg.LoadFromFile(pathArchivoCompleto);
                        //});

                        //headerImage = PdfImage.FromFile(pathImgHeader);
                        //footerImage = PdfImage.FromFile(pathImgFooter);

                        //thread2.SetApartmentState(ApartmentState.STA);
                        //thread2.Start();
                        //thread2.Join();

                        //int xHeader = 5;
                        //int yHeader = 5;
                        //int xFooter = 1783;
                        //int yFooter = 1175;

                        //foreach (PdfPageBase pdfPage in docSKUImg.Pages)
                        //{
                        //    pdfPage.Canvas.DrawImage(headerImage, xHeader, yHeader, 712, 20);
                        //    pdfPage.Canvas.DrawImage(footerImage, xFooter, yFooter, 712, 20);
                        //}

                        //docSKUImg.SaveToFile(pathArchivoCompleto);
                        //docSKUImg.Close();


                        //REPORTE SKU
                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                                ConfigurationManager.AppSettings["UrlReporteSKU"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            url = HttpUtility.HtmlEncode(parametro.Valor);
                        }

                        parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                               ConfigurationManager.AppSettings["NombreReporteSKU"].ToString().ToUpper()).FirstOrDefault();
                        if (parametro != null)
                        {
                            nombreReporte = parametro.Valor + rentabilidad.ClaveCampania;
                        }


                        //NOMBRE REPORTE CARPETA TEMPORAL
                        pathArchivoCompletoSKU = Path.Combine(pathArchivo, nombreReporte + ".pdf");

                        //NOMBRE REPORTE 
                        urlCompletoSKU = Path.Combine(url, nombreReporte + ".pdf");
                        
                        rentabilidadENTRes.UrlReporteSKUPDF = urlCompletoSKU;


                        var t = System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                ////////////////////
                                //GENERAR PDF SKU
                                //PDFTool.PDF.GenerarPDF(htmlReporteSKU, pathImgHeader, pathImgFooter, pathArchivoCompletoSKU, 1200, 2650);

                                PDFTool.PDF.GenerarPDFSKU(ListHtmlSKU, pathImgHeader, pathImgFooter, pathArchivoCompletoSKU, 1200, 2650);

                                return "OK";
                            }
                            catch (Exception ex)
                            {
                                ArchivoLog.EscribirLog(null, "ERROR: Service - PDF SKU, Source:" + ex.Source + ", Message:" + ex.Message);

                                return "ERROR: Service - PDF SKU, Source:" + ex.Source + ", Message:" + ex.Message;
                            }
                        });

                        var c = t.ContinueWith((antecedent) =>
                        {
                            if (antecedent.Result == "OK")
                            {
                                using (WebClient client = new WebClient())
                                {
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                                    client.UploadFile(urlCompletoSKU, "PUT", pathArchivoCompletoSKU);

                                    
                                }
                            }

                        });



                        //CREAR ARCHIVO EXCEL
                        //dsReporteSKU.Tables.Add(dtReporteSKU);

                        //nombreArchivo = Guid.NewGuid().ToString() + ".xlsx";

                        //NOMBRE REPORTE CARPETA TEMPORAL
                        pathArchivoCompletoSKUEx = Path.Combine(pathArchivo, nombreReporte + ".xlsx");

                        //ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoCompleto, dsReporteSKU);


                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            excel.Workbook.Worksheets.Add("ReporteSKU");

                            var worksheet = excel.Workbook.Worksheets["ReporteSKU"];

                            worksheet.SetValue("A1", "Reporte SKU");
                            worksheet.SetValue("A2", "Campaña:");
                            worksheet.SetValue("B2", nombreCampana);
                            worksheet.SetValue("A3", "Clave Campaña:");
                            worksheet.SetValue("B3", claveCampana);
                            worksheet.SetValue("A4", "Fecha Inicio SubCanal:");
                            if (!string.IsNullOrEmpty(fechaInicioSubCanal) && fechaInicioSubCanal != "01/01/0001")
                            {
                                worksheet.SetValue("B4", fechaInicioSubCanal);
                            }
                            worksheet.SetValue("A5", "Fecha Fin SubCanal:");
                            if (!string.IsNullOrEmpty(fechaFinSubCanal) && fechaFinSubCanal != "01/01/0001")
                            {
                                worksheet.SetValue("B5", fechaFinSubCanal);
                            }
                            worksheet.SetValue("A6", "Fecha Inicio Publico");
                            if (!string.IsNullOrEmpty(fechaInicioPublico) && fechaInicioPublico != "01/01/0001")
                            {
                                worksheet.SetValue("B6", fechaInicioPublico);
                            }
                            worksheet.SetValue("A7", "Fecha Fin Publico");
                            if (!string.IsNullOrEmpty(fechaFinPublico) && fechaFinPublico != "01/01/0001")
                            {
                                worksheet.SetValue("B7", fechaFinPublico);
                            }

                            worksheet.Cells["A1:B7"].Style.Font.Color.SetColor(1, 38, 130, 221);
                            worksheet.Cells["A1:B7"].Style.Font.Bold = true;
                            worksheet.Cells["A1:B7"].Style.Font.Size = 14;

                            worksheet.Cells["A9:BY9"].LoadFromDataTable(dtReporteSKU, true);//, OfficeOpenXml.Table.TableStyles.Light2);

                            worksheet.Cells["A9:BY9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A9:BY9"].Style.Fill.BackgroundColor.SetColor(1, 59, 137, 216);
                            worksheet.Cells["A9:BY9"].Style.Font.Color.SetColor(Color.White);

                            //worksheet.Cells["A9:BE" + (dtReporteSKU.Rows.Count + 9)].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            worksheet.Cells["A9:BY" + (dtReporteSKU.Rows.Count + 9)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A9:BY" + (dtReporteSKU.Rows.Count + 9)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A9:BY" + (dtReporteSKU.Rows.Count + 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A9:BY" + (dtReporteSKU.Rows.Count + 9)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                            if (dtReporteSKU.Rows.Count > 0)
                            {
                                worksheet.Cells["BY10:BY" + (dtReporteSKU.Rows.Count + 9)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            }

                            //COMENTARIOS
                            for (int m = 0; m < dtReporteSKU.Rows.Count; m++)
                            {
                                DataRow dr = dtReporteSKU.Rows[m];

                                if(string.IsNullOrEmpty(dr["COMENTARIO"].ToString()))
                                {
                                    worksheet.Cells["BY" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 255, 255, 255);
                                    worksheet.Cells["BY" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else if (dr["COMENTARIO"].ToString().ToUpper() == etiquetaRentable)
                                {
                                    worksheet.Cells["BY" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 171, 235, 198);
                                    worksheet.Cells["BY" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else if (dr["COMENTARIO"].ToString().ToUpper() == etiqueteNoRentable)
                                {
                                    worksheet.Cells["BY" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 245, 183, 177);
                                    worksheet.Cells["BY" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else if (dr["COMENTARIO"].ToString().ToUpper() == etiquetaRevisar)
                                {
                                    worksheet.Cells["BY" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 249, 231, 159);
                                    worksheet.Cells["BY" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                                else
                                {
                                    worksheet.Cells["BY" + (m + 10)].Style.Fill.BackgroundColor.SetColor(1, 255, 255, 255);
                                    worksheet.Cells["BY" + (m + 10)].Style.Font.Color.SetColor(Color.Black);
                                }
                            }

                            //worksheet.Cells["A8:BE8"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            //ENCABEZADO VERDE
                            worksheet.Cells["E9:G9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["K9:N9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["Y9:AJ9"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            worksheet.Cells["E9:G9"].Style.Fill.BackgroundColor.SetColor(1, 76, 127, 15);
                            worksheet.Cells["K9:N9"].Style.Fill.BackgroundColor.SetColor(1, 76, 127, 15);
                            worksheet.Cells["Y9:AJ9"].Style.Fill.BackgroundColor.SetColor(1, 76, 127, 15);

                            worksheet.Cells["E9:G9"].Style.Font.Color.SetColor(Color.White);
                            worksheet.Cells["K9:N9"].Style.Font.Color.SetColor(Color.White);
                            worksheet.Cells["Y9:AJ9"].Style.Font.Color.SetColor(Color.White);

                            //ENCABEZADO AGRUPADO AZUL
                            worksheet.Cells["O8:S8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["T8:X8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["AK8:AV8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["AW8:BH8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["BI8:BT8"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            worksheet.Cells["O8:S8"].Style.Font.Color.SetColor(Color.White);
                            worksheet.Cells["T8:X8"].Style.Font.Color.SetColor(Color.White);
                            worksheet.Cells["AK8:AV8"].Style.Font.Color.SetColor(Color.White);
                            worksheet.Cells["AW8:BH8"].Style.Font.Color.SetColor(Color.White);
                            worksheet.Cells["BI8:BT8"].Style.Font.Color.SetColor(Color.White);

                            worksheet.Cells["O8:S8"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                            worksheet.Cells["T8:X8"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                            worksheet.Cells["AK8:AV8"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                            worksheet.Cells["AW8:BH8"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);
                            worksheet.Cells["BI8:BT8"].Style.Fill.BackgroundColor.SetColor(1, 38, 130, 221);

                            worksheet.Cells["O8:S8"].Merge = true;
                            worksheet.Cells["T8:X8"].Merge = true;
                            worksheet.Cells["AK8:AV8"].Merge = true;
                            worksheet.Cells["AW8:BH8"].Merge = true;
                            worksheet.Cells["BI8:BT8"].Merge = true;

                            worksheet.Cells["O8:S8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells["T8:X8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells["AK8:AV8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells["AW8:BH8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells["BI8:BT8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            worksheet.Cells["O8:S8"].Value = "ACTUAL";
                            worksheet.Cells["T8:X8"].Value = "PROMOCIÓN";
                            worksheet.Cells["AK8:AV8"].Value = "PRESUPUESTO SIN PROMOCIÓN";
                            worksheet.Cells["AW8:BH8"].Value = "PRESUPUESTO CON PROMOCIÓN";
                            worksheet.Cells["BI8:BT8"].Value = "PRESUPUESTO CON PROMOCIÓN (NECESARIO)";

                            //ENCABEZADO AGRUPADO VERDE
                            worksheet.Cells["Y8:AJ8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["Y8:AJ8"].Style.Fill.BackgroundColor.SetColor(1, 76, 127, 15);
                            worksheet.Cells["Y8:AJ8"].Style.Font.Color.SetColor(Color.White);
                            worksheet.Cells["Y8:AJ8"].Merge = true;
                            worksheet.Cells["Y8:AJ8"].Value = "VENTA REAL AÑO ANTERIOR";
                            worksheet.Cells["Y8:AJ8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            //AUTO SIZE COLUMNAS
                            worksheet.Cells["A1:BY" + (dtReporteSKU.Rows.Count + 9)].AutoFitColumns();


                            FileInfo excelFile = new FileInfo(pathArchivoCompletoSKUEx);
                            excel.SaveAs(excelFile);
                        }

                        //NOMBRE REPORTE SHARE POINT
                        urlCompletoSKUEx = Path.Combine(url, nombreReporte + ".xlsx");

                        using (WebClient client = new WebClient())
                        {
                            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                            client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                            client.UploadFile(urlCompletoSKUEx, "PUT", pathArchivoCompletoSKUEx);

                            rentabilidadENTRes.UrlReporteSKUExcel = urlCompletoSKUEx;
                        }

                    }

                    #endregion


                    rentabilidadENTRes.Mensaje = "OK";
                }
                else
                {
                    rentabilidadENTRes.Mensaje = rentabilidad.Mensaje;

                    ArchivoLog.EscribirLog(null, "ERROR: Service - MostrarRentabilidad, Ocurrio un problema en el SP de Calculo de Rentabilidad. ");
                }

            }
            catch (Exception ex)
            {
                rentabilidadENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se pudo calcular correctamente la Rentabilidad de Campaña, intentelo de nuevo o consulte al administrador.";

                ArchivoLog.EscribirLog(null, "ERROR: Service - MostrarRentabilidad, Source:" + ex.Source + ", Message:" + ex.Message);
            }

            return rentabilidadENTRes;
        }
        public ReporteENT GenerarReporteDirectivo(ReporteENT reporteENTReq)
        {
            ReporteENT reporteENTres = new ReporteENT();
            CampanaDAT campanaDAT = new CampanaDAT();
            List<ReporteCEO> ListReporteCEO = new List<ReporteCEO>();

            string nombreArchivo = string.Empty;
            string pathArchivo = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;
            string pathArchivoCompleto = string.Empty;
            string url = string.Empty;
            string urlCompleto = string.Empty;
            string nombreReporte = string.Empty;

            StringBuilder html = new StringBuilder();
            DataSet dsReporte = new DataSet();
            DataTable dtReporteCEO = new DataTable();

            DataTable dtParametro = new DataTable();
            ParametroDAT parametroDAT = new ParametroDAT();
            List<Parametro> ListParametro = new List<Parametro>();
            Parametro parametro = new Parametro();

            try
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

                //OBTENER DATOS DEL REPORTE
                ListReporteCEO = campanaDAT.MostrarReporteCEO(reporteENTReq.ClaveCampana);

                if (ListReporteCEO.Count > 0)
                {
                    html.AppendLine("<h1>Reporte CEO</h1><br />");
                    html.AppendLine("<table>");
                    html.AppendLine("<thead>");
                    html.AppendLine("<tr>");
                    //AGREGAR COLUMNAS
                    html.AppendLine("<th>NOMBRE CAMPANIA</th>");
                    html.AppendLine("<th>PERIODO</th>");
                    html.AppendLine("<th>SUBTOTAL LITROS</th>");
                    html.AppendLine("<th>SUBTOTAL PIEZAS</th>");
                    html.AppendLine("<th>TOTAL LITROS / PIEZAS</th>");
                    html.AppendLine("<th>IMPORTE</th>");
                    html.AppendLine("<th>IMPORTE LITROS</th>");
                    html.AppendLine("<th>IMPORTE PZAS</th>");
                    html.AppendLine("<th>PRECIO PROMEDIO X LT</th>");
                    html.AppendLine("<th>COSTO MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>FACTOR UTILIDAD</th>");
                    //html.AppendLine("<th>MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACIÓN</th>");
                    html.AppendLine("<th>INVERSION PUBLICITARIA</th>");
                    html.AppendLine("<th>OTROS GASTOS DE OPERACIÓN PLANTA</th>");
                    html.AppendLine("<th>GASTOS DE OPERACIÓN GENERALES KROMA</th>");
                    html.AppendLine("<th>NOTAS DE CRÉDITO</th>");
                    html.AppendLine("<th>UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION GASTOS DE OPERACIÓN GENERALES KROMA</th>");
                    html.AppendLine("<th>UTILIDAD POR LT/PZ</th>");
                    html.AppendLine("<th>FACTOR DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS</th>");
                    html.AppendLine("<th>% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR</th>");
                    html.AppendLine("<th>% INCREMENTO EN LITROS Vs PRESUPUESTO</th>");
                    html.AppendLine("<th>ROI</th>");

                    html.AppendLine("</tr>");
                    html.AppendLine("</thead>");
                    html.AppendLine("<tbody>");


                    foreach (ReporteCEO reporteCEO in ListReporteCEO)
                    {
                        html.AppendLine("<tr>");
                        //AGREGAR CONTENIDO DE LAS CELDAS
                        html.AppendLine("<td>" + reporteCEO.ClaveCampania + "</td>");
                        html.AppendLine("<td>" + reporteCEO.Periodo + "</td>");
                        html.AppendLine("<td>" + reporteCEO.SubtotalLitros + "</td>");
                        html.AppendLine("<td>" + reporteCEO.SubtotalPiezas + "</td>");
                        html.AppendLine("<td>" + reporteCEO.TotalLitrosPiezas + "</td>");
                        html.AppendLine("<td>" + reporteCEO.Importe + "</td>");
                        html.AppendLine("<td>" + reporteCEO.ImporteLitros + "</td>");
                        html.AppendLine("<td>" + reporteCEO.ImportePiezas + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PrecioPromedioLitro + "</td>");
                        html.AppendLine("<td>" + reporteCEO.CostoMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.UtilidadMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.FactorUtilidad + "</td>");
                        //html.AppendLine("<td>" + reporteCEO.MP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.MargenUtilidad + "</td>");
                        html.AppendLine("<td>" + reporteCEO.InversionPublicidad + "</td>");
                        html.AppendLine("<td>" + reporteCEO.OtrosGastos + "</td>");
                        html.AppendLine("<td>" + reporteCEO.GastosOperacion + "</td>");
                        html.AppendLine("<td>" + reporteCEO.NotasCredito + "</td>");
                        html.AppendLine("<td>" + reporteCEO.UtilidadConsideraMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.UtilidadLitroPieza + "</td>");
                        html.AppendLine("<td>" + reporteCEO.FactorUtilidadMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PorcenUtilidadConsideraMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PorcenIncrementoLitros + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PorcenIncrementoLitrosPresu + "</td>");
                        html.AppendLine("<td>" + reporteCEO.Roi + "</td>");

                        html.AppendLine("</tr>");
                    }

                    html.AppendLine("</tbody>");
                    html.AppendLine("</table>");

                }

                //CREAR ARCHIVO PDF
                nombreArchivo = Guid.NewGuid().ToString() + ".pdf";

                parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                       ConfigurationManager.AppSettings["DirectorioReporte"].ToString().ToUpper()).FirstOrDefault();
                if (parametro != null)
                {
                    pathArchivo = parametro.Valor;
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

                parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                    ConfigurationManager.AppSettings["UrlReporteCEO"].ToString().ToUpper()).FirstOrDefault();
                if (parametro != null)
                {
                    url = parametro.Valor;
                }

                parametro = ListParametro.Where(n => n.Nombre.ToUpper() ==
                    ConfigurationManager.AppSettings["NombreReporteCEO"].ToString().ToUpper()).FirstOrDefault();
                if (parametro != null)
                {
                    nombreReporte = parametro.Valor + reporteENTReq.ClaveCampana + ".pdf"; ;
                }

                pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);

                var Renderer = new IronPdf.HtmlToPdf();
                var PDF = Renderer.RenderHtmlAsPdf(html.ToString());
                PDF.SaveAs(pathArchivoCompleto);

                urlCompleto = Path.Combine(url, nombreReporte);

                using (WebClient client = new WebClient())
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                    client.UploadFile(urlCompleto, "PUT", pathArchivoCompleto);
                }

                dtReporteCEO.Columns.Add("NOMBRE CAMPANIA");
                dtReporteCEO.Columns.Add("PERIODO");
                dtReporteCEO.Columns.Add("SUBTOTAL LITROS");
                dtReporteCEO.Columns.Add("SUBTOTAL PIEZAS");
                dtReporteCEO.Columns.Add("TOTAL LITROS / PIEZAS");
                dtReporteCEO.Columns.Add("IMPORTE");
                dtReporteCEO.Columns.Add("IMPORTE LITROS");
                dtReporteCEO.Columns.Add("IMPORTE PZAS");
                dtReporteCEO.Columns.Add("PRECIO PROMEDIO X LT");
                dtReporteCEO.Columns.Add("COSTO MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("FACTOR UTILIDAD");
                //dtReporteCEO.Columns.Add("MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACIÓN");
                dtReporteCEO.Columns.Add("INVERSION PUBLICITARIA");
                dtReporteCEO.Columns.Add("OTROS GASTOS DE OPERACIÓN PLANTA");
                dtReporteCEO.Columns.Add("GASTOS DE OPERACIÓN GENERALES KROMA");
                dtReporteCEO.Columns.Add("NOTAS DE CRÉDITO");
                dtReporteCEO.Columns.Add("UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION GASTOS DE OPERACIÓN GENERALES KROMA");
                dtReporteCEO.Columns.Add("UTILIDAD POR LT/PZ");
                dtReporteCEO.Columns.Add("FACTOR DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS");
                dtReporteCEO.Columns.Add("% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR");
                dtReporteCEO.Columns.Add("% INCREMENTO EN LITROS Vs PRESUPUESTO");
                dtReporteCEO.Columns.Add("ROI");

                //CREAR ARCHIVO EXCEL
                ListReporteCEO.ForEach(n =>
                {
                    dtReporteCEO.Rows.Add(n.ClaveCampania,
                                            n.Periodo,
                                            n.SubtotalLitros,
                                            n.SubtotalPiezas,
                                            n.TotalLitrosPiezas,
                                            n.Importe,
                                            n.ImporteLitros,
                                            n.ImportePiezas,
                                            n.PrecioPromedioLitro,
                                            n.CostoMP,
                                            n.UtilidadMP,
                                            n.FactorUtilidad,
                                            //.n.MP,
                                            n.MargenUtilidad,
                                            n.InversionPublicidad,
                                            n.OtrosGastos,
                                            n.GastosOperacion,
                                            n.NotasCredito,
                                            n.UtilidadConsideraMP,
                                            n.UtilidadLitroPieza,
                                            n.FactorUtilidadMP,
                                            n.MargenUtilidad,
                                            n.PorcenIncrementoLitros,
                                            n.PorcenIncrementoLitrosPresu,
                                            n.Roi);
                });

                //dtReporteCEO.Rows.Add()


                //object[] dato = { 1, 2 };

                //dtReporteCEO.LoadDataRow(dato, true);

                dsReporte.Tables.Add(dtReporteCEO);

                nombreArchivo = Guid.NewGuid().ToString() + ".xls";
                pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);

                ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoCompleto, dsReporte);

                nombreReporte = nombreReporte + reporteENTReq.ClaveCampana + ".xls";

                urlCompleto = Path.Combine(url, nombreReporte);

                using (WebClient client = new WebClient())
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                    client.UploadFile(urlCompleto, "PUT", pathArchivoCompleto);
                }

                reporteENTres.UrlArchivo = urlCompleto;
                reporteENTres.Mensaje = "OK";

            }
            catch (Exception ex)
            {
                reporteENTres.Mensaje = "ERROR: Ocurrio un error inesperado, no se pudo generar el reporte de Directivo Campaña, intente nuevamente o consulte al adminstrador de sistemas.";

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - GenerarReporteDirectivo, Source: " + ex.Source + " Message: " + ex.Message);
            }

            return reporteENTres;
        }
        public ReporteENT GenerarReporteDirectivo2(ReporteENT reporteENTReq)
        {
            ReporteENT reporteENTres = new ReporteENT();
            CampanaDAT campanaDAT = new CampanaDAT();
            List<ReporteCEO> ListReporteCEO = new List<ReporteCEO>();

            string nombreArchivo = string.Empty;
            string pathArchivo = string.Empty;
            string usuarioSharePoint = string.Empty;
            string passwordSharePoint = string.Empty;
            string pathArchivoCompleto = string.Empty;
            string url = string.Empty;
            string urlCompleto = string.Empty;
            string nombreReporte = string.Empty;

            StringBuilder html = new StringBuilder();
            DataSet dsReporte = new DataSet();
            DataTable dtReporteCEO = new DataTable();

            try
            {
                ListReporteCEO = campanaDAT.MostrarReporteCEO(reporteENTReq.ClaveCampana);

                if (ListReporteCEO.Count > 0)
                {
                    html.AppendLine("<h1>Reporte CEO</h1><br />");
                    html.AppendLine("<table>");
                    html.AppendLine("<thead>");
                    html.AppendLine("<tr>");
                    //AGREGAR COLUMNAS
                    html.AppendLine("<th>NOMBRE CAMPANIA</th>");
                    html.AppendLine("<th>PERIODO</th>");
                    html.AppendLine("<th>SUBTOTAL LITROS</th>");
                    html.AppendLine("<th>SUBTOTAL PIEZAS</th>");
                    html.AppendLine("<th>TOTAL LITROS / PIEZAS</th>");
                    html.AppendLine("<th>IMPORTE</th>");
                    html.AppendLine("<th>IMPORTE LITROS</th>");
                    html.AppendLine("<th>IMPORTE PZAS</th>");
                    html.AppendLine("<th>PRECIO PROMEDIO X LT</th>");
                    html.AppendLine("<th>COSTO MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>FACTOR UTILIDAD</th>");
                    html.AppendLine("<th>MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACIÓN</th>");
                    html.AppendLine("<th>INVERSION PUBLICITARIA</th>");
                    html.AppendLine("<th>OTROS GASTOS DE OPERACIÓN PLANTA</th>");
                    html.AppendLine("<th>GASTOS DE OPERACIÓN GENERALES KROMA</th>");
                    html.AppendLine("<th>NOTAS DE CRÉDITO</th>");
                    html.AppendLine("<th>UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION GASTOS DE OPERACIÓN GENERALES KROMA</th>");
                    html.AppendLine("<th>UTILIDAD POR LT/PZ</th>");
                    html.AppendLine("<th>FACTOR DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION</th>");
                    html.AppendLine("<th>% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS</th>");
                    html.AppendLine("<th>% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR</th>");
                    html.AppendLine("<th>% INCREMENTO EN LITROS Vs PRESUPUESTO</th>");
                    html.AppendLine("<th>ROI</th>");

                    html.AppendLine("</tr>");
                    html.AppendLine("</thead>");
                    html.AppendLine("<tbody>");


                    foreach (ReporteCEO reporteCEO in ListReporteCEO)
                    {
                        html.AppendLine("<tr>");
                        //AGREGAR CONTENIDO DE LAS CELDAS
                        html.AppendLine("<td>" + reporteCEO.ClaveCampania + "</td>");
                        html.AppendLine("<td>" + reporteCEO.Periodo + "</td>");
                        html.AppendLine("<td>" + reporteCEO.SubtotalLitros + "</td>");
                        html.AppendLine("<td>" + reporteCEO.SubtotalPiezas + "</td>");
                        html.AppendLine("<td>" + reporteCEO.TotalLitrosPiezas + "</td>");
                        html.AppendLine("<td>" + reporteCEO.Importe + "</td>");
                        html.AppendLine("<td>" + reporteCEO.ImporteLitros + "</td>");
                        html.AppendLine("<td>" + reporteCEO.ImportePiezas + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PrecioPromedioLitro + "</td>");
                        html.AppendLine("<td>" + reporteCEO.CostoMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.UtilidadMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.FactorUtilidad + "</td>");
                        html.AppendLine("<td>" + reporteCEO.MP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.MargenUtilidad + "</td>");
                        html.AppendLine("<td>" + reporteCEO.InversionPublicidad + "</td>");
                        html.AppendLine("<td>" + reporteCEO.OtrosGastos + "</td>");
                        html.AppendLine("<td>" + reporteCEO.GastosOperacion + "</td>");
                        html.AppendLine("<td>" + reporteCEO.NotasCredito + "</td>");
                        html.AppendLine("<td>" + reporteCEO.UtilidadConsideraMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.UtilidadLitroPieza + "</td>");
                        html.AppendLine("<td>" + reporteCEO.FactorUtilidadMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PorcenUtilidadConsideraMP + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PorcenIncrementoLitros + "</td>");
                        html.AppendLine("<td>" + reporteCEO.PorcenIncrementoLitrosPresu + "</td>");
                        html.AppendLine("<td>" + reporteCEO.Roi + "</td>");

                        html.AppendLine("</tr>");
                    }

                    html.AppendLine("</tbody>");
                    html.AppendLine("</table>");

                }

                //CREAR ARCHIVO PDF
                nombreArchivo = Guid.NewGuid().ToString() + ".pdf";
                pathArchivo = ConfigurationManager.AppSettings["directorioReporte"];
                usuarioSharePoint = ConfigurationManager.AppSettings["usuarioSharePoint"];
                passwordSharePoint = ConfigurationManager.AppSettings["passwordSharePoint"];
                pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);

                //var Renderer = new IronPdf.HtmlToPdf();
                //var PDF = Renderer.RenderHtmlAsPdf(html.ToString());
                //PDF.SaveAs(pathArchivoCompleto);

                url = HttpUtility.HtmlEncode(ConfigurationManager.AppSettings["UrlReporteSharePoint"]);
                nombreReporte = ConfigurationManager.AppSettings["NombreReporte"] + reporteENTReq.ClaveCampana + ".pdf";

                urlCompleto = Path.Combine(url, nombreReporte);

                using (WebClient client = new WebClient())
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                    client.UploadFile(urlCompleto, "PUT", pathArchivoCompleto);
                }

                dtReporteCEO.Columns.Add("NOMBRE CAMPANIA");
                dtReporteCEO.Columns.Add("PERIODO");
                dtReporteCEO.Columns.Add("SUBTOTAL LITROS");
                dtReporteCEO.Columns.Add("SUBTOTAL PIEZAS");
                dtReporteCEO.Columns.Add("TOTAL LITROS / PIEZAS");
                dtReporteCEO.Columns.Add("IMPORTE");
                dtReporteCEO.Columns.Add("IMPORTE LITROS");
                dtReporteCEO.Columns.Add("IMPORTE PZAS");
                dtReporteCEO.Columns.Add("PRECIO PROMEDIO X LT");
                dtReporteCEO.Columns.Add("COSTO MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("UTILIDAD BRUTA MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("FACTOR UTILIDAD");
                dtReporteCEO.Columns.Add("MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACIÓN");
                dtReporteCEO.Columns.Add("INVERSION PUBLICITARIA");
                dtReporteCEO.Columns.Add("OTROS GASTOS DE OPERACIÓN PLANTA");
                dtReporteCEO.Columns.Add("GASTOS DE OPERACIÓN GENERALES KROMA");
                dtReporteCEO.Columns.Add("NOTAS DE CRÉDITO");
                dtReporteCEO.Columns.Add("UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION GASTOS DE OPERACIÓN GENERALES KROMA");
                dtReporteCEO.Columns.Add("UTILIDAD POR LT/PZ");
                dtReporteCEO.Columns.Add("FACTOR DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION");
                dtReporteCEO.Columns.Add("% MARGEN DE UTILIDAD CONSIDERA MATERIA PRIMA + ENVASE + FABRICACION + GASTOS");
                dtReporteCEO.Columns.Add("% INCREMENTO EN LITROS Vs VENTA REAL AÑO ANTERIOR");
                dtReporteCEO.Columns.Add("% INCREMENTO EN LITROS Vs PRESUPUESTO");
                dtReporteCEO.Columns.Add("ROI");

                //CREAR ARCHIVO EXCEL
                ListReporteCEO.ForEach(n =>
                {
                    dtReporteCEO.Rows.Add(n.ClaveCampania,
                                            n.Periodo,
                                            n.SubtotalLitros,
                                            n.SubtotalPiezas,
                                            n.TotalLitrosPiezas,
                                            n.Importe,
                                            n.ImporteLitros,
                                            n.ImportePiezas,
                                            n.PrecioPromedioLitro,
                                            n.CostoMP,
                                            n.UtilidadMP,
                                            n.FactorUtilidad,
                                            n.MP,
                                            n.MargenUtilidad,
                                            n.InversionPublicidad,
                                            n.OtrosGastos,
                                            n.GastosOperacion,
                                            n.NotasCredito,
                                            n.UtilidadConsideraMP,
                                            n.UtilidadLitroPieza,
                                            n.FactorUtilidadMP,
                                            n.MargenUtilidad,
                                            n.PorcenIncrementoLitros,
                                            n.PorcenIncrementoLitrosPresu,
                                            n.Roi);
                });

                //dtReporteCEO.Rows.Add()


                //object[] dato = { 1, 2 };

                //dtReporteCEO.LoadDataRow(dato, true);

                dsReporte.Tables.Add(dtReporteCEO);

                nombreArchivo = Guid.NewGuid().ToString() + ".xls";
                pathArchivoCompleto = Path.Combine(pathArchivo, nombreArchivo);

                ExcelLibrary.DataSetHelper.CreateWorkbook(pathArchivoCompleto, dsReporte);

                nombreReporte = ConfigurationManager.AppSettings["NombreReporte"] + reporteENTReq.ClaveCampana + ".xls";

                urlCompleto = Path.Combine(url, nombreReporte);

                using (WebClient client = new WebClient())
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    client.Credentials = new NetworkCredential(usuarioSharePoint, passwordSharePoint);
                    client.UploadFile(urlCompleto, "PUT", pathArchivoCompleto);
                }

                reporteENTres.UrlArchivo = urlCompleto;
                reporteENTres.Mensaje = "OK";

            }
            catch (Exception ex)
            {
                reporteENTres.Mensaje = "ERROR: Ocurrio un error inesperado, no se pudo generar el reporte de Directivo Campaña, intente nuevamente o consulte al adminstrador de sistemas.";

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - GenerarReporteDirectivo2, Source: " + ex.Source + " Message: " + ex.Message);

            }

            return reporteENTres;
        }
        public FlujoActividadENT MostrarAprobadorPredecesor(FlujoActividadENT flujoActividadENTReq)
        {
            FlujoActividadENT flujoActividadENTRes = new FlujoActividadENT();
            FlujoActividad flujoActividad = new FlujoActividad();

            List<FlujoActividad> ListFlujoActividad = new List<FlujoActividad>();
            CampanaDAT campanaDAT = new CampanaDAT();

            try
            {
                flujoActividad = flujoActividadENTReq.ListFlujoActividad.FirstOrDefault();

                if (flujoActividad != null)
                {
                    ListFlujoActividad = campanaDAT.MostrarAprobadorPredecesor(flujoActividad);

                    flujoActividadENTRes.OK = 1;
                    flujoActividadENTRes.Mensaje = "OK";
                    flujoActividadENTRes.ListFlujoActividad = ListFlujoActividad;
                }
                else
                {
                    flujoActividadENTRes.OK = 0;
                    flujoActividadENTRes.Mensaje = "ERROR: Se debe agregar informacion de la actividad para la busqueda del aprobador.";
                    flujoActividadENTRes.ListFlujoActividad = new List<FlujoActividad>();

                    ArchivoLog.EscribirLog(null, "ERROR:  Servicio - MostrarAprobadorPredecesor, Message: Ocurrio un error al ejecuatar el SP de Aprobador Campaña, revisar la informacion que se agrega.");
                }
            }
            catch (Exception ex)
            {
                flujoActividadENTRes.OK = 0;
                flujoActividadENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se puede mostrar la informacion del Aprobador de Campaña, intentalo de nuevo o consulta al administrador de sistemas.";
                flujoActividadENTRes.ListFlujoActividad = new List<FlujoActividad>();

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarAprobadorPredecesor, Source: " + ex.Source + " Message: " + ex.Message);
            }

            return flujoActividadENTRes;
        }
        public CampanaBitacoraENT MostrarCampanaBitacoraDetalle(CampanaBitacoraENT campanaBitacoraENTReq)
        {
            CampanaBitacoraENT campanaBitacoraENTRes = new CampanaBitacoraENT();

            campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

            CampanaDAT campanaDAT = new CampanaDAT();

            EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();

            try
            {
                campana = campanaBitacoraENTReq.ListCampana.FirstOrDefault();

                if (campana != null)
                {
                    campanaBitacoraENTRes.ListCampana = campanaDAT.MostrarCampanaBitacoraDetalle(campana);

                    campanaBitacoraENTRes.Mensaje = "OK";
                    campanaBitacoraENTRes.OK = 1;                
                }
                else
                {
                    campanaBitacoraENTRes.Mensaje = "ERROR: No se agrego la informacion necesaria para la busqueda del detalle y bitacora de la Campaña.";
                    campanaBitacoraENTRes.OK = 0;
                    campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampanaBitacoraDetalle, Message: Se debe agregar la informacion necesaria para la busqueda del detalle y bitacora de la campaña.");
                }
            }
            catch(Exception ex)
            {
                campanaBitacoraENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se puede mostrar la informacion del detalle y bitacora de la Campaña, intente de nuevo o consulte a su administrador.";
                campanaBitacoraENTRes.OK = 0;
                campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampanaBitacoraDetalle, Source: " + ex.Source + " Message: " + ex.Message);
            }

            return campanaBitacoraENTRes;
        }
        public CampanaBitacoraENT MostrarCampana(CampanaBitacoraENT campanaBitacoraENTreq)
        {
            CampanaBitacoraENT campanaBitacoraENTRes = new CampanaBitacoraENT();

            campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

            CampanaDAT campanaDAT = new CampanaDAT();

            EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();

            try
            {
                campana = campanaBitacoraENTreq.ListCampana.FirstOrDefault();

                if (campana != null)
                {
                    campanaBitacoraENTRes.ListCampana = campanaDAT.MostrarCampana(campana);

                    campanaBitacoraENTRes.Mensaje = "OK";
                    campanaBitacoraENTRes.OK = 1;
                }
                else
                {
                    campanaBitacoraENTRes.Mensaje = "ERROR: No se agrego la informacion necesaria para la busqueda del detalle y bitacora de la Campaña.";
                    campanaBitacoraENTRes.OK = 0;
                    campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampana, Message: Se debe agregar la informacion necesaria para la busqueda de la campaña.");
                }
            }
            catch(Exception ex)
            {
                campanaBitacoraENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se puede mostrar la informacion de la Campaña, intente de nuevo o consulte a su administrador.";
                campanaBitacoraENTRes.OK = 0;
                campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampana, Source: " + ex.Source + " Message: " + ex.Message);

            }

            return campanaBitacoraENTRes;
        }
        public CampanaBitacoraENT MostrarCampanaDetAnioAnterior(CampanaBitacoraENT campanaBitacoraENTreq)
        {
            CampanaBitacoraENT campanaBitacoraENTRes = new CampanaBitacoraENT();

            campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

            CampanaDAT campanaDAT = new CampanaDAT();

            EntidadesCampanasPPG.Modelo.Campana campana = new EntidadesCampanasPPG.Modelo.Campana();

            try
            {
                campana = campanaBitacoraENTreq.ListCampana.FirstOrDefault();

                if (campana != null)
                {
                    campanaBitacoraENTRes.ListCampana = campanaDAT.MostrarCampanaDetAnioAnterior(campana);

                    campanaBitacoraENTRes.Mensaje = "OK";
                    campanaBitacoraENTRes.OK = 1;
                }
                else
                {
                    campanaBitacoraENTRes.Mensaje = "ERROR: No se agrego la informacion necesaria para la busqueda de Campañas del año anterior.";
                    campanaBitacoraENTRes.OK = 0;
                    campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

                    ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampanaDetAnioAnterior, Message: Se debe agregar la informacion necesaria para la busqueda de la Campaña del año anterior.");
                }
            }
            catch (Exception ex)
            {
                campanaBitacoraENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se puede mostrar la informacion de la Campaña, intente de nuevo o consulte a su administrador.";
                campanaBitacoraENTRes.OK = 0;
                campanaBitacoraENTRes.ListCampana = new List<EntidadesCampanasPPG.Modelo.Campana>();

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarCampanaDetAnioAnterior, Source: " + ex.Source + " Message: " + ex.Message);

            }

            return campanaBitacoraENTRes;
        }
        public GastoPlantaENT MostrarGastoPlantaHistorico(GastoPlantaENT gastoPlantaENTReq)
        {
            GastoPlantaENT gastoPlantaENTRes = new GastoPlantaENT();
            GastoPlantaHistorico gastoPlantaHistorico = new GastoPlantaHistorico();

            CampanaDAT campanaDAT = new CampanaDAT();

            try
            {
                gastoPlantaENTRes.ListGastoPlantaHistorico = new List<GastoPlantaHistorico>();

                gastoPlantaHistorico = gastoPlantaENTReq.ListGastoPlantaHistorico.FirstOrDefault();

                if (gastoPlantaHistorico != null)
                {
                    gastoPlantaENTRes.ListGastoPlantaHistorico = campanaDAT.MostrarGastoPlantaHistorico(gastoPlantaHistorico);
                }

                gastoPlantaENTRes.Mensaje = "OK";
                gastoPlantaENTRes.OK = 1;
            }
            catch(Exception ex)
            {
                gastoPlantaENTRes.Mensaje = "ERROR: Ocurrio un error inesperado, no se puede mostrar la informacion del historico de Gasto Planta, intente de nuevo o consulte a su administrador.";
                gastoPlantaENTRes.OK = 0;
                gastoPlantaENTRes.ListGastoPlantaHistorico = new List<GastoPlantaHistorico>();

                ArchivoLog.EscribirLog(null, "ERROR: Servicio - MostrarGastoPlantaHistorico, Source: " + ex.Source + " Message: " + ex.Message);
            }

            return gastoPlantaENTRes;
        }
        private void ValidarDatos(ref bool Correcto, List<Cronograma> ListCronograma)
        {
            StringBuilder validacion = new StringBuilder();
            List<Cronograma> ListCronogramaValidacion;

            ListCronogramaValidacion = ListCronograma.Where(n => n.IDTarea == 0
                                                                            || n.IDPadre == 0
                                                                            || string.IsNullOrEmpty(n.Actividad)
                                                                            || n.Duracion == 0
                                                                            || string.IsNullOrEmpty(n.FechaInicio.ToString())
                                                                            || string.IsNullOrEmpty(n.FechaFin.ToString())
                                                                            || n.TiempoOptimista == 0
                                                                            || n.TiempoPesimista == 0
                                                                            || string.IsNullOrEmpty(n.Correo)
                                                                            || string.IsNullOrEmpty(n.Correo_2)
                                                                            || string.IsNullOrEmpty(n.TipoFlujo)
                                                                            || string.IsNullOrEmpty(n.Incluir)).ToList();

            ListCronogramaValidacion.ForEach(n =>
            {
                n.Padre = ListCronograma.Where(m => m.IDPadre == n.IDTarea).Count() > 0 ? true : false;
            });

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validacion.Clear();

                if (cronograma.IDTarea <= 0)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto ID_Tarea");

                    Correcto = false;
                }

                if (cronograma.IDPadre < 0)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto ID_Padre");

                    Correcto = false;
                }

                if (string.IsNullOrEmpty(cronograma.Actividad))
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Actividad");

                    Correcto = false;
                }

                //if (cronograma.Duracion == 0)
                //{
                //    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Duracion");

                //    Correcto = false;
                //}

                if (cronograma.FechaInicio == null || string.IsNullOrEmpty(cronograma.FechaInicio.ToString()))
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Fecha_Inicio");

                    Correcto = false;
                }

                if (cronograma.FechaFin == null || string.IsNullOrEmpty(cronograma.FechaFin.ToString()))
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Fecha_Fin");

                    Correcto = false;
                }

                if (cronograma.TiempoOptimista <= 0 && !cronograma.Padre)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Tiempo_Optimista");

                    Correcto = false;
                }

                if (cronograma.TiempoPesimista <= 0 && !cronograma.Padre)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Tiempo_Pesimista");

                    Correcto = false;
                }

                if (string.IsNullOrEmpty(cronograma.Correo) && !cronograma.Padre)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Correo");

                    Correcto = false;
                }

                if (string.IsNullOrEmpty(cronograma.Correo_2) && !cronograma.Padre)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Correo_2");

                    Correcto = false;
                }

                if (string.IsNullOrEmpty(cronograma.TipoFlujo) && !cronograma.Padre)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Tipo_Flujo");

                    Correcto = false;
                }

                if (string.IsNullOrEmpty(cronograma.Incluir) && !cronograma.Padre)
                {
                    validacion.AppendLine("Error: No tiene informacion o esta incorrecto Incluir");

                    Correcto = false;
                }

                cronograma.ValidarDatos = cronograma.ValidarDatos + validacion.ToString();
            }
        }
        private void ValidarPredecesor(ref bool Correcto, List<Cronograma> ListCronograma)
        {
            List<Cronograma> ListCronogramaValidacion;
            StringBuilder validar = new StringBuilder();

            ListCronogramaValidacion = ListCronograma.Where(n => string.IsNullOrEmpty(n.Predecesor)).ToList();

            ListCronogramaValidacion.ForEach(n =>
            {
                n.Padre = ListCronograma.Where(m => m.IDPadre == n.IDTarea).Count() > 0 ? true : false;
            });

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();
                
                if (string.IsNullOrEmpty(cronograma.Predecesor) && !cronograma.Padre)
                {
                    validar.AppendLine("Precaucion: No se agrego Predecesor, revise si es correcto.");

                    Correcto = false;
                }

                cronograma.ValidarPredecesor = cronograma.ValidarPredecesor + validar.ToString();
            }

        }
        private void ValidarFechasVsActual(ref bool Correcto, List<Cronograma> ListCronograma)
        {
            List<Cronograma> ListCronogramaValidacion;
            StringBuilder validar = new StringBuilder();

            //VALIDAR TIEMPO VS LO ACTUAL
            ListCronogramaValidacion = ListCronograma.Where(n => n.FechaInicio < Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"))).ToList();

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();

                //validar.AppendLine("Precaucion: La Fecha_Inicio programada de la actividad es menor a la actual, por lo cual esta fecha ya aparecera como atrasada en el Cronograma.");

                validar.AppendLine("Precaucion");

                cronograma.ValidarFechaVsActual = cronograma.ValidarFechaVsActual + validar.ToString();

                Correcto = false;
            }

        }
        private void ValidarFechasReales(ref bool Correcto, List<Cronograma> ListCronograma)
        {
            List<Cronograma> ListCronogramaValidacion;
            StringBuilder validar = new StringBuilder();
            Cronograma cronogramaPadre = new Cronograma();

            cronogramaPadre = ListCronograma.Where(n => n.IDTarea == 1 || n.IDPadre == 0).FirstOrDefault();

            if (cronogramaPadre == null)
            {
                validar.Clear();

                //validar.AppendLine("Error: Se debe agregar correctamente la actividad Padre para cargar el archivo de Cronograma.");

                validar.AppendLine("Error");

                ListCronograma.FirstOrDefault().ValidarFecha = validar.ToString();

                Correcto = false;

                return;
            }

            //VALIDAR FECHAS MENORES A LA ACTIVIDAD INICIAL PADRE
            ListCronogramaValidacion = ListCronograma.Where(n => n.FechaInicio.Date < cronogramaPadre.FechaInicio.Date).ToList();

            ListCronogramaValidacion.ForEach(n =>
            {
                n.Padre = ListCronograma.Where(m => m.IDPadre == n.IDTarea).Count() > 0 ? true : false;
            });

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();

                if (cronograma.FechaInicio < cronogramaPadre.FechaInicio && !cronograma.Padre)
                {
                    //validar.AppendLine("Error: La Fecha_Inicio no puede ser menor a la fecha del cronograma.");

                    validar.AppendLine("Error");

                    Correcto = false;
                }

                cronograma.ValidarFecha = cronograma.ValidarFecha + validar.ToString();
            }

            //VALIDAR FECHAS MAYORES A LA ACTIVIDAD FINAL PADRE
            ListCronogramaValidacion = ListCronograma.Where(n => n.FechaFin.Date > cronogramaPadre.FechaFin.Date).ToList();

            ListCronogramaValidacion.ForEach(n =>
            {
                n.Padre = ListCronograma.Where(m => m.IDPadre == n.IDTarea).Count() > 0 ? true : false;
            });

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();
                
                if (cronograma.FechaFin > cronogramaPadre.FechaFin && !cronograma.Padre)
                {
                    //validar.AppendLine("Error: La Fecha_Final no puede ser mayor a la fecha del cronograma.");

                    validar.AppendLine("Error");

                    Correcto = false;
                }

                cronograma.ValidarFecha = cronograma.ValidarFecha + validar.ToString();
            }


            //VALIDAR FECHAS MENORES A LA ACTIVIDAD PADRE
            ListCronogramaValidacion = (from crono in ListCronograma
                                       from crono_2 in ListCronograma
                                       where crono.IDPadre == crono_2.IDTarea
                                        && crono.FechaInicio.Date < crono_2.FechaInicio.Date
                                       select crono).ToList();

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();

                //validar.AppendLine("Error: La Fecha_Inicio no puede ser menor a la fecha de su padre.");

                validar.AppendLine("Error");

                Correcto = false;

                cronograma.ValidarFecha = cronograma.ValidarFecha + validar.ToString();
            }

            //VALIDAR FECHAS MAYORES A LA ACTIVIDAD PADRE
            ListCronogramaValidacion = (from crono in ListCronograma
                                        from crono_2 in ListCronograma
                                        where crono.IDPadre == crono_2.IDTarea
                                         && crono.FechaFin.Date > crono_2.FechaFin.Date
                                        select crono).ToList();

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();

                //validar.AppendLine("Error: La Fecha_Final no puede ser mayor a la fecha de su padre.");

                validar.AppendLine("Error");

                Correcto = false;

                cronograma.ValidarFecha = cronograma.ValidarFecha + validar.ToString();
            }


            //VALIDAR QUE LA FECHA INICIAL NO SEA MAYOR A LA FINAL
            ListCronogramaValidacion = ListCronograma.Where(n => n.FechaInicio.Date > n.FechaFin.Date).ToList();

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();

                //validar.AppendLine("Error: La Fecha_Inicio no puede ser mayor a la Fecha_Final.");

                validar.AppendLine("Error");

                Correcto = false;

                cronograma.ValidarFecha = cronograma.ValidarFecha + validar.ToString();
            }

        }
        private void ValidarFechasProgramadas(ref bool Correcto, List<Cronograma> ListCronograma)
        {
            List<Cronograma> ListCronogramaValidacion;
            StringBuilder validar = new StringBuilder();

            //VALIDAR TIEMPO OPTIMISTA
            ListCronogramaValidacion = ListCronograma.Where(n => !n.Padre && n.Duracion < n.TiempoOptimista).ToList();

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();

                //validar.AppendLine("Precaucion: La Duracion de la actividad es menor al Tiempo_Optimista.");

                validar.AppendLine("Precaucion");

                cronograma.ValidarFechaProgramada = cronograma.ValidarFechaProgramada + validar.ToString();

                Correcto = false;
            }

            //VALIDAR TIEMPO PESIMISTA
            ListCronogramaValidacion = ListCronograma.Where(n => !n.Padre && n.Duracion > n.TiempoPesimista).ToList();

            foreach (Cronograma cronograma in ListCronogramaValidacion)
            {
                validar.Clear();

                //validar.AppendLine("Precaucion: La Duracion de la actividad es mayor al Tiempo_Pesimista.");

                validar.AppendLine("Precaucion");

                cronograma.ValidarFechaProgramada = cronograma.ValidarFechaProgramada + validar.ToString();

                Correcto = false;
            }

        }
        public DateTime FromUnixTime(long unixTimeMillis)
        {
            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return epoch.AddMilliseconds(unixTimeMillis);
        }

        //public void AgregarHijos(int IdCampana, Task Padre, List<object> ListActividades, List<Cronograma> ListCronograma)
        //{
        //    Cronograma cronograma;

        //    string responsable = string.Empty;
        //    string predecesor = string.Empty;

        //    foreach (Task Hijo in ListActividades)
        //    {
        //        cronograma = new Cronograma();

        //        responsable = string.Empty;

        //        if (Hijo.ResourceAssignments.size() > 0)
        //        {
        //            Hijo.ResourceAssignments.toArray().Cast<ResourceAssignment>().ToList().ForEach(n => responsable = responsable + "," + n.Resource.Name);
        //        }

        //        predecesor = string.Empty;

        //        if (Hijo.Predecessors.size() > 0)
        //        {
        //            Hijo.Predecessors.toArray().Cast<Relation>().ToList().ForEach(n => predecesor = predecesor + "," + n.TargetTask.ID.toString());
        //        }

        //        cronograma.IDCampania = IdCampana;
        //        cronograma.IDTarea = Hijo.ID.intValue();
        //        cronograma.IDPadre = Padre.ID.intValue();
        //        cronograma.Actividad = Hijo.Name;
        //        cronograma.NombreResponsable = responsable;
        //        cronograma.PPGID = Hijo.ResourceInitials;
        //        cronograma.Padre = Hijo.ChildTasks.size() > 0 ? true : false;
        //        cronograma.PorcentajeUsuario = Convert.ToDecimal(Hijo.PercentageComplete.floatValue());
        //        cronograma.PorcentajeSistema = 0;
        //        cronograma.Duracion = Hijo.Duration.Duration.ToString();
        //        cronograma.FechaInicio = FromUnixTime(Hijo.Start.getTime()).ToString("dd/MM/yyyy");
        //        cronograma.FechaFin = FromUnixTime(Hijo.Finish.getTime()).ToString("dd/MM/yyyy");
        //        cronograma.FechaHoy = DateTime.Now.ToString("dd/MM/yyyy");
        //        cronograma.FechaCreacion = DateTime.Now.ToString("dd/MM/yyyy");
        //        cronograma.FechaModificacion = DateTime.Now.ToString("dd/MM/yyyy");
        //        cronograma.Comentario = Hijo.Notes;
        //        cronograma.Predecesor = predecesor;
        //        cronograma.IdTreePadre = "treegrid-" + Hijo.ID.intValue() + " treegrid-parent-" + Padre.ID.intValue();
        //        cronograma.VersionCronograma = 0;

        //        ListCronograma.Add(cronograma);

        //        if (Hijo.ChildTasks.size() > 0)
        //        {
        //            AgregarHijos(IdCampana, Hijo, Hijo.ChildTasks.toArray().ToList(), ListCronograma);
        //        }
        //    }
        //}
        
    }
}

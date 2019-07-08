using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace PDFTool
{
    public class PDF
    {
        public static void GenerarPDF(string html, string pathImgHeader, string pathImgFooter, string pathPDF, double pageSizeHeigth, double pageSizeWidth)
        {
            try
            {
                PdfGenerateConfig config = new PdfGenerateConfig();
                config.PageSize = PageSize.Undefined;
                config.ManualPageSize = new XSize(pageSizeWidth, pageSizeHeigth);
                //config.ManualPageSize.Height = pageSizeHeigth;
                //config.ManualPageSize.Width = pageSizeWidth;
                //config.SetMargins(30);
                config.MarginLeft = 15;
                config.MarginRight = 15;
                config.MarginBottom = 30;
                config.MarginTop = 30;

                //string html = File.ReadAllText("C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\Ejemplo_3.html");

                var doc = PdfGenerator.GeneratePdf(html, config);
                
                foreach (PdfPage page in doc.Pages)
                {
                    
                    //page.Height = pageSizeHeigth;
                    //page.Width = pageSizeWidth;
                    
                    double h = page.Height.Value;

                    double w = page.Width.Value;

                    XGraphics graph = XGraphics.FromPdfPage(page);
                    
                    //string pathImgHeader = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\header.png";

                    XImage xImageHeader = XImage.FromFile(pathImgHeader);

                    graph.DrawImage(xImageHeader, 0, 0, w - 10, 20);

                    //string pathImgFooter = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\footer.png";

                    XImage xImageFooter = XImage.FromFile(pathImgFooter);

                    graph.DrawImage(xImageFooter, 10, h - 20, w - 10, 20);
                }

                //var tmpFile = Path.GetTempFileName();
                //tmpFile = Path.GetFileNameWithoutExtension(tmpFile) + ".pdf";

                doc.Save(pathPDF);
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        public static void GenerarPDFCEO(List<string> ListHtml, string pathImgHeader, string pathImgFooter, string pathPDF, double pageSizeHeigth, double pageSizeWidth)
        {
            try
            {
                PdfGenerateConfig config = new PdfGenerateConfig();
                config.PageSize = PageSize.Undefined;
                config.ManualPageSize = new XSize(pageSizeWidth, pageSizeHeigth);
                //config.ManualPageSize.Height = pageSizeHeigth;
                //config.ManualPageSize.Width = pageSizeWidth;
                //config.SetMargins(30);
                config.MarginLeft = 15;
                config.MarginRight = 15;
                config.MarginBottom = 30;
                config.MarginTop = 30;

                //string html = File.ReadAllText("C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\Ejemplo_3.html");

                //var doc = PdfGenerator.GeneratePdf(html, config);

                var doc = new PdfDocument();

                foreach (string html in ListHtml)
                {

                    PdfGenerator.AddPdfPages(doc, html, config);

                    double h = doc.Pages.Cast<PdfPage>().Last().Height.Value;

                    double w = doc.Pages.Cast<PdfPage>().Last().Width.Value;

                    XGraphics graph = XGraphics.FromPdfPage(doc.Pages.Cast<PdfPage>().Last());

                    //string pathImgHeader = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\header.png";

                    XImage xImageHeader = XImage.FromFile(pathImgHeader);

                    graph.DrawImage(xImageHeader, 0, 0, w - 10, 20);

                    //string pathImgFooter = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\footer.png";

                    XImage xImageFooter = XImage.FromFile(pathImgFooter);

                    graph.DrawImage(xImageFooter, 10, h - 20, w - 10, 20);
                }

                //var tmpFile = Path.GetTempFileName();
                //tmpFile = Path.GetFileNameWithoutExtension(tmpFile) + ".pdf";

                doc.Save(pathPDF);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void GenerarPDFMercadotecnia(List<string> ListHtml, string pathImgHeader, string pathImgFooter, string pathPDF, double pageSizeHeigth, double pageSizeWidth)
        {
            try
            {
                PdfGenerateConfig config = new PdfGenerateConfig();
                config.PageSize = PageSize.Undefined;
                config.ManualPageSize = new XSize(pageSizeWidth, pageSizeHeigth);
                //config.ManualPageSize.Height = pageSizeHeigth;
                //config.ManualPageSize.Width = pageSizeWidth;
                //config.SetMargins(30);
                config.MarginLeft = 15;
                config.MarginRight = 15;
                config.MarginBottom = 30;
                config.MarginTop = 30;

                //string html = File.ReadAllText("C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\Ejemplo_3.html");

                //var doc = PdfGenerator.GeneratePdf(html, config);

                var doc = new PdfDocument();

                foreach (string html in ListHtml)
                {
                    
                    PdfGenerator.AddPdfPages(doc, html, config);

                    double h = doc.Pages.Cast<PdfPage>().Last().Height.Value;

                    double w = doc.Pages.Cast<PdfPage>().Last().Width.Value;

                    XGraphics graph = XGraphics.FromPdfPage(doc.Pages.Cast<PdfPage>().Last());

                    //string pathImgHeader = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\header.png";

                    XImage xImageHeader = XImage.FromFile(pathImgHeader);

                    graph.DrawImage(xImageHeader, 0, 0, w - 10, 20);

                    //string pathImgFooter = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\footer.png";

                    XImage xImageFooter = XImage.FromFile(pathImgFooter);

                    graph.DrawImage(xImageFooter, 10, h - 20, w - 10, 20);
                }

                //var tmpFile = Path.GetTempFileName();
                //tmpFile = Path.GetFileNameWithoutExtension(tmpFile) + ".pdf";

                doc.Save(pathPDF);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void GenerarPDFSKU(List<string> ListHtml, string pathImgHeader, string pathImgFooter, string pathPDF, double pageSizeHeigth, double pageSizeWidth)
        {
            try
            {
                PdfGenerateConfig config = new PdfGenerateConfig();
                config.PageSize = PageSize.Undefined;
                config.ManualPageSize = new XSize(pageSizeWidth, pageSizeHeigth);
                //config.ManualPageSize.Height = pageSizeHeigth;
                //config.ManualPageSize.Width = pageSizeWidth;
                //config.SetMargins(30);
                config.MarginLeft = 15;
                config.MarginRight = 15;
                config.MarginBottom = 30;
                config.MarginTop = 30;

                //string html = File.ReadAllText("C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\Ejemplo_3.html");

                //var doc = PdfGenerator.GeneratePdf(html, config);

                var doc = new PdfDocument();

                foreach (string html in ListHtml)
                {

                    PdfGenerator.AddPdfPages(doc, html, config);

                    double h = 792;

                    double w = 670;

                    XGraphics graph = XGraphics.FromPdfPage(doc.Pages.Cast<PdfPage>().Last());

                    //string pathImgHeader = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\header.png";

                    XImage xImageHeader = XImage.FromFile(pathImgHeader);

                    graph.DrawImage(xImageHeader, 0, 0, w - 10, 20);

                    //string pathImgFooter = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\footer.png";

                    XImage xImageFooter = XImage.FromFile(pathImgFooter);

                    graph.DrawImage(xImageFooter, pageSizeWidth - w, pageSizeHeigth - 20, w - 10, 20);
                }

                //var tmpFile = Path.GetTempFileName();
                //tmpFile = Path.GetFileNameWithoutExtension(tmpFile) + ".pdf";

                doc.Save(pathPDF);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

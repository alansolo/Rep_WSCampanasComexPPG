using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
//using System.Linq;
using System.Text;
using TheArtOfDev.HtmlRenderer.PdfSharp;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("UNO");

           
        }

        public void GenerarPDF()
        {
            PdfGenerateConfig config = new PdfGenerateConfig();
            config.PageSize = PageSize.Letter;
            //config.SetMargins(30);
            config.MarginLeft = 15;
            config.MarginRight = 15;
            config.MarginBottom = 30;
            config.MarginTop = 30;

            string html = File.ReadAllText("C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\Ejemplo_3.html");

            var doc = PdfGenerator.GeneratePdf(html, config);

            foreach (PdfPage page in doc.Pages)
            {
                double h = page.Height.Value;

                double w = page.Width.Value;

                XGraphics graph = XGraphics.FromPdfPage(page);

                string pathImgHeader = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\header.png";

                XImage xImageHeader = XImage.FromFile(pathImgHeader);

                graph.DrawImage(xImageHeader, 0, 0, w - 10, 20);

                string pathImgFooter = "C:\\Users\\k697344\\Documents\\Comex PPG\\Documentacion\\footer.png";

                XImage xImageFooter = XImage.FromFile(pathImgFooter);

                graph.DrawImage(xImageFooter, 10, h - 20, w - 10, 20);
            }

            var tmpFile = Path.GetTempFileName();
            tmpFile = Path.GetFileNameWithoutExtension(tmpFile) + ".pdf";

            //var tmpFile = Path.GetFileNameWithoutExtension(pathPDF) + ".pdf";

            doc.Save(tmpFile);

            Process.Start(tmpFile);
        }
    }
}

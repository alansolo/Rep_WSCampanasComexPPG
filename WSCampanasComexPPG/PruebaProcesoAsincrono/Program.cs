using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PruebaProcesoAsincrono
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Cero");



            var t = Task.Run(() =>
            {

                PdfDocument pdf = new PdfDocument();
                pdf.Info.Title = "My First PDF";
                PdfPage pdfPage = pdf.AddPage();
                XGraphics graph = XGraphics.FromPdfPage(pdfPage);
                XFont font = new XFont("Verdana", 20, XFontStyle.Bold);
                graph.DrawString("This is my first PDF document", font, XBrushes.Black, new XRect(0, 0, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.Center);
                string pdfFilename = "firstpage.pdf";
                pdf.Save(pdfFilename);
                Process.Start(pdfFilename);

                DateTime dat = DateTime.Now;
                if (dat == DateTime.MinValue)
                    throw new ArgumentException("The clock is not working.");

                if (dat.Hour > 17)
                    return "evening";
                else if (dat.Hour > 12)
                    return "afternoon";
                else
                    return "morning";

            });
            var c = t.ContinueWith((antecedent) =>
            {
                if (t.Status == TaskStatus.RanToCompletion)
                {
                    Console.WriteLine("Good {0}!",
                                      antecedent.Result);
                    Console.WriteLine("And how are you this fine {0}?",
                                   antecedent.Result);
                }
                else if (t.Status == TaskStatus.Faulted)
                {
                    Console.WriteLine(t.Exception.GetBaseException().Message);
                }
            });



            //t.Start();

            //PDFAsync();

            Console.WriteLine("Uno");

            Console.Read();
        }

        static async Task PDFAsync()
        {
            await Task.Run(() => 
            {
                PdfDocument pdf = new PdfDocument();
                pdf.Info.Title = "My First PDF";
                PdfPage pdfPage = pdf.AddPage();
                XGraphics graph = XGraphics.FromPdfPage(pdfPage);
                XFont font = new XFont("Verdana", 20, XFontStyle.Bold);
                graph.DrawString("This is my first PDF document", font, XBrushes.Black, new XRect(0, 0, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.Center);
                string pdfFilename = "firstpage.pdf";
                pdf.Save(pdfFilename);
                Process.Start(pdfFilename);
            });
           
        }

        public static async Task Pro()
        {
            // Execute the antecedent.
            Task<DayOfWeek> taskA = Task.Run(() =>
            {
                PdfDocument pdf = new PdfDocument();
                pdf.Info.Title = "My First PDF";
                PdfPage pdfPage = pdf.AddPage();
                XGraphics graph = XGraphics.FromPdfPage(pdfPage);
                XFont font = new XFont("Verdana", 20, XFontStyle.Bold);
                graph.DrawString("This is my first PDF document", font, XBrushes.Black, new XRect(0, 0, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.Center);
                string pdfFilename = "firstpage.pdf";
                pdf.Save(pdfFilename);
                Process.Start(pdfFilename);

                return DateTime.Today.DayOfWeek;
            });

            // Execute the continuation when the antecedent finishes.
            await taskA.ContinueWith(antecedent => Console.WriteLine("Today is {0}.", antecedent.Result));
        }

        public static async Task Proceso1Async()
        {
            Task t = Task.Run(() =>
            {
                //Proceso1();

                //Thread.Sleep(5000);

                for (var i = 0; i < 10; i++)
                {
                    Console.WriteLine("do something {0}", i + 1);
                }
            });

            await t.ContinueWith(x =>
            {
                //Proceso2();

                //Thread.Sleep(1000);

                Console.WriteLine("done.");
            });
        }

        public static void Proceso2()
        {
            Thread.Sleep(1000);
        }
    }
}

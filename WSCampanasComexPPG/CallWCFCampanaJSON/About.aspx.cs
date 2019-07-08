using LinqToExcel;
using Microsoft.Office.Interop.Excel;
using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace CallWCFCampanaJSON
{
    public partial class About : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void Unnamed1_Click(object sender, EventArgs e)
        {
            var un = new ExcelQueryFactory("C:\\Users\\k697344\\Documents\\Libro3.xlsx");

            var dos = new ExcelQueryFactory("https://one.web.ppg.com/la/comex/camp_publicidad/DocumentosCampaas/Cronogramas/DatosBDWizard.xlsx");

            string url = HttpUtility.HtmlEncode("https://one.web.ppg.com/la/comex/camp_publicidad/DocumentosCampaas/Cronogramas/DatosBDWizard.xlsx");

            var excel = new Application();
            excel.Visible = false;
            string workbookPath = url;

            Workbook excelWorkbook = excel.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                                                            true, false, 0, true, false, false);

            excelWorkbook.SaveAs("C:\\Users\\k697344\\Documents\\Libro22.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, 
                                    Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

            excelWorkbook.Close(true, "C:\\Users\\k697344\\Documents\\Libro22.xlsx", false);

            excelWorkbook.Close();

            Sheets sheets = excelWorkbook.Worksheets;

            Worksheet worksheet = sheets.Cast<Worksheet>().Where(n => n.Name == "Hoja1").FirstOrDefault();

            //Worksheet worksheet = (Worksheet)sheets.Item[6];

            //bool flag = false;

            
            var data2 = worksheet.ListObjects.Cast<Object>().FirstOrDefault();

            Range range = worksheet.Range["", ""];

            var values = (System.Array)range.Cells.Value2;

            

            var book = new ExcelQueryFactory(workbookPath);

            var datos = book.Worksheet(6).AsEnumerable().Select(row => new
            {
                uno = row["uno"],
                dos = row["dos"]
            }).ToList();

        }
    }
}
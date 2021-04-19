using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace ServiceConnector
{
    internal class ExcelExporter
    {
        internal class ExportSettings
        {
            public string ExportDirectoryName { get; set; }
            public string ExportToFileName { get; set; }
            public bool EnableExportToSubFolder { get; set; }
            public string ExportToSubFolderName { get; set; }
        }

        internal string ExportToExcel<T>(List<T> responseData, ExportSettings exportSettings)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                excelApp.Visible = true;

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();

                string output = new ExportHelper().SetWorksheet(responseData, excelWorksheet);
                if (!string.IsNullOrEmpty(output))
                {
                    return output;
                }

                string currentDirectory = Environment.CurrentDirectory;
                if (!Directory.Exists(currentDirectory + exportSettings.ExportDirectoryName))
                {
                    Directory.CreateDirectory(currentDirectory + exportSettings.ExportDirectoryName);
                }
                if (exportSettings.EnableExportToSubFolder && !Directory.Exists(currentDirectory + exportSettings.ExportDirectoryName + exportSettings.ExportToSubFolderName))
                {
                    Directory.CreateDirectory(currentDirectory + exportSettings.ExportDirectoryName + exportSettings.ExportToSubFolderName);
                }

                excelApp.DisplayAlerts = false;
                try
                {
                    excelApp.ActiveWorkbook.SaveAs(currentDirectory + exportSettings.ExportDirectoryName +
                        (exportSettings.EnableExportToSubFolder ? exportSettings.ExportToSubFolderName : String.Empty)
                        + exportSettings.ExportToFileName, Excel.XlFileFormat.xlWorkbookNormal);
                }
                catch (Exception err)
                {
                    if (err.Message.ToLower().Contains("cannot access"))
                    {
                        return $"Unable to save file: {exportSettings.ExportToFileName}. Please close the existing version and try process again.";
                    }
                }                

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return string.Empty;
        }
    }
}

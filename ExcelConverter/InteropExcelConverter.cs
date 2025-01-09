using ExcelConverter.Interfaces;
using Logging;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;

namespace ExcelConverter
{
    internal sealed class InteropExcelConverter : IExcelConverter<string>
    {
        private InteropExcelApp Excel { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public InteropExcelConverter(InteropExcelApp excelApp, IProgress<object[]> progress)
        {
            Excel = excelApp;
            Progress = progress;
        }
        public void ConvertExcelFilesToXLSX(string filesFolder)
        {
            int counter = 0;
            string[] xlsFiles = Directory.GetFiles(filesFolder, "*.xls");
            
            foreach (string file in xlsFiles)
            {
                try
                {
                    Progress.Report([xlsFiles.Count(), ++counter, "Конвертация файлов *.xls  в  *.xlsx", false]);
                    XlsToXlsx(file);
                    RemoveXlsFile(file);
                }
                catch (SystemException ex)
                {
                    Logger.Log(ex);
                }
            }
            Excel.Dispose();
        }
        private void XlsToXlsx(string filePath)
        {
            string outputPath = filePath + "x";

            Excel.App.IgnoreRemoteRequests = true;
            Workbook workbook =  Excel.App.Workbooks.Open(filePath, 
                                                          UpdateLinks:0,
                                                          Missing.Value, 
                                                          Missing.Value, 
                                                          Missing.Value, 
                                                          Missing.Value, 
                                                          IgnoreReadOnlyRecommended: true);
            workbook.SaveAs(outputPath, FileFormat: XlFileFormat.xlOpenXMLWorkbook); 
            workbook.Close(SaveChanges: false, Missing.Value, RouteWorkbook: false);
            Excel.App.IgnoreRemoteRequests = false;
        }
        private void RemoveXlsFile(string filePath)
        {
            File.Delete(filePath);
        }
    }
}

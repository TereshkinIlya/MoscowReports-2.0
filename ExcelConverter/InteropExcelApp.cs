using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelConverter
{
    internal class InteropExcelApp : IDisposable
    {
        public Excel.Application App { get; set; }
        public InteropExcelApp(Excel.Application excelApp) 
        {
            App = excelApp;
        }

        public void Dispose()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (App != null)
            {
                App.DisplayAlerts = true;
                App.Visible = true;
                
                App.Quit();
                Marshal.ReleaseComObject(App);
            }
        }
    }
}

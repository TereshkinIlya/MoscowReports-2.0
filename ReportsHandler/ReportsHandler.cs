using ClosedXML.Excel;
using Dto;
using Logging;
using ReportsHandler.Interfaces;
using System.Diagnostics.Metrics;
using System.IO;

namespace ReportsHandler
{
    internal sealed class ReportsHandler : IReportHandler<string>
    {
        private MainFormDto MainForm { get; set; }
        private IExcel<XLWorkbook> Excel { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public ReportsHandler(IExcel<XLWorkbook> excel, IProgress<object[]> progress)
        {
            MainForm = MainFormDto.GetInstance();
            Excel = excel;
            Progress = progress;
        }
        public void HandleReports(string filesFolder)
        {
            ReadPiktsTable();
            ReadReports(filesFolder);
            WriteInMoscowTable();
        }
        private void ReadPiktsTable()
        {
            if (MainForm.OnlyReportsA == false)
            {
                Progress.Report([100, 0, "Чтение ПИКТС таблицы", true]);
                string path = MainForm.PiktsTablePath;
                try
                {
                    using XLWorkbook workbook = Excel.Open(path);
                    Excel.Read(workbook);
                }
                catch (ArgumentException ex)
                {
                    Logger.Log($"Файл - {Path.GetFileName(path)} : {ex.Message}");
                }
            }
        }
        private void ReadReports(string filesFolder)
        {
            int counter = 0;
            string[] files = Directory.GetFiles(filesFolder);
            foreach (string path in files)
            {
                Progress.Report([files.Count(), ++counter, "Чтение файлов Приложений А", false]);
                try
                {
                    using XLWorkbook workbook = Excel.Open(path);
                    workbook.Properties.Title = Path.GetFullPath(path);
                    Excel.Read(workbook);
                }
                catch (ArgumentException ex)
                {
                    Logger.Log(ex.Message);
                }
            }
        }
        private void WriteInMoscowTable()
        {
            string path = MainForm.MoscowTablePath;
            try
            {
                using XLWorkbook table = Excel.Open(path);
                Excel.Write(table);
                table.Save();
            }
            catch (ArgumentException ex)
            {
                Logger.Log($"Файл - {Path.GetFileName(path)} : {ex.Message}");
            }
        }
    }
}

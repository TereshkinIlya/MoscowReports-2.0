using ExcelConverter.Interfaces;
using FileManager.Interfaces;
using Infrastructure.Interfaces;
using Logging;
using ReportsHandler.Interfaces;
using TableHeadersHandler.Interfaces;

namespace Infrastructure.Launcher
{
    public sealed class AppLauncher : ILauncher
    {
        private ISpreadsheet Table { get; set; }
        private IFileManager<string> FileManager { get; set; }
        private IExcelConverter<string> ExcelConverter { get; set; }
        private IReportHandler<string> ReportHandler { get; set; }
        public AppLauncher(ISpreadsheet table,
                           IFileManager<string> fileManager,
                           IExcelConverter<string> excelConverter,
                           IReportHandler<string> reportHandler)
        {
            Table = table;
            FileManager = fileManager;
            ExcelConverter = excelConverter;
            ReportHandler = reportHandler;
        }
        public void Run()
        {
            Table.InitializeHeader();
            string reportsFolder = FileManager.GetFolderWithReports();
            ExcelConverter.ConvertExcelFilesToXLSX(reportsFolder);
            ReportHandler.HandleReports(reportsFolder);
        }
    }
}

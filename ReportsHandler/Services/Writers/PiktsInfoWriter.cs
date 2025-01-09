using ClosedXML.Excel;
using Core.Abstracts;
using Core.Data.PiktsInfo;
using Core.Interaces;
using Core.Tables.HeaderColumns;
using Microsoft.Extensions.DependencyInjection;
using ReportsHandler.Abstracts;
using ReportsHandler.Interfaces;

namespace ReportsHandler.Services.Writers
{
    internal sealed class PiktsInfoWriter : IExcelWriter<XLWorkbook>
    {
        private Repository<Queue<Content>, Content> Repository { get; set; }
        private IHeader TableHeader { get; set; }
        private ReaderBase ReaderBase { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public PiktsInfoWriter([FromKeyedServices("piktsRepo")] Repository<Queue<Content>, Content> repository,
                               IHeader tableHeader,
                               IProgress<object[]> progress)
        {
            Repository = repository;
            TableHeader = tableHeader;
            ReaderBase = new ReaderBase();
            Progress = progress;
        }
        public void Write(XLWorkbook table)
        {
            IXLWorksheet? sheet = table.Worksheets.FirstOrDefault(sheet =>
                sheet.Name.Contains("ППМН ТН", StringComparison.CurrentCultureIgnoreCase));

            if (sheet is null)
            {
                return;
            }
            else
            {
                WriteAsPikts(table);
            }
        }
        private void WriteAsPikts(XLWorkbook table)
        {
            int counter = 0;
            int total = Repository.Count();
            while (!Repository.IsEmpty())
            {
                Progress.Report([total, ++counter, "Запись ПИКТС данных в Московскую таблицу", false]);

                var piktsInfo = (Pikts)Repository.Get();
                WritePiktsInfoInTable(table, piktsInfo);
            }
        }
        private void WritePiktsInfoInTable(XLWorkbook table, Pikts piktsInfo)
        {
            int row;
            int column;
            int quantDefects;
            bool isDate;
            DateTime date;

            BigStreamsColumns columns = (BigStreamsColumns)TableHeader.GetHeader<BigStreamsColumns>();
            IXLWorksheet worksheet = ReaderBase.GetWorksheet(table, "ППМН ТН");
            worksheet.SetAutoFilter(false);
            IXLRange ids = ReaderBase.GetIdsRange(worksheet, "ID");
            IXLCell? idCell = ids.Search(piktsInfo.Id.ToString()).FirstOrDefault();

            if (idCell is null) return;

            row = idCell.WorksheetRow().RowNumber();

            column = columns["State"];
            worksheet.Cell(row, column).Value = piktsInfo.State;

            column = columns["SteelGrade"];
            worksheet.Cell(row, column).Value = piktsInfo.SteelGrade;

            column = columns["SteelGrade"];
            worksheet.Cell(row, column).Value = piktsInfo.SteelGrade;

            column = columns["DateLastVTD"];
            isDate = DateTime.TryParse(piktsInfo.DateLastVTD, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.DateLastVTD;

            column = columns["Defects.CurrentYear.FloodplaneDefects"];
            worksheet.Cell(row, column).Clear(XLClearOptions.Contents);
            quantDefects = piktsInfo.Defects.CurrentYear.FloodplaneDefects.Quantity;
            if (quantDefects > 0) worksheet.Cell(row, column).Value = quantDefects;

            column = columns["Defects.CurrentYear.RiverbedDefects"];
            worksheet.Cell(row, column).Clear(XLClearOptions.Contents);
            quantDefects = piktsInfo.Defects.CurrentYear.RiverbedDefects.Quantity;
            if (quantDefects > 0) worksheet.Cell(row, column).Value = quantDefects;

            column = columns["Defects.CurrentYearPlusOne.FloodplaneDefects"];
            worksheet.Cell(row, column).Clear(XLClearOptions.Contents);
            quantDefects = piktsInfo.Defects.CurrentYearPlusOne.FloodplaneDefects.Quantity;
            if (quantDefects > 0) worksheet.Cell(row, column).Value = quantDefects;

            column = columns["Defects.CurrentYearPlusOne.RiverbedDefects"];
            worksheet.Cell(row, column).Clear(XLClearOptions.Contents);
            quantDefects = piktsInfo.Defects.CurrentYearPlusOne.RiverbedDefects.Quantity;
            if (quantDefects > 0) worksheet.Cell(row, column).Value = quantDefects;

            column = columns["Defects.CurrentYearPlusTwo.FloodplaneDefects"];
            worksheet.Cell(row, column).Clear(XLClearOptions.Contents);
            quantDefects = piktsInfo.Defects.CurrentYearPlusTwo.FloodplaneDefects.Quantity;
            if (quantDefects > 0) worksheet.Cell(row, column).Value = quantDefects;

            column = columns["Defects.CurrentYearPlusTwo.RiverbedDefects"];
            worksheet.Cell(row, column).Clear(XLClearOptions.Contents);
            quantDefects = piktsInfo.Defects.CurrentYearPlusTwo.RiverbedDefects.Quantity;
            if (quantDefects > 0) worksheet.Cell(row, column).Value = quantDefects;

            column = columns["SafePeriod"];
            isDate = DateTime.TryParse(piktsInfo.SafePeriod, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.SafePeriod;

            column = columns["DateReport"];
            isDate = DateTime.TryParse(piktsInfo.DateReport, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.DateReport;

            column = columns["OTSNumber"];
            worksheet.Cell(row, column).Value = piktsInfo.OTSNumber;

            column = columns["PeriodSafeOperation.DateVTD"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateVTD, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateVTD;

            column = columns["PeriodSafeOperation.DateDefect"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateDefect, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateDefect;

            column = columns["PeriodSafeOperation.DateCDS"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateCDS, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateCDS;

            column = columns["PeriodSafeOperation.DateJumpers"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateJumpers, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateJumpers;

            column = columns["PeriodSafeOperation.DateLimited"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateLimited, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateLimited;

            column = columns["PeriodSafeOperation.DateVRK"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateVRK, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateVRK;

            column = columns["PeriodSafeOperation.DateUZA"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateUZA, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateUZA;

            column = columns["PeriodSafeOperation.DateWeldedElement"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateWeldedElement, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateWeldedElement;

            column = columns["PeriodSafeOperation.DateConnectedDetails"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateConnectedDetails, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateConnectedDetails;

            column = columns["PeriodSafeOperation.DateKPSOD"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateKPSOD, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateKPSOD;

            column = columns["PeriodSafeOperation.DateDrainageContainers"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateDrainageContainers, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateDrainageContainers;

            column = columns["PeriodSafeOperation.DatePVP"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DatePVP, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DatePVP;

            column = columns["PeriodSafeOperation.DateCorrosion"];
            isDate = DateTime.TryParse(piktsInfo.PeriodSafeOperation.DateCorrosion, out date);
            worksheet.Cell(row, column).Value = isDate ? date : piktsInfo.PeriodSafeOperation.DateCorrosion;

            column = columns["Organization"];
            worksheet.Cell(row, column).Value = piktsInfo.Organization;

            column = columns["Measures"];
            worksheet.Cell(row, column).Value = piktsInfo.Measures;
        }
    }
}

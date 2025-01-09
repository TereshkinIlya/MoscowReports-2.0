using ClosedXML.Excel;
using Core.Abstracts;
using Microsoft.Extensions.DependencyInjection;
using ReportsHandler.Interfaces;
using ReportsHandler.Abstracts;
using Core.Interaces;
using Core.Data.PiktsInfo;

namespace ReportsHandler.Services.Readers
{
    internal sealed class PiktsReportReader : IExcelReader<XLWorkbook>
    {
        private Repository<Queue<Content>, Content> Repository { get; set; }
        private IData DataFactory { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public ReaderBase ReaderBase { get; set; }
        private Dictionary<string, int> Columns { get; set; } = new();
        public PiktsReportReader([FromKeyedServices("piktsRepo")] Repository<Queue<Content>, Content> repository,
                                 IData dataFactory,
                                 IProgress<object[]> progress)
        {
            Repository = repository;
            DataFactory = dataFactory;
            Progress = progress;
            ReaderBase = new ReaderBase();
        }

        public void Read(XLWorkbook workbook)
        {
            IXLWorksheet? sheet = workbook.Worksheets.FirstOrDefault(sheet =>
                sheet.Name.Contains("Отчет ПАО", StringComparison.CurrentCultureIgnoreCase));

            if (sheet is not null)
            {
                ReadAsPikts(workbook);
            }
        }
        public void ReadAsPikts(XLWorkbook workbook)
        {
            if (!Repository.IsEmpty()) return;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, "Отчет ПАО");
            worksheet.SetAutoFilter(false);
            IXLRange table = GetPiktsTableRange(worksheet);
            IEnumerable<string> ids = GetIDsFromPiktsTable(worksheet);

            SetColumnsNumbers(worksheet);
            ReadPiktsTable(worksheet, table, ids);
        }
        private IXLRange GetPiktsTableRange(IXLWorksheet worksheet)
        {
            worksheet.SetAutoFilter(false);

            IXLCell firstUsedCell = worksheet.Cell("A1");
            IXLCell lastUsedCell = worksheet.LastCellUsed();
            IXLRange table = worksheet.Range(firstUsedCell, lastUsedCell);

            return table;
        }
        private IEnumerable<string> GetIDsFromPiktsTable(IXLWorksheet worksheet)
        {
            IXLCell startCell = ReaderBase.GetEntryCell(worksheet, "ID");
            IXLCell endCell = worksheet.Cell(worksheet.LastRowUsed()!.RowNumber(), startCell.WorksheetColumn().ColumnNumber());
            IXLRange idRange = worksheet.Range(startCell, endCell);

            var ids = idRange.AsTable().DataRange.Rows()
                .Select(id => id.Cell(1))
                .Where(cell => !cell.IsEmpty())
                .Select(id => id.GetString().Trim())
                .Where(id => id.Length == 6)
                .Distinct();

            return ids;
        }
        private void ReadPiktsTable(IXLWorksheet worksheet, IXLRange table, IEnumerable<string> ids)
        {
            int counter = 0;
            foreach (var id in ids)
            {
                Progress.Report([ids.Count(), ++counter, "Чтение ПИКТС таблицы", false]);

                Pikts piktsInfo = (Pikts)DataFactory.GetInstance<Pikts>();

                var row = table.AsTable().Rows().Where(row => row.Cell(Columns["idColumn"]).GetString().Trim() == id).First();

                piktsInfo.Id = int.Parse(id);
                piktsInfo.State = row.Cell(Columns["stateMtColumn"]).GetString(); 
                piktsInfo.SteelGrade = row.Cell(Columns["steelColumn"]).GetString();
                piktsInfo.DateLastVTD = row.Cell(Columns["lastVtdColumn"]).GetString();
                piktsInfo.SafePeriod = row.Cell(Columns["safePeriodColumn"]).GetString();
                piktsInfo.OTSNumber = row.Cell(Columns["otsNumberColumn"]).GetString();
                piktsInfo.DateReport = row.Cell(Columns["dateReportColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateVTD = row.Cell(Columns["dateVtdColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateDefect = row.Cell(Columns["dateDefectColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateCDS = row.Cell(Columns["dateCdsColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateJumpers = row.Cell(Columns["dateJumpersColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateLimited = row.Cell(Columns["dateLimitedColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateVRK = row.Cell(Columns["dateVrkColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateUZA = row.Cell(Columns["dateUzaColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateWeldedElement = row.Cell(Columns["dateWeldedElementColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateConnectedDetails = row.Cell(Columns["dateConnectedDetailsColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateKPSOD = row.Cell(Columns["dateKpsodColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateDrainageContainers = row.Cell(Columns["dateDrainageContainersColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DatePVP = row.Cell(Columns["datePvpColumn"]).GetString();
                piktsInfo.PeriodSafeOperation.DateCorrosion = row.Cell(Columns["dateCorrosionColumn"]).GetString();
                piktsInfo.Organization = row.Cell(Columns["organizationColumn"]).GetString();
                piktsInfo.Measures = row.Cell(Columns["measuresColumn"]).IsEmpty() ? "" : "Да";

                CountQuantityDefects(table, id, piktsInfo);

                Repository.Put(piktsInfo);
            }
        }
        private void CountQuantityDefects(IXLRange table, string id, Pikts piktsInfo)
        {
            string curYear = DateTime.Now.Year.ToString();
            string curYearPlusOne = (DateTime.Now.Year + 1).ToString();
            string curYearPlusTwo = (DateTime.Now.Year + 2).ToString();

            piktsInfo.Defects.CurrentYear.FloodplaneDefects.Quantity = CountDefects(table, id, curYear, "пойма");
            piktsInfo.Defects.CurrentYear.RiverbedDefects.Quantity = CountDefects(table, id, curYear, "русло");

            piktsInfo.Defects.CurrentYearPlusOne.FloodplaneDefects.Quantity = CountDefects(table, id, curYearPlusOne, "пойма");
            piktsInfo.Defects.CurrentYearPlusOne.RiverbedDefects.Quantity = CountDefects(table, id, curYearPlusOne, "русло");

            piktsInfo.Defects.CurrentYearPlusTwo.FloodplaneDefects.Quantity = CountDefects(table, id, curYearPlusTwo, "пойма");
            piktsInfo.Defects.CurrentYearPlusTwo.RiverbedDefects.Quantity = CountDefects(table, id, curYearPlusTwo, "русло");
        }
        private int CountDefects(IXLRange table, string id, string year, string position)
        {
            var rows = table.AsTable().Rows().Where(row => row.Cell(Columns["idColumn"]).GetString().Trim() == id);
            var quantity = rows
                .Where(row => !row.Cell(Columns["limitDateColumn"]).IsEmpty() &&
                              !row.Cell(Columns["signRepairColumn"]).IsEmpty() &&
                              !row.Cell(Columns["defectPositionColumn"]).IsEmpty())
                .Count(row => row.Cell(Columns["limitDateColumn"]).GetString().Contains(year) &&
                       row.Cell(Columns["signRepairColumn"]).GetString().Equals("Без ремонта", StringComparison.OrdinalIgnoreCase) &&
                       row.Cell(Columns["defectPositionColumn"]).GetString().Equals(position, StringComparison.OrdinalIgnoreCase));

            return quantity;
        }
        private void SetColumnsNumbers(IXLWorksheet worksheet)
        {
            int idColumn = ReaderBase.GetEntryCell(worksheet, "ID").WorksheetColumn().ColumnNumber();
            int stateMtColumn = ReaderBase.GetEntryCell(worksheet, "Состояние нитки").WorksheetColumn().ColumnNumber();
            int steelColumn = ReaderBase.GetEntryCell(worksheet, "Характеристика труб руслового участка").WorksheetColumn().ColumnNumber();
            int lastVtdColumn = ReaderBase.GetEntryCell(worksheet, "Дата прогона").WorksheetColumn().ColumnNumber();
            int safePeriodColumn = ReaderBase.GetEntryCell(worksheet, "Дата окончания срока безопасной эксплуатации").WorksheetColumn().ColumnNumber();
            int otsNumberColumn = ReaderBase.GetEntryCell(worksheet, "Номер заключения").WorksheetColumn().ColumnNumber();
            int dateReportColumn = ReaderBase.GetEntryCell(worksheet, "Дата выдачи заключения").WorksheetColumn().ColumnNumber();
            int dateVtdColumn = ReaderBase.GetEntryCell(worksheet, "Внутритрубное диагностирование").WorksheetColumn().ColumnNumber();
            int dateDefectColumn = ReaderBase.GetEntryCell(worksheet, "Дефекты").WorksheetColumn().ColumnNumber();
            int dateCdsColumn = ReaderBase.GetEntryCell(worksheet, "Спиралешовные трубы").WorksheetColumn().ColumnNumber();
            int dateJumpersColumn = ReaderBase.GetEntryCell(worksheet, "не подлежащие ВТД").WorksheetColumn().ColumnNumber();
            int dateLimitedColumn = ReaderBase.GetEntryCell(worksheet, "Объекты с ограниченными").WorksheetColumn().ColumnNumber();
            int dateVrkColumn = ReaderBase.GetEntryCell(worksheet, "Временные ремонтные").WorksheetColumn().ColumnNumber();
            int dateUzaColumn = ReaderBase.GetEntryCell(worksheet, "Запорная арматура").WorksheetColumn().ColumnNumber();
            int dateWeldedElementColumn = ReaderBase.GetEntryCell(worksheet, "Приварные элементы").WorksheetColumn().ColumnNumber();
            int dateConnectedDetailsColumn = ReaderBase.GetEntryCell(worksheet, "Соединительные детали").WorksheetColumn().ColumnNumber();
            int dateKpsodColumn = ReaderBase.GetEntryCell(worksheet, "КПП СОД").WorksheetColumn().ColumnNumber();
            int dateDrainageContainersColumn = ReaderBase.GetEntryCell(worksheet, "Дренажные емкости").WorksheetColumn().ColumnNumber();
            int datePvpColumn = ReaderBase.GetEntryCell(worksheet, "Планово-высотное").WorksheetColumn().ColumnNumber();
            int dateCorrosionColumn = ReaderBase.GetEntryCell(worksheet, "Коррозионное состояние").WorksheetColumn().ColumnNumber();
            int organizationColumn = ReaderBase.GetEntryCell(worksheet, "проводившая экспертную").WorksheetColumn().ColumnNumber();
            int measuresColumn = ReaderBase.GetEntryCell(worksheet, "Рекомендации по приведению").WorksheetColumn().ColumnNumber();
            int limitDateColumn = ReaderBase.GetEntryCell(worksheet, "Предельная дата").WorksheetColumn().ColumnNumber();
            int defectPositionColumn = ReaderBase.GetEntryCell(worksheet, "Местоположение").WorksheetColumn().ColumnNumber();
            int signRepairColumn = ReaderBase.GetEntryCell(worksheet, "Признак ремонта").WorksheetColumn().ColumnNumber();

            Columns.Add("idColumn", idColumn);
            Columns.Add("stateMtColumn", stateMtColumn);
            Columns.Add("steelColumn", steelColumn);
            Columns.Add("lastVtdColumn", lastVtdColumn);
            Columns.Add("safePeriodColumn", safePeriodColumn);
            Columns.Add("otsNumberColumn", otsNumberColumn);
            Columns.Add("dateReportColumn", dateReportColumn);
            Columns.Add("dateVtdColumn", dateVtdColumn);
            Columns.Add("dateDefectColumn", dateDefectColumn);
            Columns.Add("dateCdsColumn", dateCdsColumn);
            Columns.Add("dateJumpersColumn", dateJumpersColumn);
            Columns.Add("dateLimitedColumn", dateLimitedColumn);
            Columns.Add("dateVrkColumn", dateVrkColumn);
            Columns.Add("dateUzaColumn", dateUzaColumn);
            Columns.Add("dateWeldedElementColumn", dateWeldedElementColumn);
            Columns.Add("dateConnectedDetailsColumn", dateConnectedDetailsColumn);
            Columns.Add("dateKpsodColumn", dateKpsodColumn);
            Columns.Add("dateDrainageContainersColumn", dateDrainageContainersColumn);
            Columns.Add("datePvpColumn", datePvpColumn);
            Columns.Add("dateCorrosionColumn", dateCorrosionColumn);
            Columns.Add("organizationColumn", organizationColumn);
            Columns.Add("measuresColumn", measuresColumn);
            Columns.Add("limitDateColumn", limitDateColumn);
            Columns.Add("defectPositionColumn", defectPositionColumn);
            Columns.Add("signRepairColumn", signRepairColumn);
        }
    }
}

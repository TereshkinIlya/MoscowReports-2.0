using ClosedXML.Excel;
using Core.Abstracts;
using Core.Data.Watercourse;
using Core.Data.Watercourse.Parts;
using Core.Interaces;
using Microsoft.Extensions.DependencyInjection;
using ReportsHandler.Abstracts;
using ReportsHandler.Interfaces;
using System.Reflection;

namespace ReportsHandler.Services.Readers
{
    internal sealed class SmallReportReader : IExcelReader<XLWorkbook>
    {
        private Repository<Queue<Content>, Content> Repository { get; set; }
        private IData DataFactory { get; set; }
        private ReaderBase ReaderBase { get; set; }
        public SmallReportReader([FromKeyedServices("smallsRepo")] Repository<Queue<Content>, Content> repository,
                                 IData dataFactory)
        {
            Repository = repository;
            DataFactory = dataFactory;
            ReaderBase = new ReaderBase();
        }
        public void Read(XLWorkbook workbook)
        {
            IXLWorksheet? sheet = workbook.Worksheets.FirstOrDefault(sheet =>
                sheet.Name.Contains("УЗА", StringComparison.CurrentCultureIgnoreCase));

            if (sheet is null && workbook.Worksheets.Count() > 2)
            {
                ReadAsSmall(workbook);
            }
        }
        private void ReadAsSmall(XLWorkbook workbook)
        {
            var watercourse = (Watercourse)DataFactory.GetInstance<Watercourse>();
            ReadSheet_1(workbook, watercourse, "Общие сведения");
            ReadSheet_4(workbook, watercourse, "хар-ка");
            ReadSheet_6(workbook, watercourse, "Ремонты");
            ReadSheet_9(workbook, watercourse, "ФГХ");
            ReadSheet_13(workbook, watercourse, "ПВП");
            ReadSheet_15_1(workbook, watercourse, "недозаглубления");
            ReadSheet_15_2(workbook, watercourse, "оголения");
            ReadSheet_15_3(workbook, watercourse, "провисы");
            SetPositionMT(watercourse);

            Repository.Put(watercourse);
        }
        private void ReadSheet_1(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 2;
            int offsetColumn = 0;
            DateTime date;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);
            IXLCell cell = ReaderBase.GetEntryCell(worksheet, "Дата обследования");

            _ = ReaderBase.TryGetRequiredCellValue<DateTime>(cell, offsetRow, offsetColumn, out date);
            if (date == default)
            {
                throw new ArgumentException($"Вкладка {cell.Worksheet}. В ячейке {cell.Address} отсутствует дата обследования перехода");
            }
            else
            {
                watercourse.DateInspection = date;
            }
        }
        private void ReadSheet_4(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 0;
            int offsetColumn = 1;
            int id;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);
            IXLCell cell = ReaderBase.GetEntryCell(worksheet, "ID");

            _ = ReaderBase.TryGetRequiredCellValue<int>(cell, offsetRow, offsetColumn, out id);
            var digitsCount = Math.Floor(Math.Log10(id) + 1);

            if (id == 0 || digitsCount != 6)
            {
                throw new ArgumentException($"Вкладка {cell.Worksheet}. В ячейке {cell.Address} отсутствует или некорректный ID перехода");
            }
            else
            {
                watercourse.Id = id;
            }
        }
        private void ReadSheet_6(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            string repairInfo;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);
            IXLCell startCell = ReaderBase.GetEntryCell(worksheet, "Наименование выполненных работ").CellBelow(2);
            IXLCell endCell = startCell.WorksheetColumn().LastCellUsed();
            IXLRange repairRange = worksheet.Range(startCell, endCell);

            IEnumerable<string> info = repairRange.Cells()
                .Where(cell => !cell.IsEmpty())
                .Select(cell => cell.Value.ToString())
                .Where(cell => cell.Length > 1);
            
            repairInfo = string.Join(". ", info);

            watercourse.RepairInfo = repairInfo;
        }
        private void ReadSheet_9(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 0;
            int offsetColumn = 1;
            bool success;
            string coordinate;
            IXLCell cell;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);

            cell = ReaderBase.GetEntryCell(worksheet, "Широта");
            success = ReaderBase.TryGetRequiredCellValue<string>(cell, offsetRow, offsetColumn, out coordinate);
            if (success)
            {
                watercourse.Latitude = ReaderBase.GetGeoCoordinate(cell, coordinate);
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Долгота");
            success = ReaderBase.TryGetRequiredCellValue<string>(cell, offsetRow, offsetColumn, out coordinate);
            if (success)
            {
                watercourse.Longitude = ReaderBase.GetGeoCoordinate(cell, coordinate);
            }
        }
        private void ReadSheet_13(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 0;
            int offsetColumn = 1;
            float value;
            bool success;
            IXLCell cell;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);

            cell = ReaderBase.GetEntryCell(worksheet, "глубиной заложения");
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.DeviationsPVP.RiverbedDeviations.NGZRiverbed = value;
            }

            offsetColumn = 2;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.DeviationsPVP.FloodplainDeviations.NGZFloodplain = value;
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Длина оголения");
            offsetColumn = 1;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.DeviationsPVP.RiverbedDeviations.DenudationRiverbed = value;
            }

            offsetColumn = 2;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.DeviationsPVP.FloodplainDeviations.DenudationFloodplain = value;
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Длина провиса");
            offsetColumn = 1;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.DeviationsPVP.RiverbedDeviations.SagRiverbed = value;
            }

            offsetColumn = 2;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.DeviationsPVP.FloodplainDeviations.SagFloodplain = value;
            }
        }
        private void ReadSheet_15_1(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int ngzLengthColumn = 3;
            int minThicknessColumn = 6;

            IXLRange table = ReadPVP_GetTableRange(workbook, sheetName);
            IXLTableRows rows = ReadPVP_GetRows(table, "русло");
            IXLCell[] thicknessCells = ReadPVP_GetCells(rows, minThicknessColumn);
            IXLCell[] ngzCells = ReadPVP_GetCells(rows, ngzLengthColumn);

            IEnumerable<float> ngzValues = ReaderBase.CleanCellAndGetValue<float>(ngzCells);
            IEnumerable<float> thicknessValues = ReaderBase.CleanCellAndGetValue<float>(thicknessCells);

            if (ngzValues.Any())
            {
                float totalNgz = ngzValues.Sum();
                watercourse.DeviationsRivebed.LengthNGZ = MathF.Round(totalNgz,2);
            }
            if (thicknessValues.Any())
            {
                float minThikness = thicknessValues.Min();
                watercourse.DeviationsRivebed.MinThicknessProtectingLayer = minThikness;
            }           
        }
        private void ReadSheet_15_2(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int ogLengthColumn = 3;
            int minDepthColumn = 6;

            IXLRange table = ReadPVP_GetTableRange(workbook, sheetName);
            IXLTableRows rows = ReadPVP_GetRows(table, "русло");
            IXLCell[] depthCells = ReadPVP_GetCells(rows, minDepthColumn);
            IXLCell[] ogCells = ReadPVP_GetCells(rows, ogLengthColumn);

            IEnumerable<float> ogValues = ReaderBase.CleanCellAndGetValue<float>(ogCells);
            IEnumerable<float> depthValues = ReaderBase.CleanCellAndGetValue<float>(depthCells);

            if (ogValues.Any())
            {
                float totalOg = ogValues.Sum();
                watercourse.DeviationsRivebed.LengthDenudation = MathF.Round(totalOg, 2);
            }
            if (depthValues.Any())
            {
                float maxDepth = depthValues.Max();
                watercourse.DeviationsRivebed.MaxDepthDenudation = maxDepth;
            }
        }
        private void ReadSheet_15_3(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int sagLengthColumn = 3;

            IXLRange table = ReadPVP_GetTableRange(workbook, sheetName);
            IXLTableRows rows = ReadPVP_GetRows(table, "русло");
            IXLCell[] sagCells = ReadPVP_GetCells(rows, sagLengthColumn);

            IEnumerable<float> sagValues = ReaderBase.CleanCellAndGetValue<float>(sagCells);

            if (sagValues.Any())
            {
                float totalSag = sagValues.Sum();
                float maxSag = sagValues.Max();
                watercourse.DeviationsRivebed.LengthSag = MathF.Round(totalSag, 2);
                watercourse.DeviationsRivebed.MaxLengthSinglePart = maxSag;
            }
        }
        private IXLRange ReadPVP_GetTableRange(XLWorkbook workbook, string sheetName)
        {
            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);
            IXLCell firstUsedCell = worksheet.FirstCellUsed();
            IXLCell lastUsedCell = worksheet.LastCellUsed();
            IXLRange table = worksheet.Range(firstUsedCell, lastUsedCell);

            return table;
        }
        private IXLTableRows ReadPVP_GetRows(IXLRange table, string param)
        {
            IXLTableRows rows = table.AsTable().DataRange
                .Rows(row => row.Cell(2).GetString() == param);

            return rows;
        }
        private IXLCell[] ReadPVP_GetCells(IXLTableRows rows, int column)
        {
            IXLCell[] sagCells = rows.Select(row => row.Cell(column))
                .Where(cell => !cell.IsEmpty())
                .ToArray();

            return sagCells;
        }
        private void SetPositionMT(Watercourse watercourse)
        {
            DeviationsRivebed obj = watercourse.DeviationsRivebed;

            PropertyInfo[] props = watercourse.DeviationsRivebed.GetType().GetProperties();

            var allZeros = props.All(prop => (float)prop.GetValue(obj)! == 0);

            if (allZeros)
            {
                watercourse.PositionMT = "ниже";
            }
            else
            {
                watercourse.PositionMT = "выше";
            }
        }
    }
}

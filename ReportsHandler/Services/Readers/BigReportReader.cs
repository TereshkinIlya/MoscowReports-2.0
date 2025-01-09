using ClosedXML.Excel;
using Core.Abstracts;
using Core.Data.Watercourse;
using Core.Interaces;
using Microsoft.Extensions.DependencyInjection;
using ReportsHandler.Abstracts;
using ReportsHandler.Interfaces;

namespace ReportsHandler.Services.Readers
{
    internal sealed class BigReportReader : IExcelReader<XLWorkbook>
    {
        private Repository<Queue<Content>, Content> Repository { get; set; }
        private IData DataFactory { get; set; }
        private ReaderBase ReaderBase { get; set; }
        public BigReportReader([FromKeyedServices("bigsRepo")] Repository<Queue<Content>, Content> repository,
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

            if (sheet is not null)
            {
                ReadAsBig(workbook);
            }
        }
        private void ReadAsBig(XLWorkbook workbook)
        {
            var watercourse = (Watercourse)DataFactory.GetInstance<Watercourse>();
            ReadSheet_1(workbook, watercourse, "Общие сведения");
            ReadSheet_4(workbook, watercourse, "хар-ка");
            ReadSheet_9(workbook, watercourse, "ФГХ");
            ReadSheet_12(workbook, watercourse, "1% и 10%");
            ReadSheet_15(workbook, watercourse, "РП и ППРР");
            ReadSheet_16(workbook, watercourse, "Анализ РП и ПВП");

            Repository.Put(watercourse);
        }
        private void ReadSheet_1(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 2;
            int offsetColumn = 0;
            DateTime date;

            IXLWorksheet worksheet= ReaderBase.GetWorksheet(workbook, sheetName);
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
        private void ReadSheet_9(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 0;
            int offsetColumn = 1;
            bool success;
            string coordinate;
            IXLCell cell;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);
            cell = ReaderBase.GetEntryCell(worksheet, "Широта, левый берег");

            success = ReaderBase.TryGetRequiredCellValue<string>(cell, offsetRow, offsetColumn, out coordinate);
            if (success)
            {
                watercourse.Coordinates.LeftCoast.Latitude = ReaderBase.GetGeoCoordinate(cell, coordinate);
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Долгота, левый берег");

            success = ReaderBase.TryGetRequiredCellValue<string>(cell, offsetRow, offsetColumn, out coordinate);
            if (success)
            {
                watercourse.Coordinates.LeftCoast.Longitude = ReaderBase.GetGeoCoordinate(cell, coordinate);
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Широта, правый берег");

            success = ReaderBase.TryGetRequiredCellValue<string>(cell, offsetRow, offsetColumn, out coordinate);
            if (success)
            {
                watercourse.Coordinates.RightCoast.Latitude = ReaderBase.GetGeoCoordinate(cell, coordinate);
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Долгота, правый берег");

            success = ReaderBase.TryGetRequiredCellValue<string>(cell, offsetRow, offsetColumn, out coordinate);
            if (success)
            {
                watercourse.Coordinates.RightCoast.Longitude = ReaderBase.GetGeoCoordinate(cell, coordinate);
            }
        }
        private void ReadSheet_12(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 0;
            int offsetColumn = 1;
            bool success;
            float value;
            IXLCell cell;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);
            cell = ReaderBase.GetEntryCell(worksheet, "Расход воды");

            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.WaterFlowRate.MaxValues.OnePercentWaterLevel = value;
            }

            offsetColumn = 3;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.WaterFlowRate.MaxValues.TenPercentWaterLevel = value;
            }

            offsetColumn = 4;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.WaterFlowRate.AverageWaterLevel = value;
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Максимальная скорость течения");

            offsetColumn = 1;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.MaxSpeedWaterFlow.MaxValues.OnePercentWaterLevel = value;
            }

            offsetColumn = 3;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.MaxSpeedWaterFlow.MaxValues.TenPercentWaterLevel = value;
            }

            offsetColumn = 4;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.MaxSpeedWaterFlow.AverageWaterLevel = value;
            }
        }
        private void ReadSheet_15(XLWorkbook workbook, Watercourse watercourse, string sheetName)
        {
            int offsetRow = 0;
            int offsetColumn = 1;
            string position;

            IXLWorksheet worksheet = ReaderBase.GetWorksheet(workbook, sheetName);
            IXLCell cell = ReaderBase.GetEntryCell(worksheet, "Залегание МТ");
            
            bool success = ReaderBase.TryGetRequiredCellValue<string>(cell, offsetRow, offsetColumn, out position);
            if (success)
            {
                watercourse.PositionMT = position;
            }
        }
        private void ReadSheet_16(XLWorkbook workbook, Watercourse watercourse, string sheetName)
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

            cell = ReaderBase.GetEntryCell(worksheet, "Скорость");
            offsetColumn = 1;
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.RivebedProcesses.SpeedOffsetRiverbed = value;
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Амплитуда отметок");
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.RivebedProcesses.AmplitudeRiverbed = value;
            }

            cell = ReaderBase.GetEntryCell(worksheet, "Высота подвижных");
            success = ReaderBase.TryGetRequiredCellValue<float>(cell, offsetRow, offsetColumn, out value);
            if (success)
            {
                watercourse.RivebedProcesses.HeightMicroforms = value;
            }

            float speed = watercourse.RivebedProcesses.SpeedOffsetRiverbed;
            float amplitude = watercourse.RivebedProcesses.AmplitudeRiverbed;
            float height = watercourse.RivebedProcesses.HeightMicroforms;

            if (speed > 2 || amplitude > 1 || height > 1.5)
            {
                watercourse.RivebedProcesses.Character = "интенсивный";
            }
            else if (speed < 0.5 && amplitude <= 1 && height < 0.5)
            {
                watercourse.RivebedProcesses.Character = "стабильный";
            }        
            else
            {
                watercourse.RivebedProcesses.Character = "умеренный";
            }
        }
    }
}

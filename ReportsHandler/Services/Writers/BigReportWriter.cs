using ClosedXML.Excel;
using Core.Abstracts;
using Core.Data.Watercourse;
using Core.Interaces;
using Core.Tables.HeaderColumns;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.DependencyInjection;
using ReportsHandler.Abstracts;
using ReportsHandler.Interfaces;

namespace ReportsHandler.Services.Writers
{
    internal sealed class BigReportWriter : IExcelWriter<XLWorkbook>
    {
        private Repository<Queue<Content>, Content> Repository { get; set; }
        private IHeader TableHeader { get; set; }
        private ReaderBase ReaderBase { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public BigReportWriter([FromKeyedServices("bigsRepo")] Repository<Queue<Content>, Content> repository, 
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
            
            if( sheet is null)
            {
                return;
            }
            else
            {
                WriteAsBigs(table);
            }
        }
        private void WriteAsBigs(XLWorkbook table)
        {
            int counter = 0;
            int total = Repository.Count();
            while (!Repository.IsEmpty())
            {
                Progress.Report([total, ++counter, "Запись больших водотоков в Московскую таблицу", false]);

                var watercourse = (Watercourse)Repository.Get();
                WriteBigReportsInTable(table, watercourse);
            }
        }
        private void WriteBigReportsInTable(XLWorkbook table, Watercourse watercourse)
        {
            int row;
            int column;

            BigStreamsColumns columns = (BigStreamsColumns)TableHeader.GetHeader<BigStreamsColumns>();
            IXLWorksheet worksheet = ReaderBase.GetWorksheet(table, "ППМН ТН");
            worksheet.SetAutoFilter(false);
            IXLRange ids = ReaderBase.GetIdsRange(worksheet, "ID");
            IXLCell? idCell = ids.Search(watercourse.Id.ToString()).FirstOrDefault();

            if (idCell is null) return;

            row = idCell.WorksheetRow().RowNumber();

            column = columns["DateInspection"];
            worksheet.Cell(row, column).Value = watercourse.DateInspection;

            column = columns["TypeOfSurvey"];
            worksheet.Cell(row, column).Value = watercourse.TypeOfSurvey;

            column = columns["PositionMT"];
            worksheet.Cell(row, column).Value = watercourse.PositionMT;

            column = columns["DeviationsPVP.RiverbedDeviations.NGZRiverbed"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsPVP.RiverbedDeviations.NGZRiverbed;

            column = columns["DeviationsPVP.RiverbedDeviations.DenudationRiverbed"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsPVP.RiverbedDeviations.DenudationRiverbed;

            column = columns["DeviationsPVP.RiverbedDeviations.SagRiverbed"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsPVP.RiverbedDeviations.SagRiverbed;

            column = columns["DeviationsPVP.FloodplainDeviations.NGZFloodplain"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsPVP.FloodplainDeviations.NGZFloodplain;

            column = columns["DeviationsPVP.FloodplainDeviations.DenudationFloodplain"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsPVP.FloodplainDeviations.DenudationFloodplain;

            column = columns["DeviationsPVP.FloodplainDeviations.SagFloodplain"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsPVP.FloodplainDeviations.SagFloodplain;

            column = columns["RivebedProcesses.SpeedOffsetRiverbed"];
            worksheet.Cell(row, column).Value = watercourse.RivebedProcesses.SpeedOffsetRiverbed;

            column = columns["RivebedProcesses.AmplitudeRiverbed"];
            worksheet.Cell(row, column).Value = watercourse.RivebedProcesses.AmplitudeRiverbed;

            column = columns["RivebedProcesses.HeightMicroforms"];
            worksheet.Cell(row, column).Value = watercourse.RivebedProcesses.HeightMicroforms;

            column = columns["RivebedProcesses.Character"];
            worksheet.Cell(row, column).Value = watercourse.RivebedProcesses.Character;

            column = columns["Coordinates.LeftCoast.Latitude"];
            worksheet.Cell(row, column).Value = watercourse.Coordinates.LeftCoast.Latitude;

            column = columns["Coordinates.LeftCoast.Longitude"];
            worksheet.Cell(row, column).Value = watercourse.Coordinates.LeftCoast.Longitude;

            column = columns["Coordinates.RightCoast.Latitude"];
            worksheet.Cell(row, column).Value = watercourse.Coordinates.RightCoast.Latitude;

            column = columns["Coordinates.RightCoast.Longitude"];
            worksheet.Cell(row, column).Value = watercourse.Coordinates.RightCoast.Longitude;

            column = columns["WaterFlowRate.MaxValues.OnePercentWaterLevel"];
            worksheet.Cell(row, column).Value = watercourse.WaterFlowRate.MaxValues.OnePercentWaterLevel;

            column = columns["WaterFlowRate.MaxValues.TenPercentWaterLevel"];
            worksheet.Cell(row, column).Value = watercourse.WaterFlowRate.MaxValues.TenPercentWaterLevel;

            column = columns["WaterFlowRate.AverageWaterLevel"];
            worksheet.Cell(row, column).Value = watercourse.WaterFlowRate.AverageWaterLevel;

            column = columns["MaxSpeedWaterFlow.MaxValues.OnePercentWaterLevel"];
            worksheet.Cell(row, column).Value = watercourse.MaxSpeedWaterFlow.MaxValues.OnePercentWaterLevel;

            column = columns["MaxSpeedWaterFlow.MaxValues.TenPercentWaterLevel"];
            worksheet.Cell(row, column).Value = watercourse.MaxSpeedWaterFlow.MaxValues.TenPercentWaterLevel;

            column = columns["MaxSpeedWaterFlow.AverageWaterLevel"];
            worksheet.Cell(row, column).Value = watercourse.MaxSpeedWaterFlow.AverageWaterLevel;            
        }
    }
}

using ClosedXML.Excel;
using Core.Abstracts;
using Core.Data.Watercourse;
using Core.Interaces;
using Core.Tables.HeaderColumns;
using Microsoft.Extensions.DependencyInjection;
using ReportsHandler.Abstracts;
using ReportsHandler.Interfaces;

namespace ReportsHandler.Services.Writers
{
    internal sealed class SmallReportWriter : IExcelWriter<XLWorkbook>
    {
        private Repository<Queue<Content>, Content> Repository { get; set; }
        private IHeader TableHeader { get; set; }
        private ReaderBase ReaderBase { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public SmallReportWriter([FromKeyedServices("smallsRepo")] Repository<Queue<Content>, Content> repository,
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
                sheet.Name.Contains("МВ ТН", StringComparison.CurrentCultureIgnoreCase));

            if (sheet is null)
            {
                return;
            }
            else
            {
                WriteAsSmalls(table);
            }
        }
        private void WriteAsSmalls(XLWorkbook table)
        {
            int counter = 0;
            int total = Repository.Count();
            while (!Repository.IsEmpty())
            {
                Progress.Report([total, ++counter, "Запись малых водотоков в Московскую таблицу", false]);

                var watercourse = (Watercourse)Repository.Get();
                WriteSmallReportsInTable(table, watercourse);
            }
        }
        private void WriteSmallReportsInTable(XLWorkbook table, Watercourse watercourse)
        {
            int row;
            int column;

            SmallStreamsColumns columns = (SmallStreamsColumns)TableHeader.GetHeader<SmallStreamsColumns>();
            IXLWorksheet worksheet = ReaderBase.GetWorksheet(table, "МВ ТН");
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

            column = columns["RepairInfo"];
            worksheet.Cell(row, column).Value = watercourse.RepairInfo;

            column = columns["Latitude"];
            worksheet.Cell(row, column).Value = watercourse.Latitude;

            column = columns["Longitude"];
            worksheet.Cell(row, column).Value = watercourse.Longitude;

            column = columns["DeviationsRivebed.LengthNGZ"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsRivebed.LengthNGZ;

            column = columns["DeviationsRivebed.MinThicknessProtectingLayer"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsRivebed.MinThicknessProtectingLayer;

            column = columns["DeviationsRivebed.LengthDenudation"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsRivebed.LengthDenudation;

            column = columns["DeviationsRivebed.MaxDepthDenudation"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsRivebed.MaxDepthDenudation;

            column = columns["DeviationsRivebed.LengthSag"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsRivebed.LengthSag;

            column = columns["DeviationsRivebed.MaxLengthSinglePart"];
            worksheet.Cell(row, column).Value = watercourse.DeviationsRivebed.MaxLengthSinglePart;
        }
    }
}

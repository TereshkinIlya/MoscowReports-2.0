using Core.Abstracts;
using Core.Attributes;
using Core.Data.PiktsInfo.Parts;

namespace Core.Data.PiktsInfo
{
    public class Pikts : Content
    {
        [HeaderCell("ID")]
        public int Id { get; set; }
        [HeaderCell("Состояние (в работе/отключена)")]
        public string State { get; set; } = string.Empty;
        [HeaderCell("Марка стали")]
        public string? SteelGrade { get; set; }
        [HeaderCell("Дата  последнего ВТД")]
        public string? DateLastVTD { get; set; }
        [HeaderCell("Наличие дефектов")]
        public DefectsPresence Defects { get; set; }
        [HeaderCell("Срок безопасной эксплуатации по отчету ОТС")]
        public string? SafePeriod { get; set; }
        [HeaderCell("Дата выдачи отчета по ОТС")]
        public string? DateReport { get; set; }
        [HeaderCell("Номер ОТС")]
        public string? OTSNumber { get; set; }
        [HeaderCell("Cрок безопасной эксплуатации по параметрам")]
        public PeriodSafeOperation PeriodSafeOperation { get; set; }
        [HeaderCell("Организация, выдавшая заключение о сроке безопасной эксплуатации")]
        public string? Organization { get; set; }
        [HeaderCell("Наличие мероприятий по приведению")]
        public string? Measures { get; set; }
        public Pikts()
        {
            PeriodSafeOperation = new();
            Defects = new();
        }
    }
}

using Core.Attributes;

namespace Core.Data.PiktsInfo.Parts
{
    public class PeriodSafeOperation
    {
        [HeaderCell("Проведение ВТД")]
        public string DateVTD { get; set; } = string.Empty;
        [HeaderCell("Неустраненный дефект")]
        public string DateDefect { get; set; } = string.Empty;
        [HeaderCell("Диагностика спиралешовных труб")]
        public string DateCDS { get; set; } = string.Empty;
        [HeaderCell("Диагностирование объектов не подлежащих ВТД")]
        public string DateJumpers { get; set; } = string.Empty;
        [HeaderCell("Диагностика объектов с ограниченными возможностями ВТД")]
        public string DateLimited { get; set; } = string.Empty;
        [HeaderCell("Диагностика ВРК")]
        public string DateVRK { get; set; } = string.Empty;
        [HeaderCell("Диагностик запорной арматуры")]
        public string DateUZA { get; set; } = string.Empty;
        [HeaderCell("Диагностика приварных")]
        public string DateWeldedElement { get; set; } = string.Empty;
        [HeaderCell("Диагностика соед. деталей")]
        public string DateConnectedDetails { get; set; } = string.Empty;
        [HeaderCell("Диагностика КПП СОД")]
        public string DateKPSOD { get; set; } = string.Empty;
        [HeaderCell("Диагностика дренажных емкостей")]
        public string DateDrainageContainers { get; set; } = string.Empty;
        [HeaderCell("Обследование ПВП и русловых процессов")]
        public string DatePVP { get; set; } = string.Empty;
        [HeaderCell("Обследование корр. сост.")]
        public string DateCorrosion { get; set; } = string.Empty;
    }
}

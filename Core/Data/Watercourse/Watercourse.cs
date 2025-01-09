using Core.Abstracts;
using Core.Attributes;
using Core.Data.Watercourse.Parts;

namespace Core.Data.Watercourse
{
    public class Watercourse : Content
    {
        [HeaderCell("ID")]
        public int Id { get; set; }
        [HeaderCell("Дата последнего обследования")]
        public DateTime DateInspection { get; set; }
        [HeaderCell("Вид обследования")]
        public string TypeOfSurvey { get; set; } = "обследование";
        [HeaderCell("Положение МТ по отношению к ППРР (выше/ниже)")]
        public string PositionMT { get; set; } = string.Empty;
        [HeaderCell("Разделение на пойму и русло")]
        public DeviationsPVP DeviationsPVP { get; set; }
        [HeaderCell("информация о ремонтах выявленных отклонений")]
        public string RepairInfo { get; set; } = string.Empty; // для малых водотоков
        [HeaderCell("Русловые процессы")]
        public RivebedProcesses RivebedProcesses { get; set; }
        [HeaderCell("Координаты")]
        public Coordinates Coordinates { get; set; }
        [HeaderCell("Широта на момент обследования")]
        public double Latitude { get; set; } // для малых водотоков
        [HeaderCell("Долгота на момент обследования")]
        public double Longitude { get; set; } // для малых водотоков
        [HeaderCell("Наличие отклонений ПВП в русле")]
        public DeviationsRivebed DeviationsRivebed { get; set; } // для малых водотоков
        [HeaderCell("Расход воды")]
        public WaterMovement WaterFlowRate { get; set; }
        [HeaderCell("Максимальная скорость течения воды в русле")]
        public WaterMovement MaxSpeedWaterFlow { get; set; }
        public Watercourse()
        {
            DeviationsPVP = new();
            RivebedProcesses = new();
            Coordinates = new();
            DeviationsRivebed = new();
            WaterFlowRate = new();
            MaxSpeedWaterFlow = new();
        }
    }
}

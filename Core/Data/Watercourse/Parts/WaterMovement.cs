using Core.Attributes;

namespace Core.Data.Watercourse.Parts
{
    public class WaterMovement
    {
        [HeaderCell("При максимальных значениях заданной обеспеченности")]
        public MaxValues MaxValues { get;set; }
        [HeaderCell("При среднемеженном уровне воды")]
        public float AverageWaterLevel { get; set; }
        public WaterMovement() => MaxValues = new();
    }
    public class MaxValues
    {
        [HeaderCell("1")]
        public float OnePercentWaterLevel { get; set; }
        [HeaderCell("10")]
        public float TenPercentWaterLevel { get; set; }
    }
}

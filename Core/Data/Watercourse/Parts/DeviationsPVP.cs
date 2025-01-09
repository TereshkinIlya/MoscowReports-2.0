using Core.Attributes;

namespace Core.Data.Watercourse.Parts
{
    public class DeviationsPVP
    {
        [HeaderCell("Наличие отклонений ПВП в русле")]
        public RiverbedDeviations RiverbedDeviations {  get; set; }
        [HeaderCell("Наличие отклонений ПВП в пойме")]
        public FloodplainDeviations FloodplainDeviations { get;set; }
        public DeviationsPVP()
        {
            RiverbedDeviations = new();
            FloodplainDeviations = new();
        }
    }

    public class RiverbedDeviations
    {
        [HeaderCell("недостаточное заглубление, м")]
        public float NGZRiverbed { get; set; }
        [HeaderCell("в т.ч. оголения, м")]
        public float DenudationRiverbed { get; set; }
        [HeaderCell("в т.ч. провис, м")]
        public float SagRiverbed { get; set; }
    }
    public class FloodplainDeviations
    {
        [HeaderCell("недостаточное заглубление, м")]
        public float NGZFloodplain { get; set; }
        [HeaderCell("в т.ч. оголения, м")]
        public float DenudationFloodplain { get; set; }
        [HeaderCell("в т.ч. провис, м")]
        public float SagFloodplain { get; set; }
    }
}

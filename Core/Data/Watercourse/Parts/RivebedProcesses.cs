using Core.Attributes;

namespace Core.Data.Watercourse.Parts
{
    public class RivebedProcesses
    {
        [HeaderCell("Ср")]
        public float SpeedOffsetRiverbed { get; set; }
        [HeaderCell("Ам")]
        public float AmplitudeRiverbed { get; set; }
        [HeaderCell("Вм")]
        public float HeightMicroforms { get; set; }
        [HeaderCell("характер")]
        public string Character { get; set; } = string.Empty;
    }
}

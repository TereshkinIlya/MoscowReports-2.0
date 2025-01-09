using Core.Attributes;

namespace Core.Data.Watercourse.Parts
{
    public class DeviationsRivebed
    {
        [HeaderCell("недостаточное заглубление")]
        public float LengthNGZ { get; set; }
        [HeaderCell("min толщина защитного слоя")]
        public float MinThicknessProtectingLayer { get; set; }
        [HeaderCell("в т.ч. оголения")]
        public float LengthDenudation { get; set; }
        [HeaderCell("max глубина оголения")]
        public float MaxDepthDenudation { get; set; }
        [HeaderCell("в т.ч. провис")]
        public float LengthSag { get; set; }
        [HeaderCell("длина единичного участка с max протяженностью")]
        public float MaxLengthSinglePart { get; set; }
    }
}

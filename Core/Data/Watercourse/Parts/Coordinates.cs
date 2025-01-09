using Core.Attributes;

namespace Core.Data.Watercourse.Parts
{
    public class Coordinates
    {
        [HeaderCell("Левый берег")]
        public LeftCoast LeftCoast { get; set; }
        [HeaderCell("Правый берег")]
        public RightCoast RightCoast { get; set; }
        public Coordinates()
        {
            LeftCoast = new();
            RightCoast = new();
        }
    }
    public class LeftCoast
    {
        [HeaderCell("Широта, на момент обследования")]
        public double Latitude { get; set; }
        [HeaderCell("Долгота, на момент обследования")]
        public double Longitude { get; set; }
    }
    public class RightCoast
    {
        [HeaderCell("Широта, на момент обследования")]
        public double Latitude { get; set; }
        [HeaderCell("Долгота, на момент обследования")]
        public double Longitude { get; set; }
    }
}

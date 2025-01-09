using Core.Attributes;

namespace Core.Data.PiktsInfo.Parts
{
    public class DefectsPresence
    {
        [HeaderCell("CurrentYear")]
        public Deadline CurrentYear { get; set; }
        [HeaderCell("CurrentYearPlusOne")]
        public Deadline CurrentYearPlusOne { get; set;}
        [HeaderCell("CurrentYearPlusTwo")]
        public Deadline CurrentYearPlusTwo { get; set; }
        public DefectsPresence()
        {
            CurrentYear = new();
            CurrentYearPlusOne = new();
            CurrentYearPlusTwo = new();
        }
    }
    public class Deadline
    {
        [HeaderCell("пойма")]
        public FloodplaneDefects FloodplaneDefects { get; set; }
        [HeaderCell("русло")]
        public RiverbedDefects RiverbedDefects { get; set; }
        public Deadline()
        {
            FloodplaneDefects = new();
            RiverbedDefects = new();
        }
    }
    public class FloodplaneDefects
    {
        public int Quantity { get; set; }
    }
    public class RiverbedDefects
    {
        public int Quantity { get; set; }
    }
}

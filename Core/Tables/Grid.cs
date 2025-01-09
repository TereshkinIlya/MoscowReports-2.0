using Core.Abstracts;

namespace Core.Tables
{
    public class Grid : Cell
    {
        public Cell[] Cells { get; set; } = [];
        public Grid() : base() { }
        public Grid(string value) : base(value) { }
    }
}

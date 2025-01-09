namespace Core.Abstracts
{
    public abstract class Cell
    {
        public string Value { get; } = string.Empty;
        public int ColumnNumber {  get; set; }
        public Cell() { }
        public Cell(string value) => Value = value;
    }
}

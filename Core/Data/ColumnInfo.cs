using Core.Abstracts;

namespace Core.Data
{
    public sealed class ColumnInfo : Content
    {
        public string PropertyName { get; set; } = string.Empty;
        public string AttributeValue { get; set; } = string.Empty;
        public int ColumnNumber { get; set; }
    }
}

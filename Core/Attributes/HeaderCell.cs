namespace Core.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class HeaderCell : Attribute
    {
        public string Value { get; set; } = string.Empty;
        private HeaderCell() { }
        public HeaderCell(string value)
        {
            Value = value switch
            {
                "CurrentYear" => $"с предельным сроком эксплуатации {DateTime.Now.Year}",
                "CurrentYearPlusOne" => $"с предельным сроком эксплуатации {DateTime.Now.Year + 1}",
                "CurrentYearPlusTwo" => $"с предельным сроком эксплуатации {DateTime.Now.Year + 2}",
                _ => value
            };
        } 
    }
}

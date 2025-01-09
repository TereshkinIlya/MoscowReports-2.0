namespace Dto
{
    public class MainFormDto
    {
        private static MainFormDto? _instance;
        public string MoscowTablePath { get; set; } = string.Empty;
        public string PiktsTablePath { get; set; } = string.Empty;
        public string ReportsA_Path { get; set; } = string.Empty;
        public bool OnlyReportsA { get; set; }
        public DateTime LimitDate {  get; set; }
        private MainFormDto() { }
        public static MainFormDto GetInstance()
        {
            _instance ??= new MainFormDto();
            return _instance;
        }
    }
}

namespace ExcelConverter.Interfaces
{
    public interface IExcelConverter<TDirectory>
    {
        void ConvertExcelFilesToXLSX(TDirectory filesFolder);
    }
}

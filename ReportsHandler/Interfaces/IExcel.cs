namespace ReportsHandler.Interfaces
{
    internal interface IExcel<TExcel>  where TExcel : class
    {
        TExcel Open(string filePath);
        void Read(TExcel workbook);
        void Write(TExcel table);
    }
    internal interface IExcelReader<TSource> where TSource : class 
    {
        void Read(TSource source);
    }
    internal interface IExcelWriter<TTable> where TTable : class
    {
        void Write(TTable table);
    }
}

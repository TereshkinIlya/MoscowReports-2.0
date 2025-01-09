namespace ReportsHandler.Interfaces
{
    public interface IReportHandler<TDirectory>
    {
        void HandleReports(TDirectory filesFolder);
    }
}

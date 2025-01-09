namespace FileManager.Interfaces
{
    public interface IFileManager<TDirectory>
    {
        TDirectory GetFolderWithReports();
    }
}

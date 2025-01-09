namespace FileManager.Interfaces
{
    internal interface ISearcher<TPath>
    {
        IEnumerable<TPath> GetFilePaths();
    }
}

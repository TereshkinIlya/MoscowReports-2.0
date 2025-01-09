using FileManager.Interfaces;

namespace FileManager
{
    internal class FileManager : IFileManager<string>
    {
        private ISearcher<string> Searcher { get; set; }
        private ICopier<string, string> Copier { get; set; }
        public FileManager(ISearcher<string> searcher, ICopier<string, string> copier)
        {
            Searcher = searcher;
            Copier = copier;
        }
        public string GetFolderWithReports()
        {
            IEnumerable<string> paths = Searcher.GetFilePaths();
            string filesFolder = Copier.CopyFiles(paths);

            return filesFolder;
        }
    }
}

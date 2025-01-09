namespace FileManager.Interfaces
{
    internal interface ICopier<TFile, TDirectory>
    {
        TDirectory CopyFiles(IEnumerable<TFile> filePaths);
    }
}

using FileManager.Interfaces;
using Logging;
using System.Diagnostics.Metrics;
using System.IO;

namespace FileManager
{
    internal sealed class ReportA_Copier : ICopier<string, string>
    {
        private string DestinationPath {get; set;}
        private IProgress<object[]> Progress { get; set; }
        public ReportA_Copier(IProgress<object[]> progress)
        {
            Progress = progress;
            DestinationPath = CreateSaveFolder();
        }
        public string CopyFiles(IEnumerable<string> filePaths)
        {
            int counter = 0;
            foreach (string file in filePaths)
            {
                Progress.Report([filePaths.Count(), ++counter, "Копирование файлов Приложений А", false]);
                CopyFile(file);
            }
            return DestinationPath;
        }
        private void CopyFile(string file)
        {
            try
            {
                if (!Path.GetFileName(file).StartsWith("~$"))
                {
                    File.Copy(file, Path.Combine(DestinationPath, Path.GetFileName(file)), true);
                }
            }
            catch (FileNotFoundException ex)
            {
                Logger.Log(ex);
            }
            catch (SystemException ex)
            {
                Logger.Log(ex);
            }
        }
        private static string CreateSaveFolder()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string savingPath = Path.Combine(desktopPath, "Приложения А для московской таблицы");

            if (!Directory.Exists(savingPath))
                Directory.CreateDirectory(savingPath);

            return savingPath;
        }
    }
}

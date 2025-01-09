using Dto;
using FileManager.Extensions;
using FileManager.Interfaces;
using Logging;
using System.IO;

namespace FileManager
{
    internal sealed class ReportA_Searcher : ISearcher<string>
    {
        private string Pattern { get; set; }
        private MainFormDto MainForm { get; set; }
        private IProgress<object[]> Progress { get; set; }
        public ReportA_Searcher(IProgress<object[]> progress)
        {
            Progress = progress;
            MainForm = MainFormDto.GetInstance();
            Pattern = "*_А.xls*";
        }
        public IEnumerable<string> GetFilePaths()
        {
            string rootFolder = MainForm.ReportsA_Path;

            IEnumerable<string> yearsOfReleaseReports = GetYearsOfReleaseReports(MainForm.LimitDate);

            IEnumerable<string> paths = GetReportA_Paths(rootFolder, SearchOption.TopDirectoryOnly);

            if (paths.IsNullOrEmptyCollection())
            {
                paths = KeepSearchingReportA_Paths(rootFolder, yearsOfReleaseReports);
                return paths;
            }
            else
            {
                return paths;
            }
        }
        /// <summary>
        /// Получение годов выпуска отчетов по обследованию
        /// </summary>
        /// <param name="lastReportsDate">"Приложения А с : "</param>
        private static IEnumerable<string> GetYearsOfReleaseReports(DateTime lastReportsDate)
        {
            int numberOfYears = (DateTime.Now.Year - lastReportsDate.Year) + 1;
            string[] years = new string[numberOfYears];

            years[0] = lastReportsDate.Year.ToString();
            for (int i = 1; i < years.Length; i++)
            {
                years[i] = (lastReportsDate.Year + i).ToString();
            }
            return years;
        }
        private IEnumerable<string> GetReportA_Paths(string rootFolder, SearchOption option)
        {
            string[] filePaths = [];

            try
            {
                filePaths = Directory.EnumerateFiles(rootFolder, Pattern, option).
                Where(file => File.GetLastWriteTime(file) > MainForm.LimitDate.Date
                &&
                (Path.GetExtension(file).EndsWith("xls") ||
                Path.GetExtension(file).EndsWith("xlsx") ||
                Path.GetExtension(file).EndsWith("xlsm"))
                ).ToArray();
            }
            catch (IOException ex)
            {
                Logger.Log(ex);
            }
            catch (SystemException ex)
            {
                Logger.Log(ex);
            }

            return filePaths;
        }
        private IEnumerable<string> KeepSearchingReportA_Paths(string rootFolder, IEnumerable<string> yearsOfReleaseReports)
        {
            int counter = 0;
            string[] subDirectories;
            List<string> filtredDirs = [];
            List<string> filePaths = [];

            subDirectories = Directory.GetDirectories(rootFolder);
            foreach (string year in yearsOfReleaseReports)
            {
                Progress.Report([100, 0, "Поиск каталогов Приложений А", true]);
                LookForInSubFolders(filtredDirs, subDirectories, year);
            }
            foreach (string directory in filtredDirs)
            {
                Progress.Report([filtredDirs.Count(), ++counter, "Поиск файлов Приложений А", false]);
                filePaths.AddRange(GetReportA_Paths(directory, SearchOption.AllDirectories));
            }

            return filePaths;
        }
        private void LookForInSubFolders(List<string> filtredDirs, string[] subDirectories, string year)
        {
            foreach (string directory in subDirectories)
            {
                string[] folders = Directory.GetDirectories(directory);
                if (folders.Any(dir => dir.Contains(year)))
                {
                    filtredDirs.AddRange(folders.Where(path => path.Contains(year)));
                }
                else
                {
                    subDirectories = Directory.GetDirectories(directory);
                    LookForInSubFolders(filtredDirs, subDirectories, year);
                }
            }
        }
    }
}

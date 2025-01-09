using System.Collections;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Text;

namespace Logging
{
    public static class Logger
    {
        private static string LogFolderPath { get; set; }
        private static string LogFilePath { get; set; }
        static Logger()
        {
            LogFolderPath = Path.Combine(Environment.
                GetFolderPath(Environment.SpecialFolder.Desktop),
                "Приложения А для московской таблицы", "Файлы с ошибками");

            LogFilePath = Path.Combine(LogFolderPath, "!Ошибки.txt");
        }
        public static void Log(string message)
        {
            try
            {
                CreateLogFolder();
                File.AppendAllText(LogFilePath, message + "\n");
            }
            catch (SystemException)
            {
            }
        }
        public static void Log(string message, string filePath)
        {
            try
            {
                CreateLogFolder();
                File.AppendAllText(LogFilePath, message + "\n");
                CopyFile(filePath);
            }
            catch (SystemException)
            {
            }
        }
        public static void Log(string message, Exception exception)
        {
            try
            {
                CreateLogFolder();
                File.AppendAllText(LogFilePath, message + "\n");
                if(exception.Data.Count > 0)
                {
                    StringBuilder data = new StringBuilder();
                    foreach (DictionaryEntry item in exception.Data)
                    {
                        data.Append($"{item.Key}) - {item.Value}" + "\n");
                    }
                    File.AppendAllText(LogFilePath, data.ToString());
                }
            }
            catch (SystemException)
            {
            }
        }
        public static void Log(Exception exception)
        {
            try
            {
                CreateLogFolder();
                File.AppendAllText(LogFilePath, exception.Message + "\n");
                if (exception.Data.Count > 0)
                {
                    foreach (DictionaryEntry item in exception.Data)
                    {
                        File.AppendAllText(LogFilePath, $"{item.Key}) - {item.Value}" + "\n");
                    }
                }
            }
            catch (SystemException)
            {
            }
        }
        public static void OpenLogFile()
        {
            if (File.Exists(LogFilePath))
            {
                try
                {
                    if (new FileInfo(LogFilePath).Length == 0)
                    {
                        return;
                    }

                    ProcessStartInfo startInfo = new ProcessStartInfo(LogFilePath)
                    {
                        UseShellExecute = true
                    };

                    Process.Start(startInfo);
                }
                catch (Exception)
                {
                }
            }
        }
        public static void ClearLogFile()
        {
            if (Directory.Exists(LogFolderPath))
            {
                try
                {
                    File.Create(LogFilePath).Close();
                }
                catch (SystemException)
                {
                }
            }
        }
        private static void CreateLogFolder()
        {
            if (!Directory.Exists(LogFolderPath))
            {
                try
                {
                    Directory.CreateDirectory(LogFolderPath);
                }
                catch (SystemException)
                {
                }
            }
        }
        private static void CopyFile(string pathFile)
        {
            try
            {
                File.Copy(pathFile, Path.Combine(LogFolderPath, Path.GetFileName(pathFile)), true);
            }
            catch (SystemException)
            {
            }
        }
    }
}

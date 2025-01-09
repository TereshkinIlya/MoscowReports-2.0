using ClosedXML.Excel;
using Core.Abstracts;
using Core.Interaces;
using Core.Tables;
using DocumentFormat.OpenXml.Office2019.Drawing.Animation;
using DocumentFormat.OpenXml.Wordprocessing;
using Dto;
using Logging;
using System.Diagnostics;
using TableHeadersHandler.Interfaces;

namespace TableHeadersHandler
{
    internal class TableHeadline : ISpreadsheet
    {
        private MainFormDto MainForm { get; set; }
        private IHeadline<XLWorkbook> HeaderInitializer { get; set; }
        private ITable TableFactory { get; set; }
        IProgress<object[]> Progress { get; set; }
        public TableHeadline(IHeadline<XLWorkbook> headerInitializer, 
                             ITable tableFactory, 
                             IProgress<object[]> progress)
        {
            MainForm = MainFormDto.GetInstance();
            HeaderInitializer = headerInitializer;
            TableFactory = tableFactory;
            Progress = progress;
        }
        public void InitializeHeader()
        {
            string path = string.Empty;
            try
            {
                CheckTargetTables();

                path = MainForm.MoscowTablePath;
                Progress.Report([100, 0, "Инициализация заголовка Московской таблицы", true]);
                OpenMoscowTable(path);

                if(!MainForm.OnlyReportsA)
                {
                    path = MainForm.PiktsTablePath;
                    Progress.Report([100, 0, "Инициализация заголовка ПИКТС таблицы", true]);
                    OpenPiktsTable(path);
                }
            }
            catch (ArgumentException ex)
            {
                Logger.Log(ex.Message);
                throw new UnreachableException($"Невозможно открыть таблицу : " +
                    $"{Path.GetFileName(path)}. См. лог файл");
            }
            catch (IOException)
            {
                throw new UnreachableException($"Закройте таблицу {Path.GetFileName(path)}");
            }
        }
        private void CheckTargetTables()
        {
            IXLWorksheet? sheet;
            string path;

            Progress.Report([100, 0, "Проверка загружаемых таблиц", true]);

            path = MainForm.MoscowTablePath;
            using (XLWorkbook table = new XLWorkbook(path))
            {
                sheet = table.Worksheets.FirstOrDefault(sheet =>
                    sheet.Name.Contains("ППМН ТН", StringComparison.CurrentCultureIgnoreCase));

                if (sheet is null)
                    throw new UnreachableException($"Загружена НЕ Московская таблица !\n\n" +
                        $"Не найден лист \"ППМН ТН\" и/или \"МВ ТН\" \n\n{path}");
            }

            if (MainForm.OnlyReportsA) return;

            path = MainForm.PiktsTablePath;
            using (XLWorkbook table = new XLWorkbook(path))
            {
                sheet = table.Worksheets.FirstOrDefault(sheet =>
                    sheet.Name.Contains("Отчет ПАО", StringComparison.CurrentCultureIgnoreCase));

                if (sheet is null)
                    throw new UnreachableException($"Загружена НЕ ПИКТС таблица !\n\n" +
                        $"Не найден лист \"31. Отчет ПАО\"\n\n{path}");
            }
        }
        private void OpenMoscowTable(string path)
        {
            using XLWorkbook workbook = new XLWorkbook(path);

            Table<Grid> bigsTable = TableFactory.CreateTable<BigStreamsTable>();
            HeaderInitializer.SetColumnNumbers(bigsTable, workbook, "ППМН ТН");

            Table<Grid> smallsTable = TableFactory.CreateTable<SmallStreamsTable>();
            HeaderInitializer.SetColumnNumbers(smallsTable, workbook, "МВ ТН");
        }
        private void OpenPiktsTable(string path)
        {
            using XLWorkbook workbook = new XLWorkbook(path);

            Table<Grid> piktsTable = TableFactory.CreateTable<PiktsTable>();
            HeaderInitializer.SetColumnNumbers(piktsTable, workbook, "31. Отчет ПАО");
        }
    }
}

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ReportsHandler.Interfaces;
using ReportsHandler.Services.Readers;
using ReportsHandler.Services.Writers;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace ReportsHandler
{
    internal sealed class ExcelBehaviour : IExcel<XLWorkbook>
    {
        private IEnumerable<IExcelReader<XLWorkbook>> Readers { get; set; }
        private IEnumerable<IExcelWriter<XLWorkbook>> Writers { get; set; }
        public ExcelBehaviour(IEnumerable<IExcelReader<XLWorkbook>> readers, 
                              IEnumerable<IExcelWriter<XLWorkbook>> writers)
        {
            Readers = readers;
            Writers = writers;
        }
        public XLWorkbook Open(string filePath)
        {
            XLWorkbook workbook = GetWorkbook(filePath);
            return workbook;
        }
        public void Read(XLWorkbook workbook)
        {
            _ = workbook ?? throw new ArgumentNullException(nameof(workbook));

            foreach (IExcelReader<XLWorkbook> reader in Readers)
            {
                reader.Read(workbook);
            }
        }
        public void Write(XLWorkbook table)
        {
            _ = table ?? throw new ArgumentNullException(nameof(table));

            foreach (IExcelWriter<XLWorkbook> writer in Writers)
            {
                writer.Write(table);
            }
        }
        private XLWorkbook GetWorkbook(string path)
        {
            XLWorkbook workbook = new();
            try
            {
                workbook = TryGetWorkbook(path, workbook);
            }
            catch (XmlException)
            {
                workbook = FixXmlMarkup(path, workbook);
            }
            return workbook;
        }
        private XLWorkbook TryGetWorkbook(string path, XLWorkbook workbook)
        {
            try
            {
                workbook = new XLWorkbook(path);
            }
            catch (ArgumentException ex)
            {
                if (ex.Message == "Could not load comments file")
                {
                    throw new XmlException(ex.Message);
                }
                else
                {
                    throw new ArgumentException(ex.Message);
                }
            }
            catch (IOException)
            {
                throw new UnreachableException($"Таблица \"{Path.GetFileName(path)}\" открыта пользователем ! " +
                    $"Закройте таблицу и перезапустите программу ! P.S. Ну вот нафига ?");
            }
            return workbook;
        }
        private XLWorkbook FixXmlMarkup(string path, XLWorkbook workbook)
        {
            MemoryStream memoryStream = new MemoryStream();
            CopyToMemoryStream(path, memoryStream);
            memoryStream = SetStartPosition(memoryStream);
            FindAndReplaceTag(memoryStream);
            memoryStream = SetStartPosition(memoryStream);
            workbook = CreateWorkbook(memoryStream, path);
            return workbook;
        }
        private void CopyToMemoryStream(string path, MemoryStream ms)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                fs.CopyTo(ms);
            }
        }
        private MemoryStream SetStartPosition(MemoryStream ms)
        {
            ms.Seek(0, SeekOrigin.Begin);
            return ms;
        }
        private void FindAndReplaceTag(MemoryStream ms)
        {
            using (SpreadsheetDocument dSpreadsheet = SpreadsheetDocument.Open(ms, true))
            {
                _ = dSpreadsheet.WorkbookPart ?? throw new ArgumentNullException(nameof(dSpreadsheet.WorkbookPart));

                foreach (WorksheetPart wsPars in dSpreadsheet.WorkbookPart.WorksheetParts)
                {
                    foreach (VmlDrawingPart vmlDrawingPart in wsPars.VmlDrawingParts)
                    {
                        string text;
                        using (TextReader tr = new StreamReader(vmlDrawingPart.GetStream(FileMode.Open)))
                        {
                            text = tr.ReadToEnd();
                        }
                        using (TextWriter tw = new StreamWriter(vmlDrawingPart.GetStream()))
                        {
                            tw.Write(text.Replace("<br>", "    "));
                        }
                    }
                }
            }
        }
        private XLWorkbook CreateWorkbook(MemoryStream ms, string path)
        {
            try
            {
                var workbook = new XLWorkbook(ms);
                ms.Dispose();
                return workbook;
            }
            catch (ArgumentException ex)
            {
                ms.Dispose();
                throw new ArgumentException($"Can not fix xml markup." +
                    $"Can not open a workbook : {Path.GetFileName(path)}./n" + ex.Message);
            }
        }
    }
}

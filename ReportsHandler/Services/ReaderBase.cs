using ClosedXML.Excel;
using Logging;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ReportsHandler.Services
{
    internal class ReaderBase
    {
        public IXLWorksheet GetWorksheet(XLWorkbook workbook, string name)
        {
            _ = workbook ?? throw new ArgumentNullException(nameof(workbook));
            try
            {
                IXLWorksheet sheet = workbook.Worksheets.First(sheet =>
                    sheet.Name.Contains(name, StringComparison.CurrentCultureIgnoreCase));
                return sheet;
            }
            catch (InvalidOperationException)
            {
                throw new ArgumentException($"Вкладка с именем \"{name}\" не найдена !");
            }
        }
        public IXLCell GetEntryCell(IXLWorksheet worksheet, string value)
        {
            IXLCell? cell = worksheet.Search(value, CompareOptions.IgnoreCase |
                                                    CompareOptions.IgnoreSymbols).FirstOrDefault();
            if (cell == null)
            {
                throw new ArgumentNullException($" Ячейка со значением \"{value}\" на листе \"{worksheet.Name}\" не найдена !");
            }
            else
            {
                return cell;
            }
        }
        public bool TryGetRequiredCellValue<TValue>(IXLCell cell, int offsetRow, int offsetColumn, out TValue value)
        {
            IXLWorksheet worksheet = cell.Worksheet;
            IXLCell requiredCell = worksheet.Cell(cell.Address.RowNumber + offsetRow, cell.Address.ColumnNumber + offsetColumn);

            bool rezult = requiredCell.TryGetValue(out value);

            if (rezult == false)
            {
                try
                {
                    CleanCellAndGetValue(requiredCell, out value);
                }
                catch (ArgumentNullException)
                {
                    return false;
                }
                catch (SystemException)
                {
                    Logger.Log($"Файл {Path.GetFileName(cell.Worksheet.Workbook.Properties.Title)}. Вкладка {worksheet}. " +
                        $"Ячейка {requiredCell.Address}. Невозможно прочитать значение ячейки .", 
                        Path.GetFullPath(cell.Worksheet.Workbook.Properties.Title!));
                    return false;
                }
            }
            return true;
        }
        public double GetGeoCoordinate(IXLCell cell, string coordinate)
        {
            try
            {
                if (!coordinate.Any(symb => char.IsDigit(symb)) &&
                    !coordinate.Contains('.') &&
                    !coordinate.Contains('.')
                   )
                {
                    throw new ArgumentException();
                }

                coordinate = coordinate.Replace(".", ",");

                if (coordinate.IndexOf(",") == 2 || coordinate.IndexOf(",") == 3)
                {
                    return double.Parse(coordinate);
                }
                else
                {
                    coordinate = coordinate.Trim().Substring(0, coordinate.IndexOf(","));
                    coordinate = Regex.Replace(coordinate, @"\s+", " ");
                    string[] geoValues = coordinate.Split(' ');

                    double rezult =
                        double.Parse(geoValues[0]) +
                        double.Parse(geoValues[1]) / 60 +
                        double.Parse(geoValues[2]) / 3600;

                    return Math.Round(rezult, 6);
                }
            }
            catch (SystemException)
            {
                Logger.Log($"Файл {Path.GetFileName(cell.Worksheet.Workbook.Properties.Title)}. Вкладка {cell.Worksheet}. " +
                    $"Ячейка {cell.Address}. Проверьте значение координаты.", Path.GetFullPath(cell.Worksheet.Workbook.Properties.Title!));
                return default;
            }
        }
        public IEnumerable<TValue> CleanCellAndGetValue<TValue>(IXLCell[] cells) where TValue : struct
        {
            List<TValue> values = [];
            foreach (IXLCell cell in cells)
            {
                try
                {
                    string cellString = cell.Value.ToString();
                    StringBuilder newString = CleanTheString(cellString);
                    newString = newString.Replace(".", ",");

                    TValue value = (TValue)Convert.ChangeType(newString.ToString(), typeof(TValue));

                    values.Add(value);
                }
                catch (SystemException)
                {
                    Logger.Log($"Файл {Path.GetFileName(cell.Worksheet.Workbook.Properties.Title)}. Вкладка {cell.Worksheet}. " +
                        $"Ячейка {cell.Address}. Невозможно прочитать значение ячейки.", 
                        Path.GetFullPath(cell.Worksheet.Workbook.Properties.Title!));
                }
            }
            return values;
        }
        public IXLRange GetIdsRange(IXLWorksheet worksheet, string name)
        {
            IXLCell startCell = GetEntryCell(worksheet, "ID");
            IXLCell endCell = worksheet.Cell(worksheet.LastRowUsed()!.RowNumber(), startCell.WorksheetColumn().ColumnNumber());
            IXLRange range = worksheet.Range(startCell, endCell);

            return range;
        }
        private void CleanCellAndGetValue<TValue>(IXLCell requiredCell, out TValue value)
        {
            string text;

            if (requiredCell.IsEmpty())
            {
                throw new ArgumentNullException();
            }
            else
            {
                text = requiredCell.GetString();
            }

            StringBuilder newString = CleanTheString(text);

            if (typeof(TValue).Name != nameof(DateTime))
            {
                newString.Replace('.', ',');
            }

            value = (TValue)Convert.ChangeType(newString.ToString(), typeof(TValue));
        }
        private StringBuilder CleanTheString(string text)
        {
            StringBuilder newString = new StringBuilder();
            foreach (char symbol in text)
            {
                if (symbol >= '0' && symbol <= '9' || symbol == '.' || symbol == ',')
                {
                    newString.Append(symbol);
                }
            }
            return newString;
        }
    }
}

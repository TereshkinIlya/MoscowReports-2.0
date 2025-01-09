using ClosedXML.Excel;
using Core.Abstracts;
using Core.Tables;
using Logging;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using TableHeadersHandler.Interfaces;

namespace TablesHandler
{
    internal class HeaderInitializer : IHeadline<XLWorkbook>
    {
        private List<Grid> NotFoundCells { get; set; } = [];  
        public void SetColumnNumbers(Table<Grid> table, XLWorkbook workbook, string sheetName)
        {
			try
			{   
                IXLWorksheet worksheet = workbook.Worksheet(sheetName);
                SetColumnNumbersFor(table, worksheet);
                CheckNotFoundCells();
            }
			catch (ArgumentException ex)
			{
                Logger.Log(ex.Message);
                throw new ArgumentException($"Не найдена вкладка \"{sheetName}\"");
            }
            catch (UnreachableException ex)
            {
                string message = $"Не найдены ячейки заголовка таблицы на вкладке \"{sheetName}\" :";
                Logger.Log(message, ex);
                throw new ArgumentException("");
            }
            catch (StackOverflowException ex)
            {
                Logger.Log(ex);
            }
        }
        private void SetColumnNumbersFor(Table<Grid> table, IXLWorksheet worksheet)
        {
            foreach (Grid rootCell in table.Headline.Cells.OfType<Grid>())
            {
                IXLCells titleCells = FindCells(worksheet, rootCell.Value);
                if (!titleCells.Any())
                {
                    NotFoundCells.Add(rootCell);
                    continue;
                }
                SearchInTitles(titleCells, worksheet, rootCell);
            }
            NotFoundCells.RemoveAll(cell => cell.ColumnNumber != 0);
        }
        private void SearchInTitles(IXLCells titleCells, IXLWorksheet worksheet, Grid rootCell)
        {
            foreach (IXLCell xLCell in titleCells)
            {
                rootCell.ColumnNumber = xLCell.Address.ColumnNumber;
                SearchInSubtitles(worksheet, xLCell, rootCell);

                if (AllColumnsFound(rootCell)) break;
            }
        }
        private void SearchInSubtitles(IXLWorksheet worksheet, IXLCell xLCell, Grid rootCell)
        {
            foreach (Grid cell in rootCell.Cells.OfType<Grid>())
            {
                IXLRange xLRange = GetRangeForSearching(worksheet, xLCell);
                IXLCells xLCells = FindCells(xLRange, cell.Value);

                if (!xLCells.Any())
                {
                    NotFoundCells.Add(cell);
                    continue;
                }

                IXLCell nextXLCell = xLCells.First();
                cell.ColumnNumber = nextXLCell.Address.ColumnNumber;

                if (xLCell.IsMerged()) SearchInSubtitles(worksheet, nextXLCell, cell);
            }
        }
        private IXLRange GetRangeForSearching(IXLWorksheet worksheet, IXLCell xLCell)
        {
            int firstRow = xLCell.Address.RowNumber;
            int firstColumn = xLCell.Address.ColumnNumber;
            int lastRow = worksheet!.LastRowUsed()!.RowNumber();
            int lastColumn;

            if (xLCell.IsMerged())
                lastColumn = xLCell.MergedRange().LastColumn().ColumnNumber();
            else
                lastColumn = xLCell.Address.ColumnNumber;

            return worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        }
        private IXLCells FindCells(IXLWorksheet worksheet, string value)
        {
            IXLCells xLCells = worksheet.Search(value,
                    CompareOptions.IgnoreCase |
                    CompareOptions.IgnoreSymbols);
            return xLCells;
        }
        private IXLCells FindCells(IXLRange xLRange, string value)
        {
            IXLCells xLCells = xLRange.Search(value,
                    CompareOptions.IgnoreCase |
                    CompareOptions.IgnoreSymbols);
            return xLCells;
        }
        private bool AllColumnsFound(Grid rootCell)
        {
            if (rootCell.ColumnNumber == 0) return false;

            foreach (Grid subCell in rootCell.Cells.OfType<Grid>())
            {
                if (rootCell.Cells.Any(cell => cell.ColumnNumber == 0))
                    return false;
                else
                    AllColumnsFound(subCell);
            }
            return true;        
        }
        private void CheckNotFoundCells()
        {
            IEnumerable<string> notFoundCells = NotFoundCells.Select(cell => cell.Value);
            
            if(notFoundCells.Any())
            {
                int counter = 1;

                UnreachableException exception = new();
                foreach (string value in notFoundCells)
                {
                    exception.Data.Add(counter++, value);
                }
                throw exception;
            }
        }
    }
}

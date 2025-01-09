using ClosedXML.Excel;
using Core.Abstracts;
using Core.Tables;

namespace TableHeadersHandler.Interfaces
{
    internal interface IHeadline<TExcel> where TExcel : class
    {
        void SetColumnNumbers(Table<Grid> table, TExcel workbook, string sheetName);
    }
}

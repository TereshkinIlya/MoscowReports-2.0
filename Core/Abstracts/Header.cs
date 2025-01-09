using Core.Attributes;
using Core.Data;
using Core.Tables;
using Logging;
using System.Diagnostics;
using System.Reflection;

namespace Core.Abstracts
{
    public abstract class Header
    {
        public abstract IReadOnlyCollection<ColumnInfo> Columns { get; }
        public int this[string propertyName]
        {
            get
            {
                var column = Columns.FirstOrDefault(column => column.PropertyName.Equals(propertyName));
                if (column == null)
                {
                    throw new ArgumentException($"Property {propertyName} is not found!");
                }
                return column.ColumnNumber;
            }
        }
        protected virtual IEnumerable<ColumnInfo> SetHeaderColumns<TData, TTable>(TTable table) 
            where TTable : Table<Grid> 
            where TData : Content
        {
            try
            {
                PropertyInfo[] props = typeof(TData).GetProperties();
                var columns = ReadPropertiesAndGetColumnNumbers(props, table);
                
                return columns;
            }
            catch (ArgumentException ex)
            {
                Logger.Log(ex);
                throw new UnreachableException("Ошибка проверки заголовков. См. лог файл");
            }
        }
        private List<ColumnInfo> ReadPropertiesAndGetColumnNumbers(PropertyInfo[] props, Table<Grid> table)
        {
            List<ColumnInfo> columns = [];
            foreach (PropertyInfo property in props)
            {
                HeaderCell headerAttribute = GetAttribute(property);
                
                ColumnInfo column = new();

                column.PropertyName = property.Name;
                column.AttributeValue = headerAttribute.Value;
                column.ColumnNumber = FindColumnNumber(table.Headline.Cells, headerAttribute.Value);

                columns.Add(column);
                
                if (IsUserDefinedProperty(property) && column.ColumnNumber != 0)
                {
                    FindColumnNumberInSubTitles(table, property, columns);
                }
            }
            columns.RemoveAll(column => column.ColumnNumber == 0);

            return columns;
        }
        private HeaderCell GetAttribute(PropertyInfo property)
        {
            HeaderCell headerAttribute = property.GetCustomAttribute<HeaderCell>() ??
                    throw new ArgumentException($"The attribute of type \"HeaderCell\" is not found. " +
                    $"Property name is {property.Name}");
            
            return headerAttribute;
        }
        private int FindColumnNumber(Cell[] cells, string attributeValue)
        {
            Cell? cell = cells.FirstOrDefault(cell =>
                cell.Value.Contains(attributeValue, StringComparison.CurrentCultureIgnoreCase));

            if (cell != null)
            {
                return cell.ColumnNumber;
            }
            else
            {
                return 0;
            }
        }
        private bool IsUserDefinedProperty(PropertyInfo property)
        {
            if (property.PropertyType.IsClass && 
                property.PropertyType != typeof(string) && 
                !property.PropertyType.IsPrimitive)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void FindColumnNumberInSubTitles(Table<Grid> table, PropertyInfo property, List<ColumnInfo> columns)
        {
            ColumnInfo lastColumn = columns.Last();

            Grid subTitle = (Grid)table.Headline.Cells.First(cell =>
                cell.Value.Contains(lastColumn.AttributeValue, StringComparison.CurrentCultureIgnoreCase));

            PropertyInfo[] properties = property.PropertyType.GetProperties();
            KeepSearchingColumnNumber(subTitle.Cells, properties, lastColumn.PropertyName, columns);
        }
        private void KeepSearchingColumnNumber(Cell[] cells, PropertyInfo[] properties, string propertyName, List<ColumnInfo> columns)
        {
            foreach (PropertyInfo property in properties)
            {
                ColumnInfo column = new();
                HeaderCell cellAttribute = GetAttribute(property);
                column.AttributeValue = cellAttribute.Value;
                column.PropertyName = string.Concat(propertyName, ".", property.Name);
                column.ColumnNumber = FindColumnNumber(cells, cellAttribute.Value);

                columns.Add(column);

                if (IsUserDefinedProperty(property))
                {
                    Grid nextCells = (Grid)cells.First(cell =>
                        cell.Value.Contains(cellAttribute.Value, StringComparison.CurrentCultureIgnoreCase));

                    if (nextCells.Cells.Count() != 0)
                    {
                        PropertyInfo[] nextProps = property.PropertyType.GetProperties();
                        KeepSearchingColumnNumber(nextCells.Cells, nextProps, column.PropertyName, columns);
                    }
                }

            }
        }
    }
}

using Core.Abstracts;
using Core.Interaces;
using Core.Tables;

namespace Core.Fabrics
{
    internal sealed class TablesFactory : ITable
    {
        public Table<Grid> CreateTable<TTable>() => typeof(TTable) switch
        {
            Type _ when typeof(TTable) == typeof(BigStreamsTable) => new BigStreamsTable(),
            Type _ when typeof(TTable) == typeof(SmallStreamsTable) => new SmallStreamsTable(),
            Type _ when typeof(TTable) == typeof(PiktsTable) => new PiktsTable(),
            _ => throw new InvalidCastException("Type is not supported")
        };
    }
}

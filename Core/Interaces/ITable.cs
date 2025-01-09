using Core.Abstracts;
using Core.Tables;

namespace Core.Interaces
{
    public interface ITable
    {
        Table<Grid> CreateTable<TTable>();
    }
}

using Core.Abstracts;

namespace Core.Interaces
{
    public interface IData
    {
        Content GetInstance<TData>();
    }
}

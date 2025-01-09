using Core.Tables;

namespace Core.Abstracts
{
    public abstract class Table<THeadline> where THeadline : Grid
    {
        public abstract THeadline Headline { get; }
    }
}

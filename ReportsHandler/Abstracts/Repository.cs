namespace ReportsHandler.Abstracts
{
    internal abstract class Repository<TStorage, TData> where TStorage : class
                                                        where TData : class
    {
        protected abstract TStorage Storage { get; set; }
        public abstract void Put(TData item);
        public abstract TData Get();
        public abstract bool IsEmpty();
        public abstract int Count();
    }
}

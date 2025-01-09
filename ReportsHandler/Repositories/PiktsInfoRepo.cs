using Core.Abstracts;
using ReportsHandler.Abstracts;

namespace ReportsHandler.Repositories
{
    internal class PiktsInfoRepo : Repository<Queue<Content>, Content>
    {
        protected override Queue<Content> Storage { get; set; }
        public PiktsInfoRepo()
        {
            Storage = new Queue<Content>();
        }
        public override Content Get()
        {
            return Storage.Dequeue();
        }

        public override void Put(Content item)
        {
            Storage.Enqueue(item);
        }
        public override bool IsEmpty()
        {
            if (Storage.Count == 0)
                return true;
            else 
                return false;
        }
        public override int Count()
        {
            return Storage.Count;
        }
    }
}

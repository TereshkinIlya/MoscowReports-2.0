using Core.Abstracts;
using Core.Data.PiktsInfo;
using Core.Data.Watercourse;
using Core.Interaces;


namespace Core.Fabrics
{
    internal sealed class ContentFactory : IData
    {
        public Content GetInstance<TContent>() => typeof(TContent) switch
        {
            Type _ when typeof(TContent) == typeof(Pikts) => new Pikts(),
            Type _ when typeof(TContent) == typeof(Watercourse) => new Watercourse(),
            _ => throw new InvalidCastException("Type is not supported")
        };
    }
}

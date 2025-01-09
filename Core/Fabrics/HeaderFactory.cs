using Core.Abstracts;
using Core.Interaces;
using Core.Tables.HeaderColumns;

namespace Core.Fabrics
{
    internal sealed class HeaderFactory : IHeader
    {
        public Header GetHeader<THeader>() => typeof(THeader) switch
        {
            Type _ when typeof(THeader) == typeof(BigStreamsColumns) => new BigStreamsColumns(),
            Type _ when typeof(THeader) == typeof(SmallStreamsColumns) => new SmallStreamsColumns(),
            _ => throw new InvalidCastException("Type is not supported")
        };
    }
}

using Core.Abstracts;
using Core.Data;
using Core.Data.PiktsInfo;
using Core.Data.Watercourse;

namespace Core.Tables.HeaderColumns
{
    public sealed class BigStreamsColumns : Header
    {
        public override IReadOnlyCollection<ColumnInfo> Columns => _columns.AsReadOnly();
        private static ColumnInfo[] _columns = null!;
        public BigStreamsColumns()
        {
            if(_columns == null)
            {
                var table = new BigStreamsTable();
                var watercourseInfo = SetHeaderColumns<Watercourse, BigStreamsTable>(table);
                var piktsInfo = SetHeaderColumns<Pikts, BigStreamsTable>(table);

                _columns = watercourseInfo.Concat(piktsInfo).ToArray();
            }
            
        }
    }
}

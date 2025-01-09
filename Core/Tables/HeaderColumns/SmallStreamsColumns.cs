using Core.Abstracts;
using Core.Data;
using Core.Data.Watercourse;

namespace Core.Tables.HeaderColumns
{
    public class SmallStreamsColumns : Header
    {
        public override IReadOnlyCollection<ColumnInfo> Columns => _columns.AsReadOnly();
        private static ColumnInfo[] _columns = null!;
        public SmallStreamsColumns()
        {
            if (_columns == null)
            {
                var table = new SmallStreamsTable();
                _columns = SetHeaderColumns<Watercourse, SmallStreamsTable>(table).ToArray();
            }
        }
    }
}

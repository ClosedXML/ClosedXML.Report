using ClosedXML.Excel;

namespace ClosedXML.Report
{
    internal interface IDataSource
    {
        object GetValue(IXLRangeRow row);
        object[] GetAll();
    }

    internal class DataSource : IDataSource
    {
        private readonly object[] _items;

        public DataSource(object[] items)
        {
            _items = items;
        }

        public object GetValue(IXLRangeRow row)
        {
            if (row.LastCell().TryGetValue(out int key))
            {
                if (key >= 0 && key < _items.Length)
                    return _items[key];
            }
            return null;
        }

        public object[] GetAll()
        {
            return _items;
        }
    }
}

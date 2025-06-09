using System.Drawing;
using LinqToDB.Mapping;

namespace ClosedXML.Report.Tests.TestModels
{
    public partial class Item
    {
        [Association(ThisKey = "PartNo", OtherKey = "PartNo")]
        public Part Part { get; set; }

        [Association(ThisKey = "OrderNo", OtherKey = "OrderNo")]
        public Order Order { get; set; }

        public Bitmap IsOk => Resource.checkmark;
    }
}

using System.Drawing;
using LinqToDB.Mapping;

namespace ClosedXML.Report.Tests.TestModels
{
    public partial class item
    {
        [Association(ThisKey = "PartNo", OtherKey = "PartNo")]
        public part Part { get; set; }

        [Association(ThisKey = "OrderNo", OtherKey = "OrderNo")]
        public order Order { get; set; }

        public Bitmap IsOk => Resource.checkmark;
    }
}

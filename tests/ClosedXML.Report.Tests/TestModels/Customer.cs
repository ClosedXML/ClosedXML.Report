using System.Collections.Generic;
using System.Drawing;
using LinqToDB.Mapping;

namespace ClosedXML.Report.Tests.TestModels
{
    public partial class Customer
    {
        [Association(ThisKey = "CustNo", OtherKey = "CustNo")]
        public List<Order> Orders { get; set; }

        public Bitmap Logo { get; set; }
    }
}

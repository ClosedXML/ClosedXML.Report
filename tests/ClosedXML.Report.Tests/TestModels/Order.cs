using System;
using System.Collections.Generic;
using System.Drawing;
using LinqToDB.Mapping;

namespace ClosedXML.Report.Tests.TestModels
{
    public partial class order
    {
        [Association(ThisKey = "CustNo", OtherKey = "CustNo", CanBeNull = true, KeyName = "FK_Orders_Customers", BackReferenceName = "Orders")]
        public customer Customer { get; set; }

        [Association(ThisKey = "OrderNo", OtherKey = "OrderNo", IsBackReference = true)]
        public List<item> Items { get; set; }

        public Bitmap PaymentImage { get; set; }
    }
}

using System.Collections.Generic;
using System.Drawing;
using LinqToDB.Mapping;

namespace ClosedXML.Report.Tests.TestModels
{
    public partial class Order
    {
        [Association(ThisKey = "CustNo", OtherKey = "CustNo", CanBeNull = true)]
        public Customer Customer { get; set; }

        [Association(ThisKey = "OrderNo", OtherKey = "OrderNo")]
        public List<Item> Items { get; set; }

        public Bitmap PaymentImage
        {
            get
            {
                switch (PaymentMethod)
                {
                    case "Visa": return Resource.card;
                    case "Cash": return Resource.cash;
                    case "Credit": return Resource.bank;
                    default: return null;
                }
            }
        }
    }
}

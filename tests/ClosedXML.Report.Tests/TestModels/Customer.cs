using System.Collections.Generic;
using System.Drawing;
using LinqToDB.Mapping;

namespace ClosedXML.Report.Tests.TestModels
{
    public partial class customer
    {
        [Association(ThisKey = "CustNo", OtherKey = "CustNo")]
        public List<order> Orders { get; set; }

        public Bitmap Logo { get; set; }
    }
}

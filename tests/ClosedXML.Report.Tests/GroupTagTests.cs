using System.IO;
using System.Linq;
using ClosedXML.Report.Tests.TestModels;
using LinqToDB;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class GroupTagTests : XlsxTemplateTestsBase
    {
        public GroupTagTests(ITestOutputHelper output) : base(output)
        {
        }

        [Theory,
         InlineData("GroupTagTests_Simple.xlsx"),
         InlineData("GroupTagTests_Collapse.xlsx"),
         InlineData("tLists2_sum.xlsx"),
        ]
        public void Simple(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var cust = db.customers.LoadWith(x => x.Orders).OrderBy(c => c.CustNo).First();
                        tpl.AddVariable(cust);
                    }
                },
                wb =>
                {
                    CompareWithGauge(wb, templateFile);
                });
        }

        [Fact]
        public void WithHeader()
        {
            XlTemplateTest("GroupTagTests_WithHeader.xlsx",
                tpl =>
                {
                    using (var db = new DbDemos())
                        tpl.AddVariable("Orders", db.orders.LoadWith(x => x.Customer).OrderBy(c => c.OrderNo).ToArray());
                },
                wb =>
                {
                    CompareWithGauge(wb, "GroupTagTests_WithHeader.xlsx");
                });
        }
    }
}

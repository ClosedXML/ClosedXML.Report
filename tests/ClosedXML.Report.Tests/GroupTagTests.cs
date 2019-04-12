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
         InlineData("GroupTagTests_Simple_WithOutsideLink.xlsx"),
         InlineData("GroupTagTests_Collapse.xlsx"),
         InlineData("tLists1_sort.xlsx"),
         InlineData("tLists2_sum.xlsx"),
         InlineData("tLists3_options.xlsx"),
         InlineData("tLists4_complexRange.xlsx"),
         InlineData("tLists5_GlobalVars.xlsx"),
         InlineData("tPage1_options.xlsx"),
        ]
        public void Simple(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var cust = db.customers.LoadWith(x => x.Orders).OrderBy(c => c.CustNo).First(x=>x.CustNo == 1356);
                        tpl.AddVariable(cust);
                    }
                    tpl.AddVariable("Tax", 13);
                },
                wb =>
                {
                    CompareWithGauge(wb, templateFile);
                });
        }

        [Theory,
         InlineData("GroupTagTests_SummaryAbove.xlsx"),
         InlineData("GroupTagTests_MergeLabels.xlsx"),
         InlineData("GroupTagTests_MergeLabels2.xlsx"),
         InlineData("GroupTagTests_PlaceToColumn.xlsx"),
         InlineData("GroupTagTests_NestedGroups.xlsx"),
         InlineData("GroupTagTests_DisableOutline.xlsx"),
         InlineData("GroupTagTests_FormulasInGroupRow.xlsx"),
        ]
        public void Customers(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var orders = db.orders.LoadWith(x => x.Customer);
                        tpl.AddVariable("Orders", orders);
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

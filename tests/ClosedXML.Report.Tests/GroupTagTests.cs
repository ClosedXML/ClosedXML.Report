using System;
using System.Collections;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using LinqToDB;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class GroupTagTests: XlsxTemplateTestsBase
    {
        private readonly ITestOutputHelper _output;
        public GroupTagTests(ITestOutputHelper output) : base(output)
        {
            _output = output;
        }

        [Fact]
        public void Simple()
        {
            XlTemplateTest("GroupTagTests_Simple.xlsx",
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var cust = db.customers.LoadWith(x=>x.Orders).OrderBy(c=>c.CustNo).First();
                        tpl.AddVariable(cust);
                    }
                },
                wb =>
                {
                    //wb.SaveAs("GroupTagTests_Simple.xlsx");
                    CompareWithGauge(wb, "GroupTagTests_Simple.xlsx");
                });
        }

        [Fact]
        public void WithCollapseOption()
        {
            XlTemplateTest("GroupTagTests_Collapse.xlsx",
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
                    //wb.SaveAs("GroupTagTests_Collapse.xlsx");
                    CompareWithGauge(wb, "GroupTagTests_Collapse.xlsx");
                });
        }

        [Fact]
        public void WithHeader()
        {
            XlTemplateTest("GroupTagTests_WithHeader.xlsx",
                tpl =>
                {
                    using (var db = new DbDemos())
                        tpl.AddVariable("Orders", db.orders.LoadWith(x=>x.Customer).OrderBy(c => c.OrderNo).ToArray());
                },
                wb =>
                {
                    wb.SaveAs("GroupTagTests_WithHeader.xlsx");
                    CompareWithGauge(wb, "GroupTagTests_WithHeader.xlsx");
                });
        }

        [Fact]
        public void SumWithoutGroup()
        {
            XlTemplateTest("tLists2_sum.xlsx",
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
                    //wb.SaveAs("tLists2_sum.xlsx");
                    CompareWithGauge(wb, "tLists2_sum.xlsx");
                });
        }
    }
}
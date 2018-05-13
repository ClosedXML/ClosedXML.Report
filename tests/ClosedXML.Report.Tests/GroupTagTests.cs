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
                    using (var ms = new MemoryStream())
                        wb.SaveAs(ms); // as conditional formats are consolidated on saving
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
                    using (var ms = new MemoryStream())
                        wb.SaveAs(ms); // as conditional formats are consolidated on saving
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
                    using (var ms = new MemoryStream())
                        wb.SaveAs(ms); // as conditional formats are consolidated on saving
                    //wb.SaveAs("GroupTagTests_WithHeader.xlsx");
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
                    using (var ms = new MemoryStream())
                        wb.SaveAs(ms); // as conditional formats are consolidated on saving
                    //wb.SaveAs("tLists2_sum.xlsx");
                    CompareWithGauge(wb, "tLists2_sum.xlsx");
                });
        }

        [Fact]
        public void TestName()
        {
            var fileName = Path.Combine(TestConstants.GaugesFolder, "Subranges_Simple_tMD1.xlsx");
            var workbook = new XLWorkbook(fileName);
            _output.WriteLine(DateTime.Now.ToLongTimeString());
            var srcSheet = workbook.Worksheet(1);

            srcSheet.Row(8).InsertRowsAbove(1);

            var dstSheet = workbook.AddWorksheet("Copy");

            for (int i = 0; i < 147; i++)
            {
                var srcRng = srcSheet.Range(i * 10 + 1, 1, i * 10 + 10, 9);
                var dstRng = dstSheet.Cell(i * 10 + 1, 1);
                srcRng.CopyTo(dstRng);
            }
            _output.WriteLine(DateTime.Now.ToLongTimeString());
            //workbook.SaveAs(Path.Combine(TestConstants.GaugesFolder, "bigcopy.xlsx"));
        }
    }
}
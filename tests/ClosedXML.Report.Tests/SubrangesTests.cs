using System.IO;
using System.Linq;
using ClosedXML.Report.Tests.TestModels;
using LinqToDB;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class SubrangesTests : XlsxTemplateTestsBase
    {
        [Fact]
        public void Simple()
        {
            XlTemplateTest("Subranges_Simple_tMD1.xlsx",
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var items = db.items.ToList().GroupBy(i=>i.OrderNo).ToDictionary(x=>x.Key);
                        var parts = db.parts.ToList().ToDictionary(x=>x.PartNo);
                        customer[] custs = db.customers.LoadWith(x => x.Orders).OrderBy(x=>x.CustNo).ToArray();
                        foreach (var customer in custs)
                        {
                            customer.Orders.Sort((x, y) => x.OrderNo.CompareTo(y.OrderNo));
                            foreach (var o in customer.Orders)
                            {
                                var order = o;
                                o.Items = items[order.OrderNo].ToList();
                                o.Items.Sort((x,y)=>x.ItemNo.Value.CompareTo(y.ItemNo));
                                foreach (var item in o.Items)
                                    item.Part = parts[item.PartNo.Value];
                            }
                        }
                        //var cust = db.Customers.Include(x => x.Orders.Select(o=>o.Items.Select(i=>i.Part)));
                        tpl.AddVariable("Customers", custs);
                    }
                },
                wb =>
                {
#if SAVE_OUTPUT
                    wb.SaveAs("Output\\Subranges_Simple_tMD1.xlsx");
#else
                    using (var ms = new MemoryStream())
                        wb.SaveAs(ms); // as conditional formats are consolidated on saving
#endif
                    CompareWithGauge(wb, "Subranges_Simple_tMD1.xlsx");
                });
        }

        [Fact]
        public void WithSubtotals()
        {
            XlTemplateTest("Subranges_WithSubtotals_tMD2.xlsx",
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var items = db.items.ToList().GroupBy(i=>i.OrderNo).ToDictionary(x=>x.Key);
    var parts = db.parts.ToList().ToDictionary(x=>x.PartNo);
                        customer[] custs = db.customers.LoadWith(x => x.Orders).OrderBy(x=>x.CustNo).ToArray();
                        foreach (var customer in custs)
                        {
                            customer.Orders.Sort((x, y) => x.OrderNo.CompareTo(y.OrderNo));
                            foreach (var o in customer.Orders)
                            {
                                var order = o;
                                o.Items = items[order.OrderNo].ToList();
                                o.Items.Sort((x,y)=>x.ItemNo.Value.CompareTo(y.ItemNo));
                                foreach (var item in o.Items)
                                    item.Part = parts[item.PartNo.Value];
                            }
                        }
                        //var cust = db.Customers.Include(x => x.Orders.Select(o=>o.Items.Select(i=>i.Part)));
                        tpl.AddVariable("Customers", custs);
                    }
                },
                wb =>
                {
#if SAVE_OUTPUT
                    wb.SaveAs("Output\\Subranges_WithSubtotals_tMD2.xlsx");
#else
                    using (var ms = new MemoryStream())
                        wb.SaveAs(ms); // as conditional formats are consolidated on saving
#endif
                    CompareWithGauge(wb, "Subranges_WithSubtotals_tMD2.xlsx");
                });
        }

        public SubrangesTests(ITestOutputHelper output) : base(output)
        {
        }
    }
}
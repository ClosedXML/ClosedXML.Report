﻿using System.Collections.Generic;
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

        [Fact]
        public void MasterDetailWithEmptySubsetCorrectSum()
        {
            XlTemplateTest("Visitors.xlsx",
                tpl =>
                {
                    tpl.AddVariable("Visitors", GenerateVisitors());
                },
                wb =>
                {
#if SAVE_OUTPUT
                    wb.SaveAs("Output\\MasterDetailWithEmptySubsetCorrectSum.xlsx");
#endif
                    CompareWithGauge(wb, "MasterDetailWithEmptySubset.xlsx");
                });
        }

        [Fact]
        public void MasterDetailWithSingleEmptySubsetCorrectSum()
        {
            XlTemplateTest("Visitors.xlsx",
                tpl =>
                {
                    var visitors = new List<dynamic> { GenerateVisitors().First() };
                    tpl.AddVariable("Visitors", visitors);
                },
                wb =>
                {
#if SAVE_OUTPUT
                    wb.SaveAs("Output\\MasterDetailWithSingleEmptySubsetCorrectSum.xlsx");
#endif
                    CompareWithGauge(wb, "MasterDetailWithSingleEmptySubset.xlsx");
                });
        }

        [Fact]
        public void SingleEmptySubsetCorrectSum()
        {
            XlTemplateTest("Visitor.xlsx",
                tpl =>
                {
                    tpl.AddVariable(GenerateVisitors().First());
                },
                wb =>
                {
#if SAVE_OUTPUT
                    wb.SaveAs("Output\\SingleEmptySubsetCorrectSum.xlsx");
#endif
                    CompareWithGauge(wb, "SingleEmptySubset.xlsx");
                });
        }

        private IEnumerable<dynamic> GenerateVisitors()
        {
            return new List<dynamic>
                {
                    new { Name = "Alice", Attendance = new List<dynamic> { } },
                    new { Name = "Bob", Attendance = new List<dynamic> {
                        new { Month = "February", Visits = 2 },
                        new { Month = "March", Visits = 3 },
                        new { Month = "April", Visits = 7 },
                    } },
                    new { Name = "Carl", Attendance = new List<dynamic> {
                        new { Month = "January", Visits = 5 },
                        new { Month = "July", Visits = 8 },
                        new { Month = "October", Visits = 6 },
                    } },
                    new { Name = "Daniel", Attendance = new List<dynamic> { } },
                };
        }


        public SubrangesTests(ITestOutputHelper output) : base(output)
        {
        }
    }
}

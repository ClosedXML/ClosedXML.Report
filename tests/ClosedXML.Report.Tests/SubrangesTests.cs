using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Report.Tests.TestModels;
using LinqToDB;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{

    [Collection("Database")]
    public class SubrangesTests : XlsxTemplateTestsBase
    {
        [Theory]
        [InlineData("Subranges_Simple_tMD1.xlsx")]
        [InlineData("Subranges_WithSubtotals_tMD2.xlsx")]
        [InlineData("Subranges_WithSort_tMD3.xlsx", Skip = "ClosedXML 0.104.1 bug. When a sheet is copied, it doesn't insert links to calculation chain for cells with formulas, but then requires it.")]
        public void Simple(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var custs = GetCustomers(db);
                        tpl.AddVariable("Customers", custs);
                    }
                },
                wb =>
                {
                    CompareWithGauge(wb, templateFile);
                });
        }

        [Fact]
        public void MultipleSubRanges()
        {
            var random = new Random(1234);
            var templateFile = "Subranges_Multiple.xlsx";
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var custs = GetCustomers(db).Select(cust =>
                        new {
                            CustNo = cust.CustNo,
                            Company = cust.Company,
                            Orders = cust.Orders,
                            Visitors = new List<dynamic>
                            {
                                new { Name = "Alice", Age = random.Next(0, 100), Gender = "F" },
                                new { Name = "Bob", Age = random.Next(0, 100), Gender = "M" },
                                new { Name = "John", Age = random.Next(0, 100), Gender = "M" },
                            }
                        });
                        tpl.AddVariable("Customers", custs);
                        tpl.AddVariable("user", "John Doe");
                    }
                },
                wb =>
                {
                    CompareWithGauge(wb, templateFile);
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
                    CompareWithGauge(wb, "SingleEmptySubset.xlsx");
                });
        }

        private static Customer[] GetCustomers(DbDemos db)
        {
            var items = db.items.ToList().GroupBy(i => i.OrderNo).ToDictionary(x => x.Key);
            var parts = db.parts.ToList().ToDictionary(x => x.PartNo);
            var customers = db.customers.LoadWith(x => x.Orders).OrderBy(x => x.CustNo).ToArray();
            foreach (var customer in customers)
            {
                customer.Orders.Sort((x, y) => x.OrderNo.CompareTo(y.OrderNo));
                foreach (var order in customer.Orders)
                {
                    order.Items = items[order.OrderNo].ToList();
                    order.Items.Sort((x, y) => x.ItemNo.Value.CompareTo(y.ItemNo));
                    foreach (var item in order.Items)
                        item.Part = parts[item.PartNo.Value];
                }
            }
            //var cust = db.Customers.Include(x => x.Orders.Select(o=>o.Items.Select(i=>i.Part)));
            return customers;
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

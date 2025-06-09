using System;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using FluentAssertions;
using LinqToDB;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    [Collection("Database")]
    public class ReportOptionsTests : XlsxTemplateTestsBase
    {
        [Fact]
        public void Hidden_option_should_hide_sheet()
        {
            XlTemplateTest("5_options.xlsx",
                tpl => { },
                wb =>
                {
                    wb.Worksheets.Count.Should().Be(3);
                    wb.Worksheets.Count(x => x.Visibility == XLWorksheetVisibility.Visible).Should().Be(2);
                });
        }

        [Fact]
        public void OnlyValues_option_should_remove_formulas_on_sheet()
        {
            XlTemplateTest("5_options.xlsx",
                tpl => { },
                wb =>
                {
                    var worksheet = wb.Worksheet(1);
                    worksheet.Cell("B3").HasFormula.Should().BeFalse();
                    worksheet.Cell("B3").GetValue<string>().Should().Be("Begin at 19.01.2023");
                    worksheet = wb.Worksheet(3);
                    worksheet.Cell("B4").HasFormula.Should().BeFalse();
                    worksheet.Cell("B4").GetValue<int>().Should().Be(10);
                });
        }

        [Fact]
        public void ColsFit_option_should_FitWidth()
        {
            XlTemplateTest("5_options.xlsx",
                tpl => { },
                wb =>
                {
                    var worksheet = wb.Worksheet(1);
                    worksheet.Column(4).Width.Should().BeApproximately(5.0, 0.01);
                    worksheet.Column(5).Width.Should().BeApproximately(16.16, 0.01);
                });
        }

        [Fact]
        public void Sort_option_should_sort_range()
        {
            var testEntities = TestEntity.GetTestData(6).ToArray();
            XlTemplateTest("8_sort.xlsx",
                tpl => tpl.AddVariable(new
                {
                    data = testEntities,
                    dates = new[] { DateTime.Parse("2013-01-01"), DateTime.Parse("2013-01-02"), DateTime.Parse("2013-01-03") }
                }),
                wb =>
                {
                    var worksheet = wb.Worksheet(1);
                    var expectedOrder = testEntities.OrderBy(x=>x.Address.City).ThenBy(x=>x.Age).ToArray();
                    worksheet.Range("D5:D10").Cells().Select(x=>x.GetValue<int>()).ToArray().Should().ContainInOrder(expectedOrder.Select(x => x.Age));
                    worksheet.Range("E5:E10").Cells().Select(x=>x.GetString()).ToArray().Should().ContainInOrder(expectedOrder.Select(x => x.Address.City));
                });
        }

        [Fact]
        public void DeleteOptionsWithParameter()
        {
            XlTemplateTest("delete_options.xlsx",
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var cust = db.customers.LoadWith(x => x.Orders.First().Items).OrderBy(c => c.CustNo).First(x => x.CustNo == 1356);
                        tpl.AddVariable(cust);
                    }
                    tpl.AddVariable("disableCColumnDeletion", "true");
                    tpl.AddVariable("disableEColumnDeletion", "false");
                },
                wb =>
                {
                    CompareWithGauge(wb, "delete_options.xlsx");
                });
        }

        public ReportOptionsTests(ITestOutputHelper output) : base(output)
        {
        }
    }
}

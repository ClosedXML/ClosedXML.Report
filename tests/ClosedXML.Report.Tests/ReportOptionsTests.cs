using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
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
            //TODO make tests culture-independent
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("ru-RU");
            XlTemplateTest("5_options.xlsx",
                tpl => { },
                wb =>
                {
                    var worksheet = wb.Worksheet(1);
                    worksheet.Cell("B3").HasFormula.Should().BeFalse();
                    worksheet.Cell("B3").GetValue<string>().Should().Be("Begin at " + DateTime.Today.Subtract(TimeSpan.FromDays(7)).ToShortDateString());
                    worksheet = wb.Worksheet(3);
                    worksheet.Cell("B4").HasFormula.Should().BeFalse();
                    worksheet.Cell("B4").GetValue<int>().Should().Be(10);
                });
        }

        [Fact]
        public void Sort_option_should_sort_range()
        {
            XlTemplateTest("8_sort.xlsx",
                tpl => tpl.AddVariable(new
                {
                    data = TestEntity.GetTestData(6),
                    dates = new[] { DateTime.Parse("2013-01-01"), DateTime.Parse("2013-01-02"), DateTime.Parse("2013-01-03") }
                }),
                wb =>
                {
                    var worksheet = wb.Worksheet(1);
                    worksheet.Range("D5:D10").Cells().Select(x=>x.GetValue<int>()).ToArray().Should().ContainInOrder(new [] { 37, 31, 71, 24, 63, 76});
                    worksheet.Range("E5:E10").Cells().Select(x=>x.GetString()).ToArray().Should().ContainInOrder("Dallas", "Miami", "Montana", "NY", "NY", "Oklahoma");
                });
        }

        public ReportOptionsTests(ITestOutputHelper output) : base(output)
        {
        }
    }
}
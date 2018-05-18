using ClosedXML.Excel;
using FluentAssertions;
using System.Collections.Generic;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class RangeInterpreterTests: XlsxTemplateTestsBase
    {
        public RangeInterpreterTests(ITestOutputHelper output) : base(output)
        {
        }

        [Fact]
        public void ParseTags_shold_remove_all_tags()
        {
            XlTemplateTest("5_options.xlsx", tpl => {},
                wb =>
                {
                    wb.Worksheet(1).Cell("A2").IsEmpty().Should().BeTrue();
                    wb.Worksheet(2).Cell("A2").IsEmpty().Should().BeTrue();
                    wb.Worksheet(3).Cell("B4").GetString().Should().NotContain("<<OnlyValues>>");
                });
        }

        [Fact]
        public void DoNotEvaluateFormulaOnTagsParsing()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                var ws2 = wb.AddWorksheet("Sheet2");

                ws1.FirstCell().FormulaA1 = "=VLOOKUP(\"Bob\", Sheet2!B:D, 3, FALSE)";
                ws2.Cell("B2").Value = "{{item.Name}}";
                ws2.Cell("C2").Value = "{{item.Count}}";
                ws2.Cell("D2").Value = "&=C2*10";
                ws2.Range("A2:D3").AddToNamed("Items");

                var template = new XLTemplate(wb);
                template.AddVariable("Items", GenerateItems());
                template.Generate();

                ws1.FirstCell().Value.Should().Be(20.0);
            }

            IEnumerable<object> GenerateItems()
            {
                return new List<object>
                {
                    new { Name = "Alice", Count = 1 },
                    new { Name = "Bob", Count = 2 },
                    new { Name = "Carl", Count = 3 },
                };
            }
        }
    }
}

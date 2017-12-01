using FluentAssertions;
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
    }
}
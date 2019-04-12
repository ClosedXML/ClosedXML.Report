using System;
using ClosedXML.Excel;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class TemplateErrorHandlingTests: IDisposable
    {
        private XLWorkbook _wb;
        private IXLWorksheet _ws;
        private XLTemplate _template;

        public TemplateErrorHandlingTests()
        {
            _wb = new XLWorkbook();
            _ws = _wb.AddWorksheet("report");
            _template = new XLTemplate(_wb);
        }

        [Fact]
        public void UnknownVariableNameMustBeAddedToErrorList()
        {
            _ws.Cell("A1").Value = "{{unknown_variable1}}";
            _ws.Cell("A2").Value = "{{unknown_variable2}}";

            var result = _template.Generate();

            result.HasErrors.Should().BeTrue();
            result.ParsingErrors.Count.Should().Be(2);
            result.ParsingErrors[0].Message.Should().Be("Unknown identifier 'unknown_variable1'");
            result.ParsingErrors[0].Range.RangeAddress.ToString().Should().Be("A1:A1");
            result.ParsingErrors[1].Message.Should().Be("Unknown identifier 'unknown_variable2'");
            result.ParsingErrors[1].Range.RangeAddress.ToString().Should().Be("A2:A2");
        }

        [Fact]
        public void MissingRangeShouldBeAddedToErrorList()
        {
            _ws.Cell("B2").Value = "{{item.Name}}";
            _ws.Cell("C2").Value = "{{item.Value}}";

            var result = _template.Generate();

            result.ParsingErrors.Count.Should().Be(3);
            result.ParsingErrors[0].Message.Should().Be("The range does not meet the requirements of the list ranges. For details, see the documentation.");
            result.ParsingErrors[0].Range.RangeAddress.ToString().Should().Be("A1:A1");
            result.ParsingErrors[1].Message.Should().Be("Unknown identifier 'item'");
            result.ParsingErrors[1].Range.RangeAddress.ToString().Should().Be("B2:B2");
            result.ParsingErrors[2].Message.Should().Be("Unknown identifier 'item'");
            result.ParsingErrors[2].Range.RangeAddress.ToString().Should().Be("C2:C2");
        }

        [Fact]
        public void GroupTagOutsideRangeShouldBeAddedToErrorList()
        {
            _ws.Cell("B3").Value = "<<group>>";

            var result = _template.Generate();

            result.ParsingErrors.Count.Should().Be(1);
            result.ParsingErrors[0].Message.Should().Be("The GROUP tag can't be used outside the named range.");
            result.ParsingErrors[0].Range.RangeAddress.ToString().Should().Be("B3:B3");
        }

        [Fact]
        public void UnknownVariableInsideRangeMustBeAddedToErrorList()
        {
            _ws.Range("A2", "F3").AddToNamed("DataRange");
            _ws.Cell("B2").Value = "{{item.Name}}";
            _ws.Cell("C2").Value = "{{item.unknown_variable}}";

            _template.AddVariable("DataRange", new[] {new {Name = "test1"}, new {Name = "test2"}, new {Name = "test3"}});
            var result = _template.Generate();

            result.ParsingErrors.Count.Should().Be(1);
            result.ParsingErrors[0].Message.Substring(0, 54).Should().Be("No property or field 'unknown_variable' exists in type");
            result.ParsingErrors[0].Range.RangeAddress.ToString().Should().Be("C2:C2");
        }

        public void Dispose()
        {
            _template?.Dispose();
            _wb?.Dispose();
        }
    }
}

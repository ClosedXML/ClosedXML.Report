using System.Linq;
using ClosedXML.Report.Tests.TestModels;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests;

[Collection("Database")]
public class HeightTagTests : XlsxTemplateTestsBase
{
    public HeightTagTests(ITestOutputHelper output) : base(output)
    {
    }

    [Theory,
        InlineData("HeightTag.xlsx")
    ]
    public void Height(string templateFile)
    {
        XlTemplateTest(templateFile,
            tpl =>
            {

            },
            wb =>
            {
                CompareWithGauge(wb, templateFile);
            });
    }

    [Theory,
     InlineData("HeightRangeTag.xlsx")
    ]
    public void HeightRange(string templateFile)
    {
        XlTemplateTest(templateFile,
            tpl =>
            {
                using var db = new DbDemos();
                var cust = db.employees.Take(2).ToList();
                var dataTable = new
                {
                    Table = cust
                };
                tpl.AddVariable(dataTable);
            },
            wb =>
            {
                CompareWithGauge(wb, templateFile);
            });
    }
}

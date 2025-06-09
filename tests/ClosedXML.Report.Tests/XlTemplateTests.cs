using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using FluentAssertions;
using NSubstitute;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bogus;
using ClosedXML.Report.Excel;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class XlTemplateTests : XlsxTemplateTestsBase
    {
        public XlTemplateTests(ITestOutputHelper output) : base(output)
        {
        }

        [Fact]
        public void Issue246()
        {
            //Set the randomizer seed if you wish to generate repeatable data sets.
            Randomizer.Seed = new Random(8675309);
            var user = User.Generate(1).First();
            var entities = TestEntity.GetTestData(300);

            XlTemplateTest("SpreadReportTemplate.xlsx",
                tpl =>
                {
                    tpl.AddVariable(user);
                    tpl.AddVariable("EquipmentItemList", entities);
                    tpl.AddVariable("VehicleList", entities);
                },
                wb =>
                {
                });
        }

        [Fact]
        public void Add_simple_variable_should_replace_value_in_related_cell()
        {
            XlTemplateTest("1.xlsx",
                tpl => tpl.AddVariable(new { TestValue1 = "value from test", TestValue2 = 3.2 }),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell(1, 1).HasFormula.Should().BeFalse();
                    sheet.Cell(1, 1).GetValue<string>().Should().Be("value from test");
                    sheet.Cell(2, 2).FormulaA1.Should().Be("7.2*2");
                    sheet.Cell(2, 2).GetValue<double>().Should().Be(14.4);
                });
        }

        [Fact]
        public void Variables_test()
        {
            XlTemplateTest("7_vars.xlsx",
                tpl => {},
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell("B5").GetValue<string>().Should().Be("10");
                });
        }

        [Fact]
        public void Add_nullable_variable_should_replace_value_in_related_cell()
        {
            XlTemplateTest("1.xlsx",
                tpl => tpl.AddVariable(new {TestValue2 = (double?) 2.3}),
                wb => wb.Worksheet(1).Cell(2, 2).GetDouble().Should().Be((2.3 + 4) * 2)
            );
        }

        [Fact]
        public void Add_enumerable_of_simple_values_should_add_values_left_to_right()
        {
            XlTemplateTest("3.xlsx",
                tpl => tpl.AddVariable(new
                {
                    title = "title from test",
                    TestArray = new[] { 10, 22, 8, 4 }
                }),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell(1, 4).GetValue<string>().Should().Be("title from test");
                    sheet.Cell(2, 3).GetValue<int>().Should().Be(10);
                    sheet.Cell(2, 4).GetValue<int>().Should().Be(22);
                    sheet.Cell(2, 5).GetValue<int>().Should().Be(8);
                    sheet.Cell(2, 6).GetValue<int>().Should().Be(4);
                    sheet.Cell(2, 7).GetValue<int>().Should().Be(44);
                });
        }

        [Fact]
        public void Add_enumerable_variable_should_fill_range()
        {
            var testData = TestEntity.GetTestData(3).ToArray();
            XlTemplateTest("4.xlsx",
                tpl => tpl.AddVariable(new
                {
                    title = "title from test",
                    dates = new[] { DateTime.Parse("2013-01-01"), DateTime.Parse("2013-01-02"), DateTime.Parse("2013-01-03") },
                    PlanData = testData
                }),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell("H1").GetValue<string>().Should().Be("title from test");
                    sheet.Cell("B4").GetValue<string>().Should().Be("1");
                    sheet.Cell("B5").GetValue<string>().Should().Be("2");
                    sheet.Cell("B6").GetValue<string>().Should().Be("3");
                    sheet.Cell("C4").GetValue<string>().Should().Be(testData[0].Name);
                    sheet.Cell("C5").GetValue<string>().Should().Be(testData[1].Name);
                    sheet.Cell("C6").GetValue<string>().Should().Be(testData[2].Name);
                    sheet.Cell("D4").GetValue<string>().Should().Be(testData[0].Role);
                    sheet.Cell("D5").GetValue<string>().Should().Be(testData[1].Role);
                    sheet.Cell("D6").GetValue<string>().Should().Be(testData[2].Role);
                    sheet.Cell("E4").GetValue<int>().Should().Be(testData[0].Age);
                    sheet.Cell("E5").GetValue<int>().Should().Be(testData[1].Age);
                    sheet.Cell("E6").GetValue<int>().Should().Be(testData[2].Age);
                    sheet.Cell("F4").FormulaA1.Should().Be($"HYPERLINK(\"mailto:{testData[0].Email}\",\"{testData[0].Email}\")");
                    sheet.Cell("F5").FormulaA1.Should().Be($"HYPERLINK(\"mailto:{testData[1].Email}\",\"{testData[1].Email}\")");
                    sheet.Cell("F6").FormulaA1.Should().Be($"HYPERLINK(\"mailto:{testData[2].Email}\",\"{testData[2].Email}\")");
                    sheet.Cell("G4").GetValue<string>().Should().Be(testData[0].Address.City);
                    sheet.Cell("G5").GetValue<string>().Should().Be(testData[1].Address.City);
                    sheet.Cell("G6").GetValue<string>().Should().Be(testData[2].Address.City);
                    wb.DefinedName("PlanData")!.Ranges.First().RangeAddress.ToStringRelative().Should().Be("A4:K6");
                    sheet.Cell("H4").GetValue<int>().Should().Be(testData[0].Hours[0]);
                    sheet.Cell("H5").GetValue<int>().Should().Be(testData[1].Hours[0]);
                    sheet.Cell("H6").GetValue<int>().Should().Be(testData[2].Hours[0]);
                    sheet.Cell("I4").GetValue<int>().Should().Be(testData[0].Hours[1]);
                    sheet.Cell("I5").GetValue<int>().Should().Be(testData[1].Hours[1]);
                    sheet.Cell("I6").GetValue<int>().Should().Be(testData[2].Hours[1]);
                    sheet.Cell("J4").GetValue<int>().Should().Be(testData[0].Hours[2]);
                    sheet.Cell("J5").GetValue<int>().Should().Be(testData[1].Hours[2]);
                    sheet.Cell("J6").GetValue<int>().Should().Be(testData[2].Hours[2]);
                    sheet.Cell("K4").GetValue<int>().Should().Be(testData[0].Hours.Sum());
                    sheet.Cell("K5").GetValue<int>().Should().Be(testData[1].Hours.Sum());
                    sheet.Cell("K6").GetValue<int>().Should().Be(testData[2].Hours.Sum());
                    sheet.Cell("D8").GetValue<int>().Should().Be(15);
                    sheet.Cell("L6").GetValue<int>().Should().Be(4);
                });
        }

        [Fact]
        public void Add_complex_object_shold_replace_all_possible_values()
        {
            XlTemplateTest("2.xlsx",
                tpl => tpl.AddVariable(new
                {
                    title = "title from test",
                    birthdate = new DateTime(2009, 8, 17, 16, 40, 33),
                    dates = new[] { DateTime.Parse("2013-01-01"), DateTime.Parse("2013-01-02"), DateTime.Parse("2013-01-03") },
                    person = new
                    {
                        age = 35,
                        name = "Пупкин Иван",
                        car = new
                        {
                            brand = "Mercedes-Benz",
                            model = "C230"
                        }
                    },
                }),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell("B2").GetValue<string>().Should().Be("title from test");
                    sheet.Cell("C4").GetValue<int>().Should().Be(35);
                    sheet.Cell("C5").GetValue<DateTime>().Should().Be(new DateTime(2009, 8, 17, 16, 40, 33));
                    sheet.Cell("C6").GetValue<string>().Should().Be("Пупкин Иван");
                    sheet.Cell("C8").GetValue<string>().Should().Be("Mercedes-Benz");
                    sheet.Cell("C9").GetValue<string>().Should().Be("C230");
                    sheet.Cell("D11").GetValue<DateTime>().Should().Be(DateTime.Parse("2013-01-01"));
                    sheet.Cell("E11").GetValue<DateTime>().Should().Be(DateTime.Parse("2013-01-02"));
                    sheet.Cell("F11").GetValue<DateTime>().Should().Be(DateTime.Parse("2013-01-03"));
                    sheet.Cell("H11").GetValue<string>().Should().Be("should stay");
                }
            );
        }

        [Fact]
        public void RowWideRangesProcessedCorrectly()
        {
            var workbook = CreateWorkbook();
            var sheet = workbook.Worksheets.First();
            var items = GenerateItems();

            sheet.Range("2:3").AddToNamed("Items");

            var template = new XLTemplate(workbook);
            template.AddVariable("Items", items);
            template.Generate();

            sheet.Cell("B2").GetString().Should().Be("Alice");
            sheet.Cell("B3").GetString().Should().Be("Bob");
            sheet.Cell("B4").GetString().Should().Be("Carl");

            sheet.Cell("C2").GetDouble().Should().Be(20.0);
            sheet.Cell("C3").GetDouble().Should().Be(30.0);
            sheet.Cell("C4").GetDouble().Should().Be(38.0);

            sheet.Cell("F2").GetString().Should().Be("Placeholder");
            sheet.Cell("F3").GetString().Should().Be("Placeholder");
            sheet.Cell("F4").GetString().Should().Be("Placeholder");

            XLWorkbook CreateWorkbook()
            {
                var wb = new XLWorkbook();
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("B1").Value = "Name";
                ws.Cell("C1").Value = "Age";
                ws.Cell("B2").Value = "{{item.Name}}";
                ws.Cell("C2").Value = "{{item.Age}}";

                ws.Cell("F2").Value = "Placeholder";
                return wb;
            }

            IEnumerable<dynamic> GenerateItems()
            {
                return new List<dynamic>
                {
                    new { Name = "Alice", Age = 20},
                    new { Name = "Bob", Age = 30},
                    new { Name = "Carl", Age = 38},
                };
            }
        }

        [Fact]
        public void XLTemplateWithNoWorkbookFails()
        {
            IXLWorkbook wb = null;
            Assert.Throws<ArgumentNullException>(() => new XLTemplate(wb));
        }

        [Fact]
        public void XLTemplateOpenFromFile()
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, "1.xlsx");
            using (var template = new XLTemplate(fileName))
            {
                template.Workbook.Should().NotBeNull();
                template.Workbook.Worksheets.First().FirstCell().GetString().Should().Be("{{TestValue1}}");
            }
        }

        [Fact]
        public void XLTemplateOpenFromStream()
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, "1.xlsx");
            using (var stream = File.Open(fileName, FileMode.Open))
            {
                var template = new XLTemplate(stream);

                template.Workbook.Should().NotBeNull();
                template.Workbook.Worksheets.First().FirstCell().GetString().Should().Be("{{TestValue1}}");
            }
        }

        [Fact]
        public void AutoCreatedWorkbookDisposedWithTemplate()
        {
            var disposed = false;
            var wb = Substitute.For<IXLWorkbook>();
            wb.When(w => w.Dispose()).Do(w => disposed = true);

            var template = new XLTemplate(wb);

            template.Dispose();
            disposed.Should().BeFalse("Workbook specified in the constructor should not be disposed");
        }

        [Fact]
        public void WorkbookNotDisposedWithTemplate()
        {
            var disposed = false;
            var wb = Substitute.For<IXLWorkbook>();
            wb.When(w => w.Dispose()).Do(w => disposed = true);

            var fileName = Path.Combine(TestConstants.TemplatesFolder, "1.xlsx");
            var template = new XLTemplate(fileName);
            ReplaceWorkbookWithMock(template, wb);

            template.Dispose();

            disposed.Should().BeTrue("Workbook expected to be disposed");
        }

        [Fact]
        public void AccessToDisposedTemplateThrows()
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, "1.xlsx");
            var template = new XLTemplate(fileName);

            template.Dispose();

            Assert.Throws<ObjectDisposedException>(() => template.AddVariable("Test", "test"));
            Assert.Throws<ObjectDisposedException>(() => template.Generate());
        }


        [Fact]
        public void Leading_zeros_should_not_be_trimmed()
        {
            var data = new
            {
                Id = "01",
                Items = new object[] {
                    new { Id = "001" },
                    new { Id = "002" },
                }
            };

            XlTemplateTest("Leading_Zeros.xlsx",
                tpl => tpl.AddVariable(data),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell(2, 2).GetString().Should().Be("001");
                    sheet.Cell(3, 2).GetString().Should().Be("002");
                    sheet.Cell(1, 1).GetString().Should().Be("01");
                });
        }

        [Fact]
        public void DictionaryVariableTest()
        {
            var dic = new Dictionary<string, object>
            {
                { "Customer1", new Dictionary<string, object>{{"ID", "1"}, {"Name", "Customer 1"}}},
                { "Customer2", new Dictionary<string, object>{{"ID", "2"}, {"Name", "Customer 2"}}},
                { "Customer3", new Dictionary<string, object>{{"ID", "3"}, {"Name", "Customer 3"}}},
            };

            XlTemplateTest("DictionarySource.xlsx",
                tpl => tpl.AddVariable(dic),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell(1, 1).GetString().Should().Be("1");
                    sheet.Cell(1, 2).GetString().Should().Be("Customer 1");
                    sheet.Cell(2, 1).GetString().Should().Be("2");
                    sheet.Cell(2, 2).GetString().Should().Be("Customer 2");
                    sheet.Cell(3, 1).GetString().Should().Be("3");
                    sheet.Cell(3, 2).GetString().Should().Be("Customer 3");
                });
        }

        [Fact]
        public void ListOfDictionariesTest()
        {
            var dic = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>{{"ID", "1"}, {"Name", "Customer 1"}},
                new Dictionary<string, object>{{"ID", "2"}, {"Name", "Customer 2"}},
                new Dictionary<string, object>{{"ID", "3"}, {"Name", "Customer 3"}},
            };

            XlTemplateTest("ListDictionariesSource.xlsx",
                tpl => tpl.AddVariable("Customers", dic),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell(2, 2).GetString().Should().Be("1");
                    sheet.Cell(2, 3).GetString().Should().Be("Customer 1");
                    sheet.Cell(3, 2).GetString().Should().Be("2");
                    sheet.Cell(3, 3).GetString().Should().Be("Customer 2");
                    sheet.Cell(4, 2).GetString().Should().Be("3");
                    sheet.Cell(4, 3).GetString().Should().Be("Customer 3");
                });
        }

        [Fact]
        public void HorizontalSimpleRangeTest()
        {
            string[] GetRowData(IXLWorksheet sheet, int idx, int itemCnt, int offset)
            {
                var clmNumber = (char) ('a' + itemCnt * 2);
                var row = idx * 3 + offset;
                return sheet.Range($"B{row}:{clmNumber}{row}").Cells().Select(x => x.GetString()).ToArray();
            }

            var testData = TestOrder.GetTestData(4).ToArray();
            XlTemplateTest("Horizontal_SimpleTemplate.xlsx",
                tpl => tpl.AddVariable("Orders", testData),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    for (int i = 0; i < testData.Length; i++)
                    {
                        var itemCnt = testData[i].ProductsWithQuantities.Count;
                        sheet.Cell("B" + (i * 3 + 1)).GetString().Should().Be(testData[i].OrderNumber);
                        var header = GetRowData(sheet, i, itemCnt, 2);
                        header.Length.Should().Be(itemCnt * 2);
                        for (int j = 0; j < itemCnt; j += 2)
                        {
                            header[j].Should().Be("Name");
                            header[j+1].Should().Be("Quantity");
                        }
                        var data = GetRowData(sheet, i, itemCnt, 3);
                        data.Length.Should().Be(itemCnt * 2);
                        for (int j = 0; j < itemCnt; j++)
                        {
                            data[j*2].Should().Be(testData[i].ProductsWithQuantities[j].ProductName);
                            data[j*2+1].Should().Be(testData[i].ProductsWithQuantities[j].Quantity.ToString());
                        }
                    }
                });
        }

        [Fact]
        public void HorizontalRangeTest()
        {
            var testData = TestEntity.GetTestData(3).ToArray();
            XlTemplateTest("HorizontalRange.xlsx",
                tpl => tpl.AddVariable("Datas", testData),
                wb =>
                {
                    var sheet = wb.Worksheet(1);
                    sheet.Cell("C3").GetDouble().Should().Be(1d);
                    sheet.Cell("E3").GetDouble().Should().Be(2d);
                    sheet.Cell("G3").GetDouble().Should().Be(3d);
                    sheet.Cell("C4").GetString().Should().Be(testData[0].Role);
                    sheet.Cell("D4").GetString().Should().Be(testData[0].Name);
                    sheet.Cell("E4").GetString().Should().Be(testData[1].Role);
                    sheet.Cell("F4").GetString().Should().Be(testData[1].Name);
                    sheet.Cell("G4").GetString().Should().Be(testData[2].Role);
                    sheet.Cell("H4").GetString().Should().Be(testData[2].Name);
                    sheet.Cell("C5").GetString().Should().Be(testData[0].Address.City);
                    sheet.Cell("E5").GetString().Should().Be(testData[1].Address.City);
                    sheet.Cell("G5").GetString().Should().Be(testData[2].Address.City);
                });
        }

        private void ReplaceWorkbookWithMock(XLTemplate template, IXLWorkbook mock)
        {
            var property = template.GetType().GetProperty("Workbook");
            property.SetValue(template, mock);
        }
    }
}

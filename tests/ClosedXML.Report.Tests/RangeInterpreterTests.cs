using ClosedXML.Excel;
using FluentAssertions;
using System.Collections.Generic;
using System.Linq;
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

                ws1.FirstCell().GetDouble().Should().Be(20.0);
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

        [Fact]
        public void CanBindRangesToRootVariableFields()
        {
            var entity = new Parent();
            var template = CreateBaseTemplate();
            var ws = template.Workbook.Worksheets.First();
            ws.Range("A3:B4").AddToNamed("Children");
            ws.Range("E2:E2").AddToNamed("Container_ItemsInContainer");

            template.AddVariable(entity);
            template.Generate();

            AssertResultIsCorrect(ws);
        }

        [Fact]
        public void CanBindRangesToAliasedVariableFields()
        {
            var entity = new Parent();
            var template = CreateBaseTemplate();
            var ws = template.Workbook.Worksheets.First();
            ws.Range("A3:B4").AddToNamed("Model_Children");
            ws.Range("E2:E2").AddToNamed("Model_Container_ItemsInContainer");

            template.AddVariable("Model", entity);
            template.Generate();

            AssertResultIsCorrect(ws);
        }

        [Fact]
        public void ShouldNotThrowExceptionIfAccessSomeChildrenNullProp()
        {
            var entity = new Order();

            var template = CreateOrderTemplate();
            var ws = template.Workbook.Worksheets.First();

            ws.Range("A3:B4").AddToNamed("Items");

            template.AddVariable(entity);
            template.Generate();

            ws.Cell("B3").GetString().Should().Be(string.Empty);
            ws.Cell("B4").GetString().Should().Be("Material 1");
        }

        [Fact]
        public void ShouldDestroyEmptyTable()
        {
            //See #251
            var template = CreateOrderTemplate();
            var ws = template.Workbook.Worksheets.First();

            ws.Cell("B4").SetValue("This list is empty");
            ws.Cell("B5").SetValue("{{item.Name}}");
            ws.Range("A5:B6").AddToNamed("Empty");

            ws.Cell("B7").SetValue("This list is populated");
            ws.Cell("B8").SetValue("{{item.Name}}");
            ws.Range("A8:B9").AddToNamed("Populated");

            var model = new
            {
                Empty = new List<Item>(),
                Populated = new[]
                {
                    new Item("It works", null)
                }
            };
            template.AddVariable(model);
            template.Generate();

            ws.Cell("B4").GetString().Should().Be("This list is empty");
            ws.Cell("B5").GetString().Should().Be("This list is populated");
            ws.Cell("B6").GetString().Should().Be("It works");
        }

        private XLTemplate CreateOrderTemplate()
        {
            var wbTemplate = new XLWorkbook();
            var ws = wbTemplate.AddWorksheet();

            ws.Cell("B3").Value = "{{item.Material.Name}}";

            return new XLTemplate(wbTemplate);
        }

        private XLTemplate CreateBaseTemplate()
        {
            var wbTemplate = new XLWorkbook();
            var ws = wbTemplate.AddWorksheet();

            ws.Cell("B1").Value = "{{Model.Name}}";
            ws.Cell("B2").Value = "Children:";
            ws.Cell("B3").Value = "{{item.ChildName}}";

            ws.Cell("D2").Value = "Items in container:";
            ws.Cell("E2").Value = "{{item.ChildName}}";

            return new XLTemplate(wbTemplate);
        }

        private void AssertResultIsCorrect(IXLWorksheet ws)
        {
            ws.Cell("B3").GetString().Should().Be("Child 1");
            ws.Cell("B5").GetString().Should().Be("Child 3");
            ws.Cell("E2").GetString().Should().Be("Item in container 1");
            ws.Cell("G2").GetString().Should().Be("Item in container 3");
        }


        private class Parent
        {
            public string Name => "Parent Name";
            public Container Container = new Container();
            public List<Child> Children { get; } = new List<Child>
            {
                new Child("Child 1"),
                new Child("Child 2"),
                new Child("Child 3"),
            };
        }

        private class Order
        {
            public string OrderNumber => "Order Number";
            public List<Item> Items { get; } = new List<Item>
            {
                new Item("noMaterial", null),
                new Item("withMaterial", new Material("Material 1"))
            };
        }

        private class Item
        {
            public string Name { get; private set; }
            public Material Material { get; private set; }

            public Item(string name, Material material)
            {
                Name = name;
                Material = material;
            }

            public Item AddMaterial(Material material)
            {
                Material = material;
                return this;
            }
        }

        private class Material
        {
            public string Name { get; }

            public Material(string name)
            {
                Name = name;
            }
        }

        public class Child
        {
            public string ChildName { get; }
            public Child(string childName)
            {
                ChildName = childName;
            }
        }

        public class Container
        {
            public List<Child> ItemsInContainer { get; } = new List<Child>
            {
                new Child("Item in container 1"),
                new Child("Item in container 2"),
                new Child("Item in container 3"),
            };
        }
    }
}

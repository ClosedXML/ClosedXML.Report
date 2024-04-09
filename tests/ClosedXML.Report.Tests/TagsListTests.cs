using System.Collections.Generic;
using System.Linq;
using ClosedXML.Report.Options;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class TagsListTests
    {
        [Fact]
        public void Add_with_same_priority_should_increase_count()
        {
            var errList = new TemplateErrors();
            var list = new TagsList(errList)
            {
                new GroupTag { Name = "val1" },
                new GroupTag { Name = "val2" }
            };
            list.Count.Should().Be(2);
        }

        [Fact]
        public void Items_should_be_sorted_by_priority()
        {
            var errList = new TemplateErrors();
            var list = new TagsList(errList)
            {
                new OnlyValuesTag { Name = "val3", Priority = 40 },
                new ProtectedTag { Name = "val4", Priority = 0 },
                new GroupTag { Name = "val1", Priority = 200 },
                new GroupTag { Name = "val2", Priority = 200 }
            };

            var expected = new List<string>() { "val1", "val2", "val3", "val4" };

            list.Count.Should().Be(4);
            list.Select(x => x.Name).Should().BeEquivalentTo(expected, options => options.WithStrictOrdering());
        }

        [Fact]
        public void Items_with_same_priority_should_be_sorted_by_row()
        {
            var errList = new TemplateErrors();
            var list = new TagsList(errList)
            {
                new GroupTag { Name = "val1", Priority = 200, Cell = new TemplateCell { Row = 3, Column = 1 } },
                new GroupTag { Name = "val2", Priority = 200, Cell = new TemplateCell { Row = 1, Column = 1 } }
            };

            var expected = new List<string>() { "val2", "val1" };

            list.Count.Should().Be(2);
            list.Select(x => x.Name).Should().BeEquivalentTo(expected, options => options.WithStrictOrdering());
        }

        [Fact]
        public void Items_with_same_priority_and_row_should_be_sorted_by_column()
        {
            var errList = new TemplateErrors();
            var list = new TagsList(errList)
            {
                new GroupTag { Name = "val1", Priority = 200, Cell = new TemplateCell { Row = 1, Column = 15 } },
                new GroupTag { Name = "val2", Priority = 200, Cell = new TemplateCell { Row = 1, Column = 1 } },
                new GroupTag { Name = "val3", Priority = 200, Cell = new TemplateCell { Row = 1, Column = 20 } }
            };

            var expected = new List<string>() { "val2", "val1", "val3" };

            list.Count.Should().Be(3);
            list.Select(x => x.Name).Should().BeEquivalentTo(expected, options => options.WithStrictOrdering());
        }

        [Fact]
        public void Get_by_type_should_return_tags_with_inherits()
        {
            var errList = new TemplateErrors();
            var list = new TagsList(errList)
            {
                new GroupTag { Name = "val1" },
                new GroupTag { Name = "val2" },
                new SortTag { Name = "val3" },
                new ProtectedTag { Name = "val4" },
                new DescTag { Name = "val5" },
                new ColsFitTag { Name = "val6" }
            };
            list.GetAll<SortTag>().Select(x => x.Name).Should().BeEquivalentTo("val1", "val2", "val3", "val5");
        }
    }
}

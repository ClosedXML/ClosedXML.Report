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
            var list = new TagsList
            {
                new GroupTag {Name = "val1"},
                new GroupTag {Name = "val2"}
            };
            list.Count.Should().Be(2);
        }

        [Fact]
        public void Items_should_be_sorted_by_priority()
        {
            var list = new TagsList
            {
                new GroupTag {Name = "val1"},
                new GroupTag {Name = "val2"},
                new OnlyValuesTag {Name = "val3"},
                new ProtectedTag {Name = "val4"}
            };
            list.Count.Should().Be(4);
            list.Select(x => x.Name).Should().BeEquivalentTo("val3", "val1", "val2", "val4");
        }

        [Fact]
        public void Get_by_type_should_return_tags_with_inherits()
        {
            var list = new TagsList
            {
                new GroupTag {Name = "val1"},
                new GroupTag {Name = "val2"},
                new SortTag {Name = "val3"},
                new ProtectedTag {Name = "val4"},
                new DescTag {Name = "val5"},
                new ColsFitTag {Name = "val6"}
            };
            list.GetAll<SortTag>().Select(x => x.Name).Should().BeEquivalentTo("val1", "val2", "val3", "val5");
        }
    }
}
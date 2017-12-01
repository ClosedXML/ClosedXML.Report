using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Report.Builders;
using ClosedXML.Report.Plugins.FieldPlugins;
using ClosedXML.Report.Plugins.ReportPlugin;
using ClosedXML.Report.Tests.TestModels;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class FiledBuilderTests
    {
        [Theory, MemberData("PrimitiveData")]
        public void Passing_primitive_should_create_PrimitiveFieldBuilder(object value)
        {
            var output = FieldBuilder.Create(value.GetType(), "test", new TestBuildContext());
            output.Should().BeOfType<PrimitiveFieldBuilder>();
        }

        [Theory, MemberData("ComplexData")]
        public void Passing_complex_should_create_ComplexFieldBuilder(object value)
        {
            var output = FieldBuilder.Create(value.GetType(), "test", new TestBuildContext());
            output.Should().BeOfType<ComplexFieldBuilder>();
        }

        [Fact]
        public void Passing_enumerable_complex_should_create_VerticalRangeBuilder()
        {
            var output = FieldBuilder.Create(typeof(Address[]), "test", new TestBuildContext());
            output.Should().BeOfType<VerticalRangeBuilder>();
        }

        [Fact]
        public void Passing_enumerable_primitive_should_create_HorizontalRangeBuilder()
        {
            var output = FieldBuilder.Create(typeof(DateTime[]), "test", new TestBuildContext());
            output.Should().BeOfType<HorizontalRangeBuilder>();
        }

        [Fact]
        public void Create_name_for_primitives_test()
        {
            var ctx = new TestBuildContext();
            var output = FieldBuilder.Create(typeof(int), "test", ctx);
            output.CreateNames();
            output.CreateNames();

            ctx.Names.Count.Should().Be(1);
            ctx.Names.Keys.Should().Contain("test");
        }

        [Fact]
        public void Create_name_for_complex_test()
        {
            var ctx = new TestBuildContext();
            var output = FieldBuilder.Create(typeof(TestEntity), "test", ctx);
            output.CreateNames();
            output.CreateNames();

            ctx.Names.Count.Should().BeGreaterOrEqualTo(7);
            ctx.Names.Keys.Should().Contain("test_Age");
            ctx.Names.Keys.Should().Contain("test_Address_City");
            ctx.Names.Keys.Should().Contain("_value");
        }

        [Fact]
        public void Create_name_for_enumerable_primitives_test()
        {
            var ctx = new TestBuildContext();
            var output = FieldBuilder.Create(typeof(int[]), "test", ctx);
            output.CreateNames();
            output.CreateNames();

            ctx.Names.Count.Should().BeGreaterOrEqualTo(1);
            ctx.Names.Keys.Should().Contain("_value");
        }

        [Fact]
        public void Create_name_for_enumerable_complex_test()
        {
            var ctx = new TestBuildContext();
            var output = FieldBuilder.Create(typeof(TestEntity[]), "test", ctx);
            output.CreateNames();
            output.CreateNames();

            ctx.Names.Count.Should().BeGreaterOrEqualTo(7);
            ctx.Names.Keys.Should().Contain("_Age");
            ctx.Names.Keys.Should().Contain("_Name");
            ctx.Names.Keys.Should().Contain("_Address_City");
            ctx.Names.Keys.Should().Contain("_value");
        }

        [Fact]
        public void Set_value_for_primitive_test()
        {
            var ctx = new TestBuildContext();
            var output = FieldBuilder.Create(typeof(string), "test", ctx);
            output.CreateNames();

            output.SetValue("test value");
            ctx.Names["test"].Should().Be("test value");

            output.SetValue(2);
            ctx.Names["test"].Should().Be(2);
        }

        [Fact]
        public void Set_value_for_complex_test()
        {
            var ctx = new TestBuildContext();
            var output = FieldBuilder.Create(typeof(TestEntity), "test", ctx);
            output.CreateNames();

            output.SetValue(new TestEntity("test name", "test role", 36, new [] { 23, 42 }) {Address = new Address("RUS", "SPB", "Nevskiy")});
            ctx.Names["test_Name"].Should().Be("test name");
            ctx.Names["test_Role"].Should().Be("test role");
            ctx.Names["test_Age"].Should().Be(36);
            ctx.Names["test_Address_City"].Should().Be("SPB");
        }

        public static IEnumerable<object[]> ComplexData
        {
            get
            {
                return new[]
                {
                    new object[] {new Address("", "", "")},
                    new object[] {new TestEntity("", "", 0, null),},
                };
            }
        }

        public static IEnumerable<object[]> PrimitiveData
        {
            get
            {
                return new[]
                {
                    new object[] {"test"},
                    new object[] {0},
                    new object[] {2},
                    new object[] {2.3f},
                    new object[] {4.2d},
                    new object[] {(double?) 3.1},
                    new object[] {DateTime.Today},
                    new object[] {true}
                };
            }
        }
    }
    public class TestBuildContext : IReportBuildContext
    {
        private int _idx = 1;
        public readonly Dictionary<string, object> Names = new Dictionary<string, object>();
        public readonly Dictionary<string, IXLNamedRange> Ranges = new Dictionary<string, IXLNamedRange>();

        public TestBuildContext()
        {
            ReportPlugins = new ReportPluginsRegister();
            FieldPlugins = new FieldPluginsRegister();
        }

        public int AllocateRow(int rowCnt = 1)
        {
            return _idx + rowCnt;
        }

        public IXLRangeBase GetHiddenCell(int rowIdx, int cIdx)
        {
            return null;
        }

        public bool ContainsName(string name)
        {
            return Names.ContainsKey(name);
        }

        public IXLNamedRange AddName(string name, int rIdx, int cIdx)
        {
            Names.Add(name, null);
            return null;
        }

        public void SetNameValue(string name, object value)
        {
            Names[name] = value;
        }

        public IXLNamedRange GetNamedRange(string alias)
        {
            return Ranges.ContainsKey(alias) ? Ranges[alias] : null;
        }

        public IXLRange GetHiddenRange(int height, int width)
        {
            AllocateRow(height);
            return null;
        }

        public ReportPluginsRegister ReportPlugins { get; private set; }
        public FieldPluginsRegister FieldPlugins { get; private set; }
    }
}
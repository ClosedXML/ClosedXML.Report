using System;
using System.IO;
using System.Linq.Expressions;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class SubtotalSummaryFuncTests : XlsxTemplateTestsBase
    {
        private IXLRange _rng;
        private XLWorkbook _workbook;

        public SubtotalSummaryFuncTests(ITestOutputHelper output) : base(output)
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, "9_plaindata.xlsx");
            _workbook = new XLWorkbook(fileName);
            _rng = _workbook.Range("range1");
            _rng.InsertColumnsAfter(1, true);
            var clmn = _rng.LastColumn().ColumnNumber() - _rng.FirstColumn().ColumnNumber() + 1;
            for (int i = 1; i <= _rng.RowCount(); i++)
            {
                _rng.Row(i).Cell(clmn).Value = i - 1;
            }
        }

        [Fact]
        public void SumIntTest()
        {
            var sum = new SubtotalSummaryFunc("sum", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => e.Compile();
            sum.DataSource = new DataSource(new object[] { new Test(1), new Test(2), new Test(3) });

            sum.Calculate(sum.DataSource).Should().Be(6);
        }

        [Fact]
        public void SumDoubleTest()
        {
            var sum = new SubtotalSummaryFunc("sum", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => e.Compile();
            sum.DataSource = new DataSource(new object[] { new Test(1.5), new Test(2d), new Test(3.5) });

            sum.Calculate(sum.DataSource).Should().Be(7d);
        }

        [Fact]
        public void SumTimeSpanTest()
        {
            var sum = new SubtotalSummaryFunc("sum", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(TimeSpan.FromHours(1)), new Test(TimeSpan.FromHours(2)), new Test(TimeSpan.FromHours(3)) });

            sum.Calculate(sum.DataSource).Should().Be(TimeSpan.FromHours(6));
        }

        [Fact]
        public void AverageIntTest()
        {
            var sum = new SubtotalSummaryFunc("average", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(1), new Test(2), new Test(3) });

            sum.Calculate(sum.DataSource).Should().Be(2);
        }

        [Fact]
        public void AverageDoubleTest()
        {
            var sum = new SubtotalSummaryFunc("average", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(1.5), new Test(2d), new Test(3.5), new Test(4.5) });

            sum.Calculate(sum.DataSource).Should().Be(2.875);
        }

        [Fact]
        public void AverageTimeSpanTest()
        {
            var sum = new SubtotalSummaryFunc("average", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(TimeSpan.FromHours(1)), new Test(TimeSpan.FromHours(2)), new Test(TimeSpan.FromHours(3)) });

            sum.Calculate(sum.DataSource).Should().Be(TimeSpan.FromHours(2));
        }

        [Fact]
        public void MinIntTest()
        {
            var sum = new SubtotalSummaryFunc("min", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(2), new Test(1), new Test(3) });

            sum.Calculate(sum.DataSource).Should().Be(1);
        }

        [Fact]
        public void MinDoubleTest()
        {
            var sum = new SubtotalSummaryFunc("min", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(2d), new Test(0.5), new Test(3.5) });

            sum.Calculate(sum.DataSource).Should().Be(0.5);
        }

        [Fact]
        public void MinDateTimeTest()
        {
            var sum = new SubtotalSummaryFunc("min", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(new DateTime(2017, 01, 01)), new Test(new DateTime(2016, 01, 01)), new Test(new DateTime(2018, 01, 01)) });

            sum.Calculate(sum.DataSource).Should().Be(new DateTime(2016, 01, 01));
        }

        [Fact]
        public void StDevDoubleTest()
        {
            var sum = new SubtotalSummaryFunc("stdev", 1);
            Expression<Func<Test, object>> e = o => o.val;
            sum.GetCalculateDelegate = type => (e.Compile());
            sum.DataSource = new DataSource(new object[] { new Test(10d), new Test(20d), new Test(30d), new Test(40d), new Test(50d), new Test(60d), new Test(70d) });

            ((double)sum.Calculate(sum.DataSource)).Should().BeInRange(21.6, 21.61);
        }

        private class Test
        {
            public dynamic val;

            public Test(dynamic val)
            {
                this.val = val;
            }
        }
    }
}

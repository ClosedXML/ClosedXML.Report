using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Linq.Expressions;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class FormulaEvaluatorTests
    {
        [Fact]
        public void FormulaEvaluatorTests1()
        {
            var eval = new FormulaEvaluator();
            eval.AddVariable("a", 2);
            eval.AddVariable("b", 3);
            eval.Evaluate("{{\"test\"}}").Should().Be("test");
            eval.Evaluate("{{a+b}}").Should().Be(5);
            eval.Evaluate("{{c}}+{{d}}={{c+d}}", new Parameter("c", 3), new Parameter("d", 6)).Should().Be("3+6=9");
            eval.Evaluate("{{c}}+{{d}}={{c+d}}", new Parameter("c", 7), new Parameter("d", 8)).Should().Be("7+8=15");
        }

        [Fact]
        public void ExpressionParseTest()
        {
            var customers = new Customer[]
            {
                new Customer {Id = 1, Name = "Customer1"},
                new Customer {Id = 2, Name = "Customer2"}
            }.AsEnumerable();

            string query = "customers.Where(c => c.Id == 1).OrderBy(c=> c.Name)";
            var lambda = DynamicExpressionParser.ParseLambda(new [] {Expression.Parameter(customers.GetType(), "customers")}, null, query);
            var dlg = lambda.Compile();
            dlg.DynamicInvoke(customers).Should().BeAssignableTo<IEnumerable<Customer>>();
            ((IEnumerable<Customer>) dlg.DynamicInvoke(customers)).Should().HaveCount(1);
            ((IEnumerable<Customer>) dlg.DynamicInvoke(customers)).First().Id.Should().Be(1);
        }

        class Customer
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }
    }
}

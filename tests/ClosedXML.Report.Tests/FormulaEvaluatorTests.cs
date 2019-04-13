using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Linq.Dynamic.Core.Exceptions;
using System.Linq.Expressions;
using FluentAssertions;
using NSubstitute.ExceptionExtensions;
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

        [Fact]
        public void MultipleExpressionsWithNullResult()
        {
            var eval = new FormulaEvaluator();
            eval.AddVariable("a", null);
            eval.AddVariable("b", 1);
            eval.Evaluate("{{a}}{{b}}").Should().Be(1);
            eval.Evaluate("{{b}}{{a}}").Should().Be("1");
        }

        [Fact]
        public void PassNullParameter()
        {
            var eval = new FormulaEvaluator();
            eval.Evaluate("{{\"Hello \"+a}}", new Parameter("a", null)).Should().Be("Hello ");
            eval.Evaluate("{{1+a}}", new Parameter("a", null)).Should().Be(null);
            //TODO: eval.Evaluate("{{\"City: \"+Iif(a==null, string.Empty, a.City}}", new Parameter("a", null)).Should().Be("City: ");
        }

        [Fact]
        public void WrongExpressionShouldThrowParseException()
        {
            var eval = new FormulaEvaluator();
            Assert.Throws<ParseException>(() => eval.Evaluate("{{missing}}"));
        }

        [Fact]
        public void ParseExceptionMessageShouldBeUnknownIdentifier()
        {
            var eval = new FormulaEvaluator();
            Assert.Throws<ParseException>(() => eval.Evaluate("{{item.id}}"))
                .Message.Should().Be("Unknown identifier 'item'");
        }

        [Fact]
        public void EvalExpressionVariableWithAt()
        {
            var eval = new FormulaEvaluator();
            eval.AddVariable("@a", 1);
            eval.Evaluate("{{@a+@a}}").Should().Be(2);
        }

        class Customer
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }
    }
}

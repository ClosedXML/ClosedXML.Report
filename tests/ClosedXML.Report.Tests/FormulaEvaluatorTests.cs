﻿using ClosedXML.Report.Utils;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Linq.Dynamic.Core.CustomTypeProviders;
using System.Linq.Dynamic.Core.Exceptions;
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
            var lambda = XLDynamicExpressionParser.ParseLambda(new [] {Expression.Parameter(customers.GetType(), "customers")}, null, query);
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

        [Fact]
        public void ExpressionParseTestNullPropagation()
        {
            var customers = new Customer[]
            {
                new Customer {Id = 1, Name = "Customer1", Manager = new Customer { Id = 3, Name = "Manager1"}},
                new Customer {Id = 2, Name = "Customer2", Manager = null}
            };
            var eval = new FormulaEvaluator();
            eval.AddVariable("a", customers[0]);
            eval.AddVariable("b", customers[1]);
            eval.Evaluate(@"{{np(a.Manager.Name, ""test"")}}").Should().Be("Manager1");
            eval.Evaluate(@"{{np(b.Manager.Name, ""test"")}}").Should().Be("test");
            eval.Evaluate(@"{{np(b.Manager.Name, null)}}").Should().BeNull();
        }

        [Fact(Skip = "It doesn't work yet")]
        public void EvalDictionaryParams()
        {
            Parameter CreateDicParameter(string name) => new Parameter("item", new Dictionary<string, object>
                {{"Name", new Dictionary<string, object> {{"FirstName", name }}}});

            var eval = new FormulaEvaluator();
            eval.Evaluate("{{item.Name.FirstName}}", CreateDicParameter("Julio")).Should().Be("Julio");
            eval.Evaluate("{{item.Name.FirstName}}", CreateDicParameter("John")).Should().Be("John");
        }

        [Fact(Skip = "It doesn't work yet")]
        public void EvalDictionaryParams2()
        {
            object CreateDicParameter(string name) => new Dictionary<string, object>
                {{"Name", new Dictionary<string, object> {{"FirstName", name }}}};

            var config = new ParsingConfig()
            {
                CustomTypeProvider = new DefaultDynamicLinqCustomTypeProvider()
            };
            var parType = new Dictionary<string, object>().GetType();
            var lambda = DynamicExpressionParser.ParseLambda(config, new [] {Expression.Parameter(parType, "item")}, typeof(object), "item.Name.FirstName").Compile();
            lambda.DynamicInvoke(CreateDicParameter("Julio")).Should().Be("Julio");
            lambda.DynamicInvoke(CreateDicParameter("John")).Should().Be("John");
        }

        [Fact]
        public void UsingDynamicLinqTypeTest()
        {
            var eval = new FormulaEvaluator();
            eval.AddVariable("a", "1");
            eval.Evaluate("{{EvaluateUtils.ParseAsInt(a).IncrementMe()}}").Should().Be(2);
        }

        class Customer
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public Customer Manager { get; set; }
        }
    }

    [DynamicLinqType]
    public static class EvaluateUtils
    {
        public static int ParseAsInt(string value)
        {
            if (value == null)
            {
                return 0;
            }

            return int.Parse(value);
        }

        public static int IncrementMe(this int values)
        {
            return values + 1;
        }
    }
}

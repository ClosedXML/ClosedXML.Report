using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

namespace ClosedXML.Report
{
    internal class FormulaEvaluator
    {
        private static readonly Regex ExprMatch = new Regex(@"\{\{.+?\}\}");
        //  private readonly Interpreter _interpreter; !!! переделать на DynamicLinq
        private readonly Dictionary<string, Delegate> _lambdaCache = new Dictionary<string, Delegate>();
        private readonly Dictionary<string, object> _variables = new Dictionary<string, object>();

        public object Evaluate(string formula, params Parameter[] pars)
        {
            var expressions = GetExpressions(formula);
            foreach (var expr in expressions)
            {
                var val = Eval(expr.Substring(2, expr.Length - 4), pars);
                if (expr == formula)
                    return val;

                formula = formula.Replace(expr, ObjToString(val));
            }
            return formula;
        }

        public void AddVariable(string name, object value)
        {
            _variables.Add(name, value);
        }

        private string ObjToString(object val)
        {
            if (val is DateTime)
                return ((DateTime)val).ToOADate().ToString(CultureInfo.InvariantCulture);

            var formattable = val as IFormattable;
            return formattable != null ? formattable.ToString(null, CultureInfo.InvariantCulture) : val?.ToString();
        }

        private IEnumerable<string> GetExpressions(string cellValue)
        {
            var matches = ExprMatch.Matches(cellValue);
            return from Match match in matches select match.Value;
        }

        private object Eval(string expression, Parameter[] pars)
        {
            Delegate lambda;
            if (!_lambdaCache.TryGetValue(expression, out lambda))
            {
                var parameters = pars.Select(p=>p.ParameterExpression).ToArray();
                lambda = DynamicExpressionParser.ParseLambda(parameters, typeof(object), expression, _variables).Compile();

                _lambdaCache.Add(expression, lambda);
            }

            return lambda.DynamicInvoke(pars.Select(p => p.Value).ToArray());
        }
    }

    internal class Parameter
    {
        public Parameter(string name, object value)
        {
            ParameterExpression = Expression.Parameter(value.GetType(), name);
            Value = value;
        }

        public ParameterExpression ParameterExpression { get; private set; }
        public object Value { get; private set; }
    }
}

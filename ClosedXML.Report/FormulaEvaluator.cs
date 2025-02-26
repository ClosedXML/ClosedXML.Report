using ClosedXML.Report.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Dynamic.Core.Exceptions;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ClosedXML.Report
{
    public class FormulaEvaluator
    {
        private static readonly Regex ExprMatch = new Regex(@"\{\{.+?\}\}");

        private readonly Dictionary<string, Delegate> _lambdaCache = new Dictionary<string, Delegate>();
        private readonly Dictionary<string, object> _variables = new Dictionary<string, object>();

        public object Evaluate(string formula, params Parameter[] pars)
        {
            var expressions = GetExpressions(formula);
            foreach (var expr in expressions)
            {
                var val = Eval(Trim(expr), pars);
                if (expr == formula)
                    return val;

                formula = formula.Replace(expr, ObjToString(val));
            }
            return formula;
        }

        public bool TryEvaluate(string formula, out object result, params Parameter[] pars)
        {
            try
            {
                result = Evaluate(formula, pars);
                return true;
            }
            catch
            {
                result = null;
                return false;
            }
        }

        public void AddVariable(string name, object value)
        {
            if (value != null && !value.GetType().IsGenericType && value is IEnumerable enumerable)
            {
                var itemType = enumerable.GetItemType();
                var newEnumerable = EnumerableCastTo(enumerable, itemType);
                _variables[name] = newEnumerable;
            }
            else
                _variables[name] = value;
        }

        private IEnumerable EnumerableCastTo(IEnumerable enumerable, Type itemType)
        {
            ParameterExpression source = Expression.Parameter(typeof(IEnumerable));

            MethodInfo method = typeof(Enumerable).GetMethod(nameof(Enumerable.Cast), BindingFlags.Public | BindingFlags.Static);
            MethodInfo castGenericMethod = method.MakeGenericMethod(itemType);
            var castExpr = Expression.Call(null, castGenericMethod, source);

            var lambda = Expression.Lambda<Func<IEnumerable, IEnumerable>>(castExpr, source);

            return lambda.Compile().Invoke(enumerable);
        }

        private string ObjToString(object val)
        {
            if (val == null) val = "";
            if (val is DateTime dateVal)
                return dateVal.ToOADate().ToString(CultureInfo.InvariantCulture);

            return val is IFormattable formattable
                ? formattable.ToString(null, CultureInfo.InvariantCulture)
                : val.ToString();
        }

        private IEnumerable<string> GetExpressions(string cellValue)
        {
            var matches = ExprMatch.Matches(cellValue);
            if (matches.Count == 0)
                return new[] { cellValue };
            return from Match match in matches select match.Value;
        }

        private string Trim(string formula)
        {
            if (formula.StartsWith("{{"))
                return formula.Substring(2, formula.Length - 4);
            else
                return formula;
        }

        internal Delegate ParseExpression(string formula, ParameterExpression[] parameters)
        {
            var cacheKey = GetCacheKey(formula, parameters);
            if (!_lambdaCache.TryGetValue(cacheKey, out var lambda))
            {
                try
                {
                    lambda = XLDynamicExpressionParser.ParseLambda(parameters, typeof(object), formula, _variables).Compile();
                }
                catch (ArgumentException)
                {
                    return null;
                }

                _lambdaCache.Add(cacheKey, lambda);
            }
            return lambda;
        }

        private string GetCacheKey(string formula, ParameterExpression[] parameters)
        {
            return formula + string.Join("+", parameters.Select(x => x.Type.Name));
        }

        private object Eval(string expression, Parameter[] pars)
        {
            var parameters = pars.Select(p => p.ParameterExpression).ToArray();
            var lambda = ParseExpression(expression, parameters);
            return lambda?.DynamicInvoke(pars.Select(p => p.Value).ToArray());
        }
    }

    public class Parameter
    {
        public Parameter(string name, object value)
        {
            ParameterExpression = Expression.Parameter(value?.GetType() ?? typeof(object), name);
            Value = value;
        }

        public ParameterExpression ParameterExpression { get; }
        public object Value { get; }
    }
}

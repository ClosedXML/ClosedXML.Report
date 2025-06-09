using System;
using System.Collections.Generic;
using System.Linq.Dynamic.Core;
using System.Linq.Dynamic.Core.CustomTypeProviders;
using System.Linq.Expressions;
using System.Text;

namespace ClosedXML.Report.Utils
{
    internal static class XLDynamicExpressionParser
    {
        /// <summary>
        /// A wrapper for <see cref="DynamicExpressionParser.ParseLambda"/> providing custom parsing config.
        /// </summary>
        /// <param name="parameters">A array from ParameterExpressions.</param>
        /// <param name="resultType">Type of the result. If not specified, it will be generated dynamically.</param>
        /// <param name="expression">The expression.</param>
        /// <param name="values">An object array that contains zero or more objects which are used as replacement values.</param>
        /// <returns>The generated <see cref="T:System.Linq.Expressions.LambdaExpression" /></returns>
        public static LambdaExpression ParseLambda(
            ParameterExpression[] parameters,
            Type resultType,
            string expression,
            params object[] values)
        {
            var config = new ParsingConfig()
            {
                CustomTypeProvider = new XLDynamicLinqCustomTypeProvider()
            };

            return DynamicExpressionParser.ParseLambda(config, parameters, resultType, expression, values);
        }

        private  class XLDynamicLinqCustomTypeProvider : DefaultDynamicLinqCustomTypeProvider
        {
            private static HashSet<Type> _customTypesCache;
            private static readonly ParsingConfig _config = new ParsingConfig();

            public XLDynamicLinqCustomTypeProvider() : base(_config, cacheCustomTypes: true)
            {
            }

            public override HashSet<Type> GetCustomTypes()
            {
                if (_customTypesCache != null)
                    return _customTypesCache;

                _customTypesCache = base.GetCustomTypes();
                return _customTypesCache;
            }
        }
    }
}

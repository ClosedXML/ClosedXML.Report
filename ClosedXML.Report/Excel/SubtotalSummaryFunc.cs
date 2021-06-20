using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq.Expressions;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Excel
{
    public class SubtotalSummaryFunc
    {
        private static readonly Dictionary<string, IFuncData<IAggregator>> TotalFuncs = new Dictionary<string, IFuncData<IAggregator>>
        {
            {"average", new FuncData<AverageAggregator>(1)},
            {"avg", new FuncData<AverageAggregator>(1)},
            {"count", new FuncData<CountAggregator>(2)},
            {"countnums", new FuncData<CountAggregator>(2)},
            {"counta", new FuncData<CountAAggregator>(3)},
            {"max", new FuncData<MaxAggregator>(4)},
            {"min", new FuncData<MinAggregator>(5)},
            {"product", new FuncData<ProductAggregator>(6)},
            {"stdev", new FuncData<StDevAggregator>(7)},
            /*{"stdevp", new FuncData(8)},*/
            {"sum", new FuncData<SumAggregator>(9)},
            {"var", new FuncData<StDevAggregator>(10)},
            /*{"varp", new FuncData(11)}*/
        };

        private static IFuncData<IAggregator> GetFunc(string funcName)
        {
            var func = TotalFuncs.ContainsKey(funcName) ? TotalFuncs[funcName] : null;
            if (func == null)
                Debug.WriteLine("Unknown function " + funcName);
            return func;
        }

        private IFuncData<IAggregator> _func;
        internal IDataSource DataSource { get; set; }

        internal SubtotalSummaryFunc(string func, int column)
        {
            Column = column;
            FuncName = func.ToLower();
            _func = GetFunc(FuncName);
        }

        public string FuncName { get; private set; }
        public int Column { get; private set; }

        public virtual int FuncNum
        {
            get
            {
                if (_func == null)
                    _func = GetFunc(FuncName);
                return GetCalculateDelegate != null ? 0 : _func.FuncNum;
            }
        }

        public Func<Type, Delegate> GetCalculateDelegate;

        internal object Calculate(IDataSource dataSource)
        {
            var items = dataSource.GetAll();
            if (items == null || items.Length == 0)
                return null;

            if (FuncNum != 0)
                return null;

            var agg = _func.CreateAggregator();

            var dlg = GetCalculateDelegate(items[0].GetType());
            //var dlg = lambda.Compile();
            foreach (var item in items)
            {
                try
                {
                    var val = dlg.DynamicInvoke(item);

                    if (val != null)
                        agg.Aggregate(val);
                }
                catch { }
            }
            return agg.Result;
        }

        /*public object Calculate(IXLRange range)
        {
            if (FuncNum != 0)
                return null;

            if (GetExpression
            var agg = _func.CreateAggregator();

            var fval = DataSource.GetValue(range.FirstRow());
            var lambda = GetExpression(fval.GetType());
            using (var rows = range.Rows())
                foreach (var row in rows)
                {
                    try
                    {
                        var item = DataSource.GetValue(row);
                        var val = lambda.Invoke(item);
                        if (val != null)
                            agg.Aggregate(val);
                    }
                    catch { }
                }
            return agg.Result;
        }*/

        #region Private classes

        private interface IFuncData<out T> where T : IAggregator
        {
            T CreateAggregator();
            int FuncNum { get; }
        }

        private class FuncData<T> : IFuncData<T> where T : IAggregator
        {
            public FuncData(int func)
            {
                FuncNum = func;
            }

            public T CreateAggregator()
            {
                return Activator.CreateInstance<T>();
            }

            public int FuncNum { get; private set; }
        }

        private interface IAggregator
        {
            void Aggregate(object value);
            object Result { get; }
        }

        private class SumAggregator : IAggregator
        {
            public void Aggregate(object value)
            {
                if (Result == null)
                    Result = value.GetType().GetDefault();

                dynamic a = value;
                dynamic b = Result;
                Result = a + b;
            }

            public object Result { get; private set; }
        }

        private class ProductAggregator : IAggregator
        {
            public void Aggregate(object value)
            {
                if (Result == null)
                    Result = value.GetType().GetDefault();

                dynamic a = value;
                dynamic b = Result;
                Result = a * b;
            }

            public object Result { get; private set; }
        }

        private class CountAggregator : IAggregator
        {
            private int _cnt = 0;
            public void Aggregate(object value)
            {
                if (value.GetType().IsNumeric())
                    _cnt++;
            }

            public object Result { get { return _cnt; } }
        }

        private class CountAAggregator : IAggregator
        {
            private int _cnt = 0;
            public void Aggregate(object value)
            {
                _cnt++;
            }

            public object Result { get { return _cnt; } }
        }

        private class MinAggregator : IAggregator
        {
            private object _min;
            public void Aggregate(object value)
            {
                if (_min == null)
                    _min = value;

                dynamic a = value;
                dynamic b = _min;
                if (a < b)
                    _min = value;
            }

            public object Result { get { return _min; } }
        }

        private class MaxAggregator : IAggregator
        {
            private object _max;
            public void Aggregate(object value)
            {
                if (_max == null)
                    _max = value;

                dynamic a = value;
                dynamic b = _max;
                if (a > b)
                    _max = value;
            }

            public object Result { get { return _max; } }
        }

        private class AverageAggregator : IAggregator
        {
            protected readonly List<object> List = new List<object>();

            public void Aggregate(object value)
            {
                List.Add(value);
            }

            public virtual object Result
            {
                get
                {
                    if (List.Count == 0)
                        return null;

                    dynamic sum = List[0].GetType().GetDefault();
                    foreach (dynamic v in List)
                        sum += v;

                    if (sum is TimeSpan)
                    {
                        return TimeSpan.FromTicks(((TimeSpan)sum).Ticks / List.Count);
                    }

                    return sum / List.Count;
                }
            }
        }

        // выборочная дисперсия
        private class VarAggregator : AverageAggregator
        {
            public override object Result
            {
                get
                {
                    dynamic avg = base.Result;

                    dynamic sum = List[0].GetType().GetDefault();
                    foreach (dynamic v in List)
                    {
                        sum += Math.Pow(v - avg, 2);
                    }

                    return sum / (List.Count - 1);
                }
            }
        }

        // Стандартное отклонение
        private class StDevAggregator : VarAggregator
        {
            public override object Result
            {
                get
                {
                    dynamic stdev = base.Result;
                    return Math.Sqrt(stdev);
                }
            }
        }

        #endregion
    }
}

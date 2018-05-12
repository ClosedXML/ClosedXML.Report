using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using ClosedXML.Report.Utils;
//using DynamicExpresso;

namespace ClosedXML.Report
{
    internal class RangeInterpreter
    {
        private readonly string _alias;
        private readonly FormulaEvaluator _evaluator;
        private readonly TagsEvaluator _tagsEvaluator;
        private readonly Dictionary<string, object> _variables = new Dictionary<string, object>();
        private readonly Dictionary<string, TagsList> _tags = new Dictionary<string, TagsList>();

        public RangeInterpreter(string alias)
        {
            _alias = alias;
            _evaluator = new FormulaEvaluator();
            _tagsEvaluator = new TagsEvaluator();
        }

        public void Evaluate(IXLRange range)
        {
            var rangeName = range.RangeAddress.ToStringRelative(true);
            ParseTags(range, rangeName);
            EvaluateValues(range);
            TagsPostprocessing(rangeName, null);
        }

        public void ParseTags(IXLRange range, string rangeName)
        {
            var innerRanges = range.GetContainingNames().Where(nr => _variables.ContainsKey(nr.Name)).ToArray();
            var cells = from c in range.CellsUsed()
                        let value = c.GetString()
                        where !c.HasFormula
                            && (value.StartsWith("<<") || value.EndsWith(">>"))
                            && !innerRanges.Any(nr => { using (var r = nr.Ranges) using (var cr = c.AsRange()) return r.Contains(cr);})
                        select c;

            if (!_tags.ContainsKey(rangeName))
                _tags.Add(rangeName, new TagsList());

            foreach (var cell in cells)
            {
                string value = cell.GetString();
                OptionTag[] tags;
                string newValue;
                var templateCell = new TemplateCell(cell.Address.RowNumber, cell.Address.ColumnNumber, cell);
                if (value.StartsWith("&="))
                {
                    tags = _tagsEvaluator.Parse(value.Substring(2), range, templateCell, out newValue);
                    cell.FormulaA1 = newValue;
                }
                else
                {
                    tags = _tagsEvaluator.Parse(value, range, templateCell, out newValue);
                    cell.Value = newValue;
                }
                _tags[rangeName].AddRange(tags);
            }
        }

        public void TagsPostprocessing(string rangeName, ProcessingContext context)
        {
            if (_tags.ContainsKey(rangeName))
            {
                var tags = _tags[rangeName];
                tags.Execute(context);
            }
        }

        public void CopyTags(string srcRangeName, string destRangeName, IXLRange destRange)
        {
            var srcTags = _tags[srcRangeName];
            if (!_tags.ContainsKey(destRangeName))
                _tags.Add(destRangeName, new TagsList());
            _tags[destRangeName].AddRange(srcTags.CopyTo(destRange));
        }

        public virtual void EvaluateValues(IXLRange range, params Parameter[] pars)
        {
            foreach (var parameter in pars)
            {
                AddParameter(parameter.Value);
            }
            range.Worksheet.SuspendEvents();
            var innerRanges = range.GetContainingNames().Where(nr => _variables.ContainsKey(nr.Name)).ToArray();
            var cells = range.CellsUsed()
                .Where(c => !c.HasFormula
                            && c.GetString().Contains("{{")
                            && !innerRanges.Any(nr => nr.Ranges.Contains(c.AsRange())))
                .ToArray();
            range.Worksheet.ResumeEvents();

            foreach (var cell in cells)
            {
                string value = cell.GetString();
                try
                {
                    if (value.StartsWith("&="))
                        cell.FormulaA1 = _evaluator.Evaluate(value.Substring(2), pars).ToString();
                    else
                        cell.Value = _evaluator.Evaluate(value, pars);
                }
                catch (ParseException ex)
                {
                    Debug.WriteLine("Cell value evaluation exception (range '{1}'): {0}", ex.Message, range.RangeAddress);
                }
            }

            foreach (var nr in innerRanges)
            {
                if (!_variables.ContainsKey(nr.Name))
                {
                    Debug.WriteLine(string.Format("Range {0} was skipped. Variable with that name was not found.", nr.Name));
                    continue;
                }

                var datas = _variables[nr.Name] as IEnumerable;
                if (datas == null)
                    continue;

                var items = datas as object[] ?? datas.Cast<object>().ToArray();
                var tplt = RangeTemplate.Parse(nr);
                var nrng = nr.Ranges.First();
                using (var buff = tplt.Generate(items))
                {
                    var trgtRng = buff.CopyTo(nrng);
                    nr.SetRefersTo(trgtRng);
                    
                    //Apparently, this is needed for Pivot tags only
                    //tplt.RangeTagsApply(trgtRng, items);
                }

                // refresh ranges for pivot tables
                foreach (var pt in range.Worksheet.Workbook.Worksheets.SelectMany(sh => sh.PivotTables))
                {
                    if (pt.SourceRange.Intersects(nrng))
                    {
                        pt.SourceRange = nrng.Offset(-1, 1, nrng.RowCount(), nrng.ColumnCount() - 1);
                    }
                }
            }
        }

        private void AddParameter(object value)
        {
            var type = value.GetType();
            if (type.IsPrimitive())
                return;

            var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance).Where(f => f.IsPublic)
                .Select(f => new { f.Name, val = f.GetValue(value), type = f.FieldType })
                .Concat(type.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(f => f.CanRead)
                    .Select(f => new { f.Name, val = f.GetValue(value), type = f.PropertyType }));

            string alias = _alias;
            if (!string.IsNullOrEmpty(alias))
                alias = alias + "_";

            foreach (var field in fields)
            {
                _variables[alias + field.Name] = field.val;
            }
        }

        public void AddVariable(string alias, object value)
        {
            _variables.Add(alias, value);
            _evaluator.AddVariable(alias, value);
        }
    }
}
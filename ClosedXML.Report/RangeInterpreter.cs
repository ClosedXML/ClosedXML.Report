using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using ClosedXML.Report.Utils;
using System.Linq.Dynamic.Core.Exceptions;


namespace ClosedXML.Report
{
    internal class RangeInterpreter
    {
        private readonly string _alias;
        private readonly FormulaEvaluator _evaluator;
        private readonly TagsEvaluator _tagsEvaluator;
        private readonly Dictionary<string, object> _variables = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, TagsList> _tags = new Dictionary<string, TagsList>();
        private readonly TemplateErrors _errors;

        public RangeInterpreter(string alias, TemplateErrors errors)
        {
            _alias = alias;
            _errors = errors;
            _evaluator = new FormulaEvaluator();
            _tagsEvaluator = new TagsEvaluator();
        }

        public void Evaluate(IXLRange range)
        {
            var rangeName = range.RangeAddress.ToStringRelative(true);
            ParseTags(range, rangeName);
            EvaluateValues(range);
            TagsPostprocessing(rangeName, new ProcessingContext(range, null, _evaluator));
        }

        public void ParseTags(IXLRange range, string rangeName)
        {
            var innerRanges = range.GetContainingNames().Where(nr => _variables.ContainsKey(nr.Name)).ToArray();
            var cellsUsed = range.CellsUsed()
                .Where(c => !c.HasFormula && !innerRanges.Any(nr => nr.Ranges.Contains(c.AsRange())))
                .ToArray();
            var cells = from c in cellsUsed
                let value = c.GetString()
                where (value.StartsWith("<<") || value.EndsWith(">>"))
                select c;

            if (!_tags.ContainsKey(rangeName))
                _tags.Add(rangeName, new TagsList(_errors));

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
                _tags.Add(destRangeName, new TagsList(_errors));
            _tags[destRangeName].AddRange(srcTags.CopyTo(destRange));
        }

        public virtual void EvaluateValues(IXLRange range, params Parameter[] pars)
        {
            foreach (var parameter in pars)
            {
                AddParameter(parameter.Value);
            }
            var innerRanges = range.GetContainingNames()
                .Select(BindToVariable)
                .Where(nr => nr != null)
                .ToArray();

            var cells = range.CellsUsed()
                .Where(c => !c.HasFormula
                            && c.GetString().Contains("{{")
                            && !innerRanges.Any(nr => nr.NamedRange.Ranges.Contains(c.AsRange())))
                .ToArray();

            foreach (var cell in cells)
            {
                string value = cell.GetString();
                try
                {
                    if (value.StartsWith("&="))
                        cell.FormulaA1 = _evaluator.Evaluate(value.Substring(2), pars).ToString();
                    else
                        cell.SetValue(_evaluator.Evaluate(value, pars));
                }
                catch (ParseException ex)
                {
                    if (ex.Message == "Unknown identifier 'item'" && pars.Length == 0)
                    {
                        var firstCell = cell.Address.RowNumber > 1
                            ? cell.CellAbove().WorksheetRow().FirstCell()
                            : cell.WorksheetRow().FirstCell();
                        var msg = "The range does not meet the requirements of the list ranges. For details, see the documentation.";
                        firstCell.Value = msg;
                        firstCell.Style.Font.FontColor = XLColor.Red;
                        _errors.Add(new TemplateError(msg, firstCell.AsRange()));
                    }
                    cell.Value = ex.Message;
                    cell.Style.Font.FontColor = XLColor.Red;
                    _errors.Add(new TemplateError(ex.Message, cell.AsRange()));
                }

                string EvalString(string str)
                {
                    try
                    {
                        return _evaluator.Evaluate(str, pars).ToString();
                    }
                    catch (ParseException ex)
                    {
                        _errors.Add(new TemplateError(ex.Message, cell.AsRange()));
                        return ex.Message;
                    }
                }

                if (cell.HasComment)
                {
                    var comment = EvalString(cell.Comment.Text);
                    cell.Comment.ClearText();
                    cell.Comment.AddText(comment);
                }

                if (cell.HasHyperlink)
                {
                    if (cell.Hyperlink.IsExternal)
                        cell.Hyperlink.ExternalAddress = new Uri(EvalString(cell.Hyperlink.ExternalAddress.ToString()));
                    else
                        cell.Hyperlink.InternalAddress = EvalString(cell.Hyperlink.InternalAddress);
                }

                if (cell.HasRichText)
                {
                    var richText = EvalString(cell.RichText.Text);
                    cell.RichText.ClearText();
                    cell.RichText.AddText(richText);
                }
            }

            foreach (var nr in innerRanges)
            {
                foreach (var rng in nr.NamedRange.Ranges)
                {
                    var items = nr.RangeData as object[] ?? nr.RangeData.Cast<object>().ToArray();
                    var tplt = RangeTemplate.Parse(nr.NamedRange.Name, rng, _errors, _variables);
                    using (var buff = tplt.Generate(items))
                    {
                        var ranges = nr.NamedRange.Ranges;
                        var trgtRng = buff.CopyTo(rng);
                        ranges.Remove(rng);
                        ranges.Add(trgtRng);
                        nr.NamedRange.SetRefersTo(ranges);

                        tplt.RangeTagsApply(trgtRng, items);
                    }

                    // refresh ranges for pivot tables
                    foreach (var pt in range.Worksheet.Workbook.Worksheets.SelectMany(sh => sh.PivotTables))
                    {
                        if (pt.SourceRange.Intersects(rng))
                        {
                            pt.SourceRange = rng.Offset(-1, 1, rng.RowCount() + 1, rng.ColumnCount() - 1);
                        }
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
                    .Select(f => new { f.Name, val = f.GetValue(value, new object[] { }), type = f.PropertyType }));

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

        private BoundRange BindToVariable(IXLNamedRange namedRange)
        {
            if (_variables.TryGetValue(namedRange.Name, out var variableValue) &&
                variableValue is IEnumerable data1)
                return new BoundRange(namedRange, data1);

            var expression = "{{" + namedRange.Name.Replace("_", ".") +"}}";

            if (_evaluator.TryEvaluate(expression, out var res) &&
                res is IEnumerable data2)
                return new BoundRange(namedRange, data2);

            return null;
        }

        private class BoundRange
        {
            public IXLNamedRange NamedRange { get; }

            public IEnumerable RangeData { get; }

            public BoundRange(IXLNamedRange namedRange, IEnumerable rangeData)
            {
                NamedRange = namedRange;
                RangeData = rangeData;
            }
        }
    }
}

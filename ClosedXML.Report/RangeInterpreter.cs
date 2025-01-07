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
            EvaluateValues(range);
            ParseTags(range, rangeName);
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
                        where TagExtensions.HasTag(value)
                        select c;

            if (!_tags.ContainsKey(rangeName))
                _tags.Add(rangeName, new TagsList(_errors));

            foreach (var cell in cells)
            {
                OptionTag[] tags = _tagsEvaluator.ApplyTagsTo(cell, range);
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

        /// <summary>
        /// <para>
        /// Apply variables to the template ranges and template cells in the <paramref name="range"/>.
        /// </para>
        /// </summary>
        public virtual void EvaluateValues(IXLRange range, params Parameter[] pars)
        {
            foreach (var parameter in pars)
            {
                AddParameter(parameter.Value);
            }

            // Get all defined names in the `range` that with the data from variables
            var boundRanges = new List<BoundRange>();
            foreach (var candidateName in range.GetContainingNames())
            {
                if (TryBindToVariable(candidateName, out var boundRange))
                    boundRanges.Add(boundRange);
            }

            // Get cells that should be templated, but aren't part of a bounded range.
            var cellsToTemplate = range.CellsUsed()
                .Where(c => !c.HasFormula
                            && c.GetString().Contains("{{")
                            && !boundRanges.Any(nr => nr.DefinedName.Ranges.Contains(c.AsRange())))
                .ToArray();

            // Apply template to the cell content, i.e. value, rich text, formula, comment or hyperlink.
            // Unlike bound ranges, this doesn't change position of a cell, so it should be done first.
            foreach (var cell in cellsToTemplate)
            {
                string value = cell.GetString();
                try
                {
                    if (value.StartsWith("&="))
                        cell.FormulaA1 = _evaluator.Evaluate(value.Substring(2), pars).ToString();
                    else
                    {
                        var cellValue = XLCellValueConverter.FromObject(_evaluator.Evaluate(value, pars));
                        cell.SetValue(cellValue);
                    }
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
                    var comment = EvalString(cell.GetComment().Text);
                    cell.GetComment().ClearText();
                    cell.GetComment().AddText(comment);
                }

                if (cell.HasHyperlink)
                {
                    if (cell.GetHyperlink().IsExternal)
                        cell.GetHyperlink().ExternalAddress = new Uri(EvalString(cell.GetHyperlink().ExternalAddress.ToString()));
                    else
                        cell.GetHyperlink().InternalAddress = EvalString(cell.GetHyperlink().InternalAddress);
                }

                if (cell.HasRichText)
                {
                    var richText = EvalString(cell.GetRichText().Text);
                    cell.GetRichText().ClearText();
                    cell.GetRichText().AddText(richText);
                }
            }

            // Render bound ranges
            foreach (var nr in boundRanges)
            {
                foreach (var rng in nr.DefinedName.Ranges)
                {
                    var grownRange = rng.GrowToMergedRanges();
                    var items = nr.RangeData as object[] ?? nr.RangeData.Cast<object>().ToArray();

                    if (!items.Any() && grownRange.IsOptionsRowEmpty())
                    {
                        // Related to #251. I am pretty sure this is wrong solution and dealing with empty items
                        // should be done through RangeTemplate below. But if there are no items and empty option
                        // row, the result is degenerated A1 rendered range in temp sheet. Deleting (empty) options
                        // row would thus delete only first cell, not full (empty) options row.
                        grownRange.Delete(XLShiftDeletedCells.ShiftCellsUp);
                        continue;
                    }

                    // Range template generates output into a new temporary sheet, as not to affect other things
                    // and then copies it to the range in the original sheet.
                    var rangeTemplate = RangeTemplate.Parse(nr.DefinedName.Name, grownRange, _errors, _variables);
                    using (var renderedBuffer = rangeTemplate.Generate(items))
                    {
                        var ranges = nr.DefinedName.Ranges;
                        var trgtRng = renderedBuffer.CopyTo(grownRange);
                        ranges.Remove(rng);
                        ranges.Add(trgtRng);
                        nr.DefinedName.SetRefersTo(ranges);

                        rangeTemplate.RangeTagsApply(trgtRng, items);
                        var isOptionsRowEmpty = trgtRng.IsOptionsRowEmpty();
                        if (isOptionsRowEmpty)
                            trgtRng.LastRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
                    }

                    // refresh ranges for pivot tables
                    foreach (var pivotCache in range.Worksheet.Workbook.PivotCaches)
                    {
                        pivotCache.Refresh();
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

        private bool TryBindToVariable(IXLDefinedName variableName, out BoundRange boundRange)
        {
            if (_variables.TryGetValue(variableName.Name, out var variableValue) &&
                variableValue is IEnumerable data1)
            {
                boundRange = new BoundRange(variableName, data1);
                return true;
            }

            var expression = "{{" + variableName.Name.Replace("_", ".") + "}}";

            if (_evaluator.TryEvaluate(expression, out var res) &&
                res is IEnumerable data2)
            {
                boundRange = new BoundRange(variableName, data2);
                return true;
            }

            boundRange = null;
            return false;
        }

        [DebuggerDisplay("Bound variable: {DefinedName.Name}")]
        private class BoundRange
        {
            public IXLDefinedName DefinedName { get; }

            public IEnumerable RangeData { get; }

            public BoundRange(IXLDefinedName definedName, IEnumerable rangeData)
            {
                DefinedName = definedName;
                RangeData = rangeData;
            }
        }
    }
}

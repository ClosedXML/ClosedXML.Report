using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    internal class TagsEvaluator
    {
        private static readonly Regex TagsMatch = new Regex(@"\<\<.+?\>\>");

        private IEnumerable<string> GetAllTags(string cellValue)
        {
            var matches = TagsMatch.Matches(cellValue);
            return from Match match in matches select match.Value;
        }

        /// <summary>
        /// Apply tags to the <paramref name="cell"/>, if it contains tags.
        /// </summary>
        /// <param name="cell">Template cell that *might* contain tags (doesn't have to).</param>
        /// <param name="range">Range each option will be associated with.</param>
        /// <returns>Created tags for the cell (may be empty).</returns>
        public OptionTag[] ApplyTagsTo(TemplateCell cell, IXLRange range)
        {
            OptionTag[] tags;
            if (cell.CellType == TemplateCellType.Formula)
            {
                tags = Parse(cell.Formula, range, cell, out var newValue);
                cell.Formula = newValue;
            }
            else
            {
                tags = Parse(cell.GetString(), range, cell, out var newValue);
                cell.Value = newValue == "" ? Blank.Value : newValue;
            }

            return tags;
        }

        public OptionTag[] Parse(string templateLiteral, IXLRange range, TemplateCell cell, out string newValue)
        {
            List<OptionTag> result = new List<OptionTag>();
            foreach (var expr in GetAllTags(templateLiteral))
            {
                var optionTag = ParseTag(expr.Substring(2, expr.Length-4));
                if (optionTag == null)
                    continue;
                optionTag.Cell = cell;
                optionTag.Range = range;
                if (cell.XLCell.Address.RowNumber != cell.Row) // is range tag
                {
                    optionTag.RangeOptionsRow = range.LastRow().RangeAddress;
                }
                result.Add(optionTag);
                templateLiteral = templateLiteral.Replace(expr, "");
            }

            newValue = templateLiteral.Trim();
            return result.ToArray();
        }

        private OptionTag ParseTag(string str)
        {
            string name;
            Dictionary<string, string> dictionary;
            using (var reader = new VernoStringReader(str))
            {
                name = reader.ReadWord();
            
                dictionary = new Dictionary<string, string>();
                foreach (var pair in reader.ReadNamedValues(" ", "="))
                {
                    dictionary.Add(pair.Key.ToLower(), pair.Value);
                }
            }

            return TagsRegister.CreateOption(name, dictionary);

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Report.Options
{
    public class TagsList : SortedSet<OptionTag>
    {
        private readonly TemplateErrors _errors;

        public TagsList(TemplateErrors errors) : base(new OptionTagComparer())
        {
            _errors = errors;
        }

        public TagsList CopyTo(IXLRange toRange)
        {
            var clone = new TagsList(_errors);
            foreach (var srcTag in this)
            {
                var tag = (OptionTag)srcTag.Clone();
                //var cell = toRange.Cell(tag.Cell.AsRange().Relative(tag.Range).FirstAddress);
                tag.Cell = srcTag.Cell.Clone();
                tag.Range = toRange;
                clone.Add(tag);
            }
            return clone;
        }

        public new void Add(OptionTag tag)
        {
            tag.List = this;
            base.Add(tag);
        }

        public void AddRange(IEnumerable<OptionTag> collection)
        {
            foreach (var tag in collection)
            {
                Add(tag);
            }
        }

        public IEnumerable<OptionTag> GetAll(string[] names)
        {
            return this.Where(x => Array.IndexOf(names, x.Name) >= 0);
        }

        public IEnumerable<T> GetAll<T>() where T : OptionTag
        {
            return this.OfType<T>().Where(x => x.Enabled);
        }

        public IEnumerable<OptionTag> GetAll(OptionTag exclude, string[] names)
        {
            return this.Where(x => x != exclude && Array.IndexOf(names, x.Name) >= 0);
        }

        public void Execute(ProcessingContext context)
        {
            while (true)
            {
                var t = this.FirstOrDefault(x => x.Enabled);
                if (t == null)
                    break;

                try
                {
                    t.Execute(context);
                }
                catch (TemplateParseException ex)
                {
                    _errors.Add(new TemplateError(ex.Message, ex.Range));
                }
                finally
                {
                    t.Enabled = false;
                }
            }
        }

        public bool HasTag(string name)
        {
            return this.Any(x => string.Equals(x.Name, name, StringComparison.InvariantCultureIgnoreCase));
        }

        public void Reset()
        {
            foreach (var item in this)
                item.Enabled = true;
        }

        internal class OptionTagComparer : IComparer<OptionTag>
        {
            public int Compare(OptionTag x, OptionTag y)
            {
                var result = -x.Priority.CompareTo(y.Priority);

                if (x.Cell != null)
                {
                    if (result == 0)
                        result = x.Cell.Row.CompareTo(y.Cell.Row);
                    if (result == 0)
                        result = x.Cell.Column.CompareTo(y.Cell.Column);
                }

                if (result == 0)
                    return 1;

                return result;
            }
        }
    }
}

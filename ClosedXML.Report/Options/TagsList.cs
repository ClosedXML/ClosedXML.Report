using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Report.Options
{
    public class TagsList : SortedSet<OptionTag>
    {
        public TagsList() : base(new OptionTagComparer())
        {
        }

        public TagsList CopyTo(IXLRange toRange)
        {
            var clone = new TagsList();
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
            var map = Enumerable.Range(byte.MinValue, byte.MaxValue + 1).ToList();
            map.RemoveAll(x => this.Any(t => t.PriorityKey == x));
            tag.PriorityKey = (byte)map.OrderBy(x => Math.Abs(tag.Priority - x)).First();
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
                var t = this.FirstOrDefault(x=>x.Enabled);
                if (t == null)
                    break;

                try
                {
                    t.Execute(context);
                }
                catch
                {
                    throw;
                    // TODO ignored
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
    }

    internal class OptionTagComparer : IComparer<OptionTag>
    {
        public int Compare(OptionTag x, OptionTag y)
        {
            return -x.PriorityKey.CompareTo(y.PriorityKey);
        }
    }
}

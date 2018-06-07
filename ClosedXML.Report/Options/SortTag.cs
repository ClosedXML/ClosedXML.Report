/*
 -----------------------------------------------------------------------
"Sort"          "\Desc"               Column       rt      Normal      
                "\num=n"
"Desc"                                Column       rt      Normal
"Asc"                                 Column       rt      Normal
-----------------------------------------------------------------------
 */
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    public class SortTag : OptionTag
    {
        public XLSortOrder Order
        {
            get { return Parameters.ContainsKey("desc") ? XLSortOrder.Descending : XLSortOrder.Ascending; }
        }

        public int Num
        {
            get { return Parameters.ContainsKey("num") ? Parameters["num"].AsInt(1) : int.MaxValue; }
        }

        public override void Execute(ProcessingContext context)
        {
            var fields = List.GetAll<SortTag>().ToArray();
            foreach (var tag in fields.OrderBy(x => x.Num).ThenBy(x => x.Column))
            {
                context.Range.SortColumns.Add(tag.Column, tag.Order);
            }
            context.Range.Sort();

            foreach (var tag in fields)
            {
                tag.Enabled = false;
            }
        }

        public override byte Priority => 128;
    }

    public class DescTag : SortTag
    {
        public override void Execute(ProcessingContext context)
        {
            Parameters["desc"] = null;
            base.Execute(context);
        }
    }
}

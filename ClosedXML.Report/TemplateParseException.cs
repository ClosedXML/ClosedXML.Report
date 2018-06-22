using System;
using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public class TemplateParseException: Exception
    {
        public TemplateErrors InnerErrors { get; }
        public IXLRange Range { get; }

        public TemplateParseException(string message, IXLRange range) : base(message)
        {
            Range = range;
        }

        public TemplateParseException(string message, TemplateErrors errors)
        {
            InnerErrors = errors;
        }
    }
}

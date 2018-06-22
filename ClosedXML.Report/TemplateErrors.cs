using System.Collections;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public class TemplateErrors: IEnumerable<TemplateError>
    {
        private readonly List<TemplateError> _errors = new List<TemplateError>();

        public IEnumerator<TemplateError> GetEnumerator()
        {
            return _errors.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public TemplateError this[int index]
        {
            get => _errors[index];
        }

        public int Count { get => _errors.Count; }

        internal void Add(TemplateError error)
        {
            if (!_errors.Exists(x => x.Range.Equals(error.Range) && x.Message == error.Message))
                _errors.Add(error);
        }
    }

    public class TemplateError
    {
        public TemplateError(string message, IXLRange range)
        {
            Message = message;
            Range = range;
        }

        public string Message { get; }
        public IXLRange Range { get; }
    }
}

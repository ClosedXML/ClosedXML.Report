using System;
using ClosedXML.Report.Tests.TestModels;

namespace ClosedXML.Report.Tests.Utils
{
    public class DatabaseFixture : IDisposable
    {
        public DatabaseFixture()
        {
            DbDemos.DefaultSettings = new DbSettings();
        }

        public void Dispose()
        {
        }
    }
}

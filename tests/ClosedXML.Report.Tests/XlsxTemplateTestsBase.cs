using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using ClosedXML.Report.Tests.Utils;
using FluentAssertions;
//using JetBrains.Profiler.Windows.Api;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class XlsxTemplateTestsBase
    {
        private readonly ITestOutputHelper _output;
        public XlsxTemplateTestsBase(ITestOutputHelper output)
        {
            _output = output;
            LinqToDB.Common.Configuration.Linq.AllowMultipleQuery = true;
        }

        // Because different fonts are installed on Unix,
        // the columns widths after AdjustToContents() will
        // cause the tests to fail.
        // Therefore we ignore the width attribute when running on Unix
        public static bool IsRunningOnUnix
        {
            get
            {
                int p = (int)Environment.OSVersion.Platform;
                return ((p == 4) || (p == 6) || (p == 128));
            }
        }

        protected void XlTemplateTest(string tmplFileName, Action<XLTemplate> arrangeCallback, Action<XLWorkbook> assertCallback)
        {
            /*if (MemoryProfiler.IsActive && MemoryProfiler.CanControlAllocations)
                MemoryProfiler.EnableAllocations();*/

            //MemoryProfiler.Dump();

            var fileName = Path.Combine(TestConstants.TemplatesFolder, tmplFileName);
            var workbook = new XLWorkbook(fileName);
            var template = new XLTemplate(workbook);

            // ARRANGE
            arrangeCallback(template);

            using (var file = new MemoryStream())
            {
                //MemoryProfiler.Dump();
                // ACT
                var start = DateTime.Now;
                template.Generate();
                _output.WriteLine(DateTime.Now.Subtract(start).ToString());
                //MemoryProfiler.Dump();
                workbook.SaveAs(file);
                //MemoryProfiler.Dump();
                file.Position = 0;

                using (var wb = new XLWorkbook(file))
                {
                    // ASSERT
                    assertCallback(wb);
                }
            }
            workbook.Dispose();
            workbook = null;
            template = null;
            GC.Collect();
            //MemoryProfiler.Dump();
        }

        protected void CompareWithGauge(Stream streamActual, string fileExpected)
        {
            fileExpected = Path.Combine(TestConstants.GaugesFolder, fileExpected);
            using (var streamExpected = File.OpenRead(fileExpected))
            {
                string message;
                var success = ExcelDocsComparer.Compare(streamActual, streamExpected, IsRunningOnUnix, out message);
                var formattedMessage =
                    String.Format(
                        "Found difference from the expected file '{0}'. The difference is: '{1}'",
                        fileExpected, message);

                success.Should().BeTrue(formattedMessage);
            }
        }

        protected void CompareWithGauge(XLWorkbook workbook, string fileExpected)
        {
            using (var stream = new MemoryStream())
            {
                workbook.SaveAs(stream, true);
                stream.Position = 0;
                CompareWithGauge(stream, fileExpected);
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class PivotTests : XlsxTemplateTestsBase
    {
        public PivotTests(ITestOutputHelper output) : base(output)
        {
        }

        [Theory,
         InlineData("tPivot1.xlsx"),
         InlineData("tPivot5_Static.xlsx")]
        public void Simple(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var rows = from o in db.orders
                            select new {o.Customer.Company, o.PaymentMethod, OrderNo = o.OrderNo.ToString(), o.ShipDate, o.ItemsTotal, o.TaxRate, o.AmountPaid};
                        tpl.AddVariable("Orders", rows);
                    }
                },
                wb =>
                {
                    CompareWithGauge(wb, templateFile);
                });
        }

        [Fact(Skip = "Pivot support is limited")]
        public void CreatePivot()
        {
            using (var db = new DbDemos())
            {
                var rows = from o in db.orders
                    select new {o.Customer.Company, o.PaymentMethod, o.OrderNo, o.ShipDate, o.ItemsTotal, o.TaxRate, o.AmountPaid};

                using (var workbook = new XLWorkbook())
                {
                    var sheet = workbook.Worksheets.Add("PastrySalesData");

                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var source = sheet.Cell(1, 1).InsertTable(rows, "PastrySalesData", true);

                    // Create a range that includes our table, including the header row
                    var range = source.DataRange;
                    var header = sheet.Range(1, 1, 1, 3);
                    var dataRange = sheet.Range(header.FirstCell(), range.LastCell());

                    // Add a new sheet for our pivot table
                    var ptSheet = workbook.Worksheets.Add("PivotTable");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    var pt = ptSheet.PivotTables.Add("PivotTable", ptSheet.Cell(8, 2), dataRange);
                    pt.MergeAndCenterWithLabels = true;
                    pt.ShowExpandCollapseButtons = false;
                    pt.ClassicPivotTableLayout = true;
                    pt.ShowGrandTotalsColumns = false;
                    pt.SortFieldsAtoZ = true;

                    var pf = pt.RowLabels.Add("PaymentMethod");
                    pf.AddSubtotal(XLSubtotalFunction.Sum);
                    pf.AddSubtotal(XLSubtotalFunction.Average);
                    pt.RowLabels.Add("OrderNo");
                    pt.RowLabels.Add("ShipDate");

                    // The rows in our pivot table will be the names of the pastries
                    /*pt.RowLabels.Add("Company");
                    pt.RowLabels.Add("PaymentMethod", "Payment Method");
                    pt.RowLabels.Add("OrderNo");*/


                    // The columns will be the months
                    pt.ColumnLabels.Add("TaxRate");

                    // The values in our table will come from the "NumberOfOrders" field
                    // The default calculation setting is a total of each row/column
                    pt.Values.Add("AmountPaid", "Amount paid");
                    pt.Values.Add("ItemsTotal", "Items Total");

                    workbook.SaveAs("pivot_example.xlsx");
                }
                using (var wb = XLWorkbook.OpenFromTemplate("pivot_example.xlsx"))
                {
                    wb.SaveAs("pivot_example1.xlsx");
                }
            }
        }
    }
}

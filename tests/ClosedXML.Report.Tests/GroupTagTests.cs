﻿using System.Drawing;
using System.Linq;
using ClosedXML.Report.Tests.TestModels;
using LinqToDB;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class GroupTagTests : XlsxTemplateTestsBase
    {
        public GroupTagTests(ITestOutputHelper output) : base(output)
        {
        }

        [Theory,
         InlineData("GroupTagTests_Simple.xlsx"),
         InlineData("GroupTagTests_Simple_WithOutsideLink.xlsx"),
         InlineData("GroupTagTests_Collapse.xlsx"),
         InlineData("tLists1_sort.xlsx"),
         InlineData("tLists1_cell_setting.xlsx"),
         InlineData("tLists2_sum.xlsx"),
         InlineData("tLists3_options.xlsx"),
         InlineData("issue#111_autofilter_with_delete.xlsx"),
         InlineData("tLists4_complexRange.xlsx"),
         InlineData("tLists5_GlobalVars.xlsx"),
         InlineData("tLists6_count.xlsx"),
         InlineData("tLists7_image.xlsx"),
         InlineData("tPage1_options.xlsx"),
        ]
        public void Simple(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var cust = db.customers.LoadWith(x => x.Orders).OrderBy(c => c.CustNo).First(x=>x.CustNo == 1356);
                        cust.Logo = Resource.toms_diving_center;
                        cust.Orders.ForEach(x =>
                        {
                            switch (x.PaymentMethod)
                            {
                                case "Visa": x.PaymentImage = Resource.card; break;
                                case "Cash": x.PaymentImage = Resource.cash; break;
                                case "Credit": x.PaymentImage = Resource.bank; break;
                            }
                        });
                        tpl.AddVariable("MoreOrders", cust.Orders.Take(5));
                        tpl.AddVariable(cust);
                    }
                    tpl.AddVariable("Tax", 13);
                },
                wb =>
                {
                    CompareWithGauge(wb, templateFile);
                });
        }

        [Theory,
         InlineData("GroupTagTests_SummaryAbove.xlsx"),
         InlineData("GroupTagTests_MergeLabels.xlsx"),
         InlineData("GroupTagTests_MergeLabels2.xlsx"),
         InlineData("GroupTagTests_PlaceToColumn.xlsx"),
         InlineData("GroupTagTests_NestedGroups.xlsx"),
         InlineData("GroupTagTests_DisableOutline.xlsx"),
         InlineData("GroupTagTests_FormulasWithTagsInGroupRow.xlsx"),
        ]
        public void EmptyDataSource(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl => tpl.AddVariable("Orders", new order[0]),
                wb => { });
        }

        [Theory,
         InlineData("GroupTagTests_SummaryAbove.xlsx"),
         InlineData("GroupTagTests_MergeLabels.xlsx"),
         InlineData("GroupTagTests_MergeLabels2.xlsx"),
         InlineData("GroupTagTests_PlaceToColumn.xlsx"),
         InlineData("GroupTagTests_NestedGroups.xlsx"),
         InlineData("GroupTagTests_DisableOutline.xlsx"),
         InlineData("GroupTagTests_FormulasInGroupRow.xlsx"),
         InlineData("GroupTagTests_MultiRanges.xlsx"),
         InlineData("GroupTagTests_FormulasWithTagsInGroupRow.xlsx"),
         InlineData("GroupTagTests_TotalLabel.xlsx"),
       ]
        public void Customers(string templateFile)
        {
            XlTemplateTest(templateFile,
                tpl =>
                {
                    using (var db = new DbDemos())
                    {
                        var orders = db.orders.LoadWith(x => x.Customer);
                        tpl.AddVariable("Orders", orders);
                    }
                },
                wb =>
                {
                    CompareWithGauge(wb, templateFile);
                });
        }

        [Fact]
        public void WithHeader()
        {
            XlTemplateTest("GroupTagTests_WithHeader.xlsx",
                tpl =>
                {
                    using (var db = new DbDemos())
                        tpl.AddVariable("Orders", db.orders.LoadWith(x => x.Customer).OrderBy(c => c.OrderNo).ToArray());
                },
                wb =>
                {
                    CompareWithGauge(wb, "GroupTagTests_WithHeader.xlsx");
                });
        }
    }
}

---
title: Sorting a list range
---

# Sorting a list range

Regions in ClosedXML.Report may be sorted by columns. This can be done by using tag `<<sort>>` in the target columns of the service row.

![tlists1_sort](../../images/sorting-01.png)

You may choose a descending order by adding the parameter `desc` to the tag (`<<sort desc>>`). When the report is built you'll see the data sorted in descending order.

![tlists1_sort_desc](../../images/sorting-02.png)

You also may add additional columns to the sort by adding the parameter `num` to the tag `sort` (`<<sort num=2>>`). On the picture below you may see the dataset will be sorted first by the column Payment Method, and then by the column Ship Date in the descending order.

![tlists1_sort_num xlsx](../../images/sorting-03.png)

Sorting in ClosedXML.Report has this limitation: all the data of the range to order must be located in a single row. If this is not a option for you consider sorting your data before transferring it to ClosedXML.Report.

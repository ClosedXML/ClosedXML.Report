---
title: Grouping
---

# Grouping

To perform grouping and create subtotals for the columns in ClosedXML.Report there is a tag `<<group>>`. The range is pre-sorted by all columns for which the tags are `<<group>>`, `<<sort>>`, `<<desc>>` and `<<asc>>`. The sort order for the `<<group>>` option is specified by an additional parameter - `<<desc>>` or `<<asc>>` (default asc). By default (without the use of additional options) the work of the `<<group>>` tag is similar to the work of the Subtotal method of the Range object in Excel. If the `<<group>>` tag is specified for several columns, the subtotals are grouped by all these columns. Grouping takes place from right to left, that is, first the totals are grouped by the rightmost column, for which the `<<group>>` tag is specified, then by the column with the `<<group>>` tag to the left of it, etc. The format of the service line of the range is used to format the rows of subtotals. After the subtotals are created, the service line is removed from the range.

To get subtotals you can use aggregation tags in the corresponding columns:
* `<<Sum>>` - displays the amount by column;
* `<<Count>>` - the number of values in the column;
* `<<CountNums>>` - the number of non-empty values in the column;
* `<<Avg>>` or `<<Average>>` - the average value of the column;
* `<<Max>>` - the maximum value in the column;
* `<<Min>>` - the minimum value in the column;
* `<<Product>>` - product by column;
* `<<StDev>>` - standard deviation;
* `<<StDevP>>` - standard deviation of the total population;
* `<<Var>>` - dispersion;
* `<<VarP>>` - dispersion for the general population.

To change behavior, the `<<group>>` tag has a number of options:
* Collapse - causes the subtotal to be collapsed to the level at which the `<<group>>` tag is located with this parameter.
* MergeLabels=[Merge1&#124;Merge2&#124;Merge3] - causes the group cells to be merged in the grouped column
* PlaceToColumn=n - allows you to specify the column in which the group header will be placed
* WithHeader - allows you to create a group header when using subtotals
* Disablesubtotals - allows you to disable the creation of subtotals for the column
* DisableOutline - turns off the creation of an Outline view for the grouped column
* PageBreaks - allows you to place each group on a separate page
* TotalLabel - allows you to set the caption text in the line of subtotals (default: 'Total')
* GrandLabel - allows you to set the caption text in the total totals line (default: 'Grand')

Also, to change the behavior of a grouping, there are range tags:
* `<<SummaryAbove>>` - in case the `<<SummaryAbove>>` tag is found in the range, subtotals are placed above the data;
* `<<DisableGrandTotal>>` - prohibits the creation of grand totals when using range grouping with subtotals.

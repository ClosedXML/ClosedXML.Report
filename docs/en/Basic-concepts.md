---
title: Basic Concepts
---

# Basic Concepts

### Templates
All work on creating a report is built on templates (XLSX-templates) - Excel books that contain a description of the report form, as well as options for books, sheets and report areas. Special field formulas and data areas describe the data in the report structure that you want to transfer to Excel. ClosedXML.Report will call the template and fill the report cells with data from the specified sets.

### Variables
The values passed in ClosedXML.Report with the method `AddVariable` are called variables. They are used to calculate the expressions in the templates. Variable can be added with or without a name. If a variable is added without a name, then all its public fields and properties are added as variables with their names.

Examples:

`template.AddVariable(cust);` 

OR

`template.AddVariable("Customer", cust);`. 


### Expressions 
Expressions are enclosed in double braces {% raw %}{{ }}{% endraw %} and utilize the syntax similar to C#. Lambda expressions are supported.

Examples: 

{% raw %}
`{{item.Product.Price * item.Product.Quantity}}`

`{{items.Where(i => i.Currency == "RUB").Count()}}`
{% endraw %}

### Tags
ClosedXML.Report has a few advanced features allowing to hide a worksheet, sort the data table, apply groupping, calculate totals, etc. These features are controlled by addit tags to the worksheet, to the entire range, or to a single column. Tag is a text embrased by double angle brackets that can be analyzed by a ClosedXML.Report parser. Different tags let you get subtotals, build pivot tables, apply auto-filter, and so on. Tags may have parameters for tuning their behavior. Parameters may require you to specify their values. In this case, the parameter name is separated from the value by the equal sign.

All tags can refer to six report objects: report, sheet, column, row, region, column of area. Report tags are specified in cell A1 of any template sheet. Sheet tags are specified in cell A2 on the sheet. The column's tags are specified in the first line of the sheet. The line tags are specified in the first column of the sheet. The area tags are specified in the leftmost cell of the options row of this area. Tags of the column of the area are specified in the cell of this column in the options row of the area.

Example: `<<Range horizontal>>`.

A list of all the tags you can find on the [Tags page](More-options)

### Ranges
To represent IEnumerable values, Excel named regions are used.

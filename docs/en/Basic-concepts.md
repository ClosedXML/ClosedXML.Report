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
Expressions are enclosed in double braces {{ }} and utilize the syntax similar to C#. Lambda expressions are supported.

Examples: 

{% raw %}
`{{item.Product.Price * item.Product.Quantity}}`

`{{items.Where(i => i.Currency == "RUB").Count()}}`
{% endraw %}

### Tags
ClosedXML.Report has a few advanced features allowing to hide a worksheet, sort the data table, apply groupping, calculate totals, etc. These features are controlled by addit tags to the worksheet, to the entire range, or to a single column. Tag is a text embrased by double angle brackets that can be analyzed by a ClosedXML.Report parser. Different tags let you get subtotals, build pivot tables, apply auto-filter, and so on. Tags may have parameters for tuning their behavior.

Example: `<<Range horizontal>>`.

### Ranges
To represent IEnumerable values, Excel named regions are used.

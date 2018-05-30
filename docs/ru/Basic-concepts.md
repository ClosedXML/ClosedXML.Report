---
title: Базовые сущности
---

# Базовые сущности

### Templates
All work on creating a report is built on templates (XLSX-templates) - Excel books that contain a description of the report form, as well as options for books, sheets and report areas. Special field formulas and data areas describe the data in the report structure that you want to transfer to Excel. ClosedXML.Report will call the template and fill the report cells with data from the specified sets.

### Variables
The values passed in ClosedXML.Report with the method `AddVariable` are called variables. They are used to calculate the expressions used in the templates. Variables can be added with or without a name. If a variable was added without a name, then all the public fields and properties of this instance will be added as variables with their names.

Examples:

`template.AddVariable(cust);` 

OR

`template.AddVariable("Customer", cust);`. 


### Expressions 
Expressions are enclosed in double braces {% raw %}{{ }}{% endraw %} . A syntax similar to C # is used. Lambda expressions are supported.

Examples: 

{% raw %}
`{{item.Product.Price * item.Product.Quantity}}`

`{{items.Where(i => i.Currency == "RUB").Count()}}`
{% endraw %}

### Tags
ClosedXML.Report имеет ряд встроенных возможностей, которые позволяют спрятать лист, отсортировать полученную область, получить итоги по ее колонкам, сгруппировать область и др. Эти дополнительные действия можно вызвать, дополнив книгу-шаблон тэгами листа, области или столбцов. Тэг – это строковое значение, заключённое в двойные угловые скобки и понятное анализатору ClosedXML.Report. Эти опции помогут вам получить промежуточные итоги, включить автофильтр, создать сводные таблицы по области и др. Теги могут иметь параметры.

Example: `<<Range horizontal>>`.

### Ranges
To represent IEnumerable values, Excel regions are used.

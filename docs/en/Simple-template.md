---
title: Simple Template
---

# Simple Template

You can use _expressions_ with braces {{ }} in any cell of any sheet of the _template_
workbook and Excel will find their values at run-time. How? ClosedXML.Report adds a hidden worksheet in a report
workbook and transfers values of all fields for the current record. Then it names all these data cells.

Excel formulas, in which _variables_ are added, must be escaped `&`. As an example: `&=CONCATENATE(Addr1;" "; Addr2)`

Cells with field formulas can be formatted by any known way, including conditional formatting.

Take a simple example:

![simpletemplate](../../images/simple-template-01.png)

```c#
...
        var template = new XLTemplate(workbook);
        var cust = db.Customers.GetById(10);

        template.AddVariable(cust);
        // OR
        template.AddVariable("Company", cust.Company);
        template.AddVariable("Addr1", cust.Addr1);
        template.AddVariable("Addr2", cust.Addr2);
...

public class Customer
{
	public double CustNo { get; set; }
	public string Company { get; set; }
	public string Addr1 { get; set; }
	public string Addr2 { get; set; }
	public string City { get; set; }
	public string State { get; set; }
	public string Zip { get; set; }
	public string Country { get; set; }
	public string Phone { get; set; }
	public string Fax { get; set; }
	public double? TaxRate { get; set; }
	public string Contact { get; set; }
}

```
# ClosedXML.Report
[![Build status](https://ci.appveyor.com/api/projects/status/y2ha69ggalbj1y47/branch/develop?svg=true)](https://ci.appveyor.com/project/ClosedXML/closedxml-report/branch/develop/artifacts)
[![Open Source Helpers](https://www.codetriage.com/closedxml/closedxml.report/badges/users.svg)](https://www.codetriage.com/closedxml/closedxml.report)
[![NuGet Badge](https://buildstats.info/nuget/ClosedXML.Report)](https://www.nuget.org/packages/ClosedXML.Report/)
[![All Contributors](https://img.shields.io/badge/all_contributors-4-orange.svg?style=flat-square)](#contributors)

ClosedXML.Report is a tool for report generation and data analysis in .NET applications through the use of Microsoft Excel.
It is a .NET-library for report generation Microsoft Excel without requiring Excel to be installed on the machine that's running the code. With ClosedXML.Report, you can easily export any data from your .NET classes to Excel using the XLSX-template.

Excel is an excellent alternative to common report generators, and using Excel’s built-in features
can make your reports much more responsive.
Use ClosedXML.Report as a tool for generating files of Excel. Then use Excel visual instruments: formatting (including
conditional formatting), AutoFilter, Pivot tables to construct a versatile data analysis system. With ClosedXML.Report, you can move a lot of report programming
and tuning into Excel. ClosedXML.Report templates are simple and our algorithms are fast – we carefully count every
millisecond – so you waste less time on routine report programming and get surprisingly fast results. If you want
to master such a versatile tool as Excel – ClosedXML.Report is an excellent choice.
Furthermore, ClosedXML.Report doesn’t operate with the usual concepts of band-oriented report tools: Footer, Header,
and Detail. So you get a much greater degree of freedom in report construction and design, and the easiest possible integration of .NET and Microsoft Excel. 

[For more information see the wiki](https://closedxml.github.io/ClosedXML.Report/)

### Install ClosedXML.Report via NuGet

If you want to include ClosedXML.Report in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/ClosedXML.Report/)

To install ClosedXML.Report, run the following command in the Package Manager Console

```
PM> Install-Package ClosedXML.Report
```

## Features

* Copying cell formatting
* Propagation conditional formatting
* Vertical and horizontal tables and subranges
* Ability to implement Excel formulas
* Using dynamically calculated formulas with the syntax of C # and Linq
* Operations with tabular data: sorting, grouping, total functions.
* Pivot tables
* Subranges

## How to use?
To create a report you must first create a report template. You can apply any formatting to any workbook cells, insert pictures, and modify any of the parameters of the workbook itself. In this example, we have turned off the zero values display and hidden the 
gridlines. ClosedXML.Report will preserve all changes to the template. 

**Template**

![template1](https://user-images.githubusercontent.com/1150085/33486458-3161eb92-d6bb-11e7-8833-d500461b18a5.png)

**Code**

```c#
    protected void Report()
    {
        const string outputFile = @".\Output\report.xlsx";
        var template = new XLTemplate(@".\Templates\report.xlsx");

        using (var db = new DbDemos())
        {
            var cust = db.customers.LoadWith(c => c.Orders).First();
            template.AddVariable(cust);
            template.Generate();
        }

        template.SaveAs(outputFile);

        //Show report
        Process.Start(new ProcessStartInfo(outputFile) { UseShellExecute = true });
    }
```

**Result**

![result1](https://user-images.githubusercontent.com/1150085/33486460-31a02542-d6bb-11e7-8899-8694157ee9dd.png)

[For more information see the wiki](https://closedxml.github.io/ClosedXML.Report/) and [tests](https://github.com/ClosedXML/ClosedXML.Report/tree/master/tests)

## Credits
* Afalinasoft team with their XLReport for the idea
* ClosedXML team for a great library
* Logo design by [@Tobaloidee](https://github.com/Tobaloidee)

## Contributors

Thanks goes to these wonderful people ([emoji key](https://github.com/kentcdodds/all-contributors#emoji-key)):

<!-- ALL-CONTRIBUTORS-LIST:START - Do not remove or modify this section -->
<!-- prettier-ignore -->
| [<img src="https://avatars3.githubusercontent.com/u/1150085?v=4" width="100px;"/><br /><sub><b>Rozhkov Alexey</b></sub>](https://github.com/b0bi79)<br />[💻](https://github.com/b0bi79/ClosedXML.Report/commits?author=b0bi79 "Code") [📖](https://github.com/b0bi79/ClosedXML.Report/commits?author=b0bi79 "Documentation") [👀](#review-b0bi79 "Reviewed Pull Requests") | [<img src="https://avatars0.githubusercontent.com/u/19576939?v=4" width="100px;"/><br /><sub><b>Aleksei</b></sub>](https://github.com/Pankraty)<br />[💻](https://github.com/b0bi79/ClosedXML.Report/commits?author=Pankraty "Code") [🌍](#translation-Pankraty "Translation") [👀](#review-Pankraty "Reviewed Pull Requests") [🚇](#infra-Pankraty "Infrastructure (Hosting, Build-Tools, etc)") | [<img src="https://avatars1.githubusercontent.com/u/145854?v=4" width="100px;"/><br /><sub><b>Francois Botha</b></sub>](http://www.vwd.co.za)<br />[📦](#platform-igitur "Packaging/porting to new platform") | [<img src="https://avatars0.githubusercontent.com/u/36028424?v=4" width="100px;"/><br /><sub><b>tobaloidee</b></sub>](https://github.com/Tobaloidee)<br />[🎨](#design-Tobaloidee "Design") |
| :---: | :---: | :---: | :---: |
<!-- ALL-CONTRIBUTORS-LIST:END -->

This project follows the [all-contributors](https://github.com/kentcdodds/all-contributors) specification. Contributions of any kind welcome!

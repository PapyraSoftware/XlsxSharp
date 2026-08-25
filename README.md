![XlsxSharp](https://github.com/PapyraSoftware/XlsxSharp/blob/develop/resources/logo/readme.png)

[![NuGet version (XlsxSharp)](https://img.shields.io/nuget/v/XlsxSharp.svg?style=flat)](https://www.nuget.org/packages/XlsxSharp/) [![.NET](https://img.shields.io/badge/.NET-10.0-512BD4.svg)](#) [![Build & Test](https://github.com/PapyraSoftware/XlsxSharp/actions/workflows/build.yml/badge.svg?branch=develop)](https://github.com/PapyraSoftware/XlsxSharp/actions/workflows/build.yml)

XlsxSharp is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

This is a fork of [ClosedXML](https://github.com/ClosedXML/ClosedXML). With the primary goal to update the library to .NET 10 and to maintain it with the latest .NET versions.

For more information see [the documentation](https://xlsxsharp.readthedocs.io/).

### Install XlsxSharp via NuGet

If you want to include XlsxSharp in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/XlsxSharp)

XlsxSharp is currently published as a prerelease, so the prerelease switch is required. To install XlsxSharp, run the following command in the Package Manager Console

```
PM> Install-Package XlsxSharp -IncludePrerelease
```

### What can you do with this?

XlsxSharp allows you to create Excel files without the Excel application. The typical example is creating Excel reports on a web server.

**Example:**
```c#
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Sample Sheet");
    worksheet.Cell("A1").Value = "Hello World!";
    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
    workbook.SaveAs("HelloWorld.xlsx");
}
```

## Formula parser

XlsxSharp ships its own Pratt (operator-precedence) formula parser (`XlsxSharp.Parser`), supporting both A1 and R1C1 reference styles. It is regression-tested against real-world formula corpora, most notably the [Enron](http://www.felienne.com/archives/3634) and [EUSES](http://eusesconsortium.org/resources.php) spreadsheet datasets used in academic formula-parsing research:

| Dataset | Formulas | Parsed successfully |
|---|---:|--------------------:|
| Enron | 946,320 |             946,303 |
| EUSES | 89,295 |              89,294 |

These figures are asserted as a floor in `XlsxSharp.Parser.Tests/PrattDataSetTests.cs` (`EnronDataSetCoverage`, `EusesDataSetCoverage`), so any regression in parser coverage fails CI.

## Developer guidelines
The [OpenXML specification](https://ecma-international.org/publications-and-standards/standards/ecma-376/) is a large and complicated beast. In order for XlsxSharp, the wrapper around OpenXML, to support all the features, we rely on community contributions. Before opening an issue to request a new feature, we'd like to urge you to try to implement it yourself and log a pull request.

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits
* Project originally created by Manuel de Leon
* Current maintainer: [Stefan Nikolei](https://github.com/stefannikolei)
* Former maintainer: [Jan Havlíček](https://github.com/jahav)
* Former maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)

![XlsxSharp](https://github.com/PapyraSoftware/XlsxSharp/blob/develop/resources/logo/readme.png)

[![Release](https://img.shields.io/badge/release-0.95.4-blue.svg)](https://github.com/PapyraSoftware/XlsxSharp/releases/latest) [![NuGet version (XlsxSharp)](https://img.shields.io/nuget/v/XlsxSharp.svg?style=flat)](https://www.nuget.org/packages/XlsxSharp/) [![.NET Framework](https://img.shields.io/badge/.NET%20Framework-%3E%3D%204.0-red.svg)](#) [![.NET Standard](https://img.shields.io/badge/.NET%20Standard-%3E%3D%202.0-red.svg)](#) [![Build status](https://ci.appveyor.com/api/projects/status/wobbmnlbukxejjgb?svg=true)](https://ci.appveyor.com/project/PapyraSoftware/XlsxSharp/branch/develop/artifacts)
[![Open Source Helpers](https://www.codetriage.com/PapyraSoftware/XlsxSharp/badges/users.svg)](https://www.codetriage.com/PapyraSoftware/XlsxSharp)

XlsxSharp is a .NET library for reading, manipulating and writing Excel 2007+ (.xlsx, .xlsm) files. It aims to provide an intuitive and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

This is a fork of [ClosedXML](https://github.com/ClosedXML/ClosedXML). With the primary goal to update the library to .NET 10 and to maintain it with the latest .NET versions.

For more information see [the documentation](https://closedxml.readthedocs.io/) or [the wiki](https://github.com/closedxml/closedxml/wiki).

### Install XlsxSharp via NuGet

If you want to include XlsxSharp in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/XlsxSharp)

To install XlsxSharp, run the following command in the Package Manager Console

```
PM> Install-Package XlsxSharp
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

## Developer guidelines
The [OpenXML specification](https://ecma-international.org/publications-and-standards/standards/ecma-376/) is a large and complicated beast. In order for XlsxSharp, the wrapper around OpenXML, to support all the features, we rely on community contributions. Before opening an issue to request a new feature, we'd like to urge you to try to implement it yourself and log a pull request.

Please read the [full developer guidelines](CONTRIBUTING.md).

## Credits
* Project originally created by Manuel de Leon
* Current maintainer: [Stefan Nikolei](https://github.com/stefannikolei)
* Former maintainer: [Jan Havlíček](https://github.com/jahav)
* Former maintainer and lead developer: [Francois Botha](https://github.com/igitur)
* Master of Computing Patterns: [Aleksei Pankratev](https://github.com/Pankraty)
* Logo design by [@Tobaloidee](https://github.com/Tobaloidee)

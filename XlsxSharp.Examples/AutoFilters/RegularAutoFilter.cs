using System.IO;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Examples.AutoFilters;

public class RegularAutoFilter : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws;

        #region Single Column Numbers

        string singleColumnNumbers = "Single Column Numbers";
        ws = wb.Worksheets.Add(singleColumnNumbers);

        // Add a bunch of numbers to filter
        ws.Cell("A1")
            .SetValue("Numbers")
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue(5)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(4)
            .CellBelow()
            .SetValue(5);

        ws.Cell("B1")
            .SetValue("Names")
            .CellBelow()
            .SetValue("John")
            .CellBelow()
            .SetValue("Jack")
            .CellBelow()
            .SetValue("Neil")
            .CellBelow()
            .SetValue("Alex")
            .CellBelow()
            .SetValue("Jason")
            .CellBelow()
            .SetValue("Patrick")
            .CellBelow()
            .SetValue("Jacques");

        // Add filters
        IXLAutoFilter autoFilter = ws.RangeUsed().SetAutoFilter();
        autoFilter.Column(1).AddFilter(3).AddFilter(1);

        autoFilter.Column(2).BeginsWith("J");

        // Sort the filtered list
        ws.AutoFilter.Sort();

        #endregion Single Column Numbers

        #region Single Column Strings

        string singleColumnStrings = "Single Column Strings";
        ws = wb.Worksheets.Add(singleColumnStrings);

        // Add a bunch of strings to filter
        ws.Cell("A1")
            .SetValue("Strings")
            .CellBelow()
            .SetValue("B")
            .CellBelow()
            .SetValue("C")
            .CellBelow()
            .SetValue("C")
            .CellBelow()
            .SetValue("E")
            .CellBelow()
            .SetValue("A")
            .CellBelow()
            .SetValue("D");

        // Add filters
        ws.RangeUsed().SetAutoFilter().Column(1).AddFilter("C").AddFilter("A");

        // Sort the filtered list
        ws.AutoFilter.Sort();

        #endregion Single Column Strings

        #region Single Column Mixed

        string singleColumnMixed = "Single Column Mixed";
        ws = wb.Worksheets.Add(singleColumnMixed);

        // Add a bunch of items to filter
        ws.Cell("A1")
            .SetValue("Mixed")
            .CellBelow()
            .SetValue("B")
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue("C")
            .CellBelow()
            .SetValue("E")
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(4);

        // Add filters
        ws.RangeUsed().SetAutoFilter().Column(1).AddFilter("C").AddFilter(1);

        // Sort the filtered list
        ws.AutoFilter.Sort();

        #endregion Single Column Mixed

        #region Multi Column

        string multiColumn = "Multi Column";
        ws = wb.Worksheets.Add(multiColumn);

        ws.Cell("A1")
            .SetValue("First")
            .CellBelow()
            .SetValue("B")
            .CellBelow()
            .SetValue("C")
            .CellBelow()
            .SetValue("C")
            .CellBelow()
            .SetValue("E")
            .CellBelow()
            .SetValue("A")
            .CellBelow()
            .SetValue("D");

        ws.Cell("B1")
            .SetValue("Numbers")
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue(5)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(4);

        ws.Cell("C1")
            .SetValue("Strings")
            .CellBelow()
            .SetValue("B")
            .CellBelow()
            .SetValue("C")
            .CellBelow()
            .SetValue("C")
            .CellBelow()
            .SetValue("E")
            .CellBelow()
            .SetValue("A")
            .CellBelow()
            .SetValue("D");

        // Add filters
        ws.RangeUsed().SetAutoFilter().Column(2).AddFilter(3).AddFilter(1);

        // Sort the filtered list
        ws.AutoFilter.Sort(3);

        #endregion Multi Column

        #region Table

        string tableSheetName = "Table";
        ws = wb.Worksheets.Add(tableSheetName);

        // Add a bunch of numbers to filter
        ws.Cell("A1")
            .SetValue("Numbers")
            .CellBelow()
            .SetValue(2)
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue(3)
            .CellBelow()
            .SetValue(5)
            .CellBelow()
            .SetValue(1)
            .CellBelow()
            .SetValue(4);

        // Add filters
        IXLTable table = ws.RangeUsed().CreateTable();
        table.ShowTotalsRow = true;
        table.Field(0).TotalsRowFunction = XLTotalsRowFunction.Sum;
        table.AutoFilter.Column(1).AddFilter(3).AddFilter(4);

        table.AutoFilter.Sort();

        #endregion Table

        using (MemoryStream ms = new())
        {
            wb.SaveAs(ms);

            XLWorkbook workbook = new(ms);

            #region Single Column Numbers

            workbook.Worksheet(singleColumnNumbers).AutoFilter.Column(1).AddFilter(5);
            workbook.Worksheet(singleColumnNumbers).AutoFilter.Sort(1, XLSortOrder.Descending);

            #endregion Single Column Numbers

            #region Single Column Strings

            workbook.Worksheet(singleColumnStrings).AutoFilter.Column(1).AddFilter("E");
            workbook.Worksheet(singleColumnStrings).AutoFilter.Sort(1, XLSortOrder.Descending);

            #endregion Single Column Strings

            #region Single Column Mixed

            workbook.Worksheet(singleColumnMixed).AutoFilter.Column(1).AddFilter("E");
            workbook.Worksheet(singleColumnMixed).AutoFilter.Column(1).AddFilter(3);
            workbook.Worksheet(singleColumnMixed).AutoFilter.Sort(1, XLSortOrder.Descending);

            #endregion Single Column Mixed

            #region Multi Column

            workbook.Worksheet(multiColumn).AutoFilter.Column(3).AddFilter("C");
            workbook.Worksheet(multiColumn).AutoFilter.Sort(3, XLSortOrder.Descending);

            #endregion Multi Column

            #region Table

            workbook.Worksheet(tableSheetName).Table(0).AutoFilter.Column(1).AddFilter(5);
            workbook.Worksheet(tableSheetName).Table(0).AutoFilter.Sort(1, XLSortOrder.Descending);

            #endregion Table

            workbook.SaveAs(filePath);
            ms.Close();
        }
    }
}

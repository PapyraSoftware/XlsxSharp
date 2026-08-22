using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class Collections : IXLExample
{
    public void Create(String filePath)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Collections");

        // From a list of strings
        List<string> listOfStrings = ["House", "Car"];
        ws.Cell(1, 1).Value = "Strings";
        ws.Cell(1, 1).AsRange().AddToNamed("Titles");
        ws.Cell(2, 1).InsertData(listOfStrings);

        // From a list of arrays
        List<int[]> listOfArr =
        [
            [1, 2, 3],
            [1],
            [1, 2, 3, 4, 5, 6],
        ];
        ws.Cell(1, 3).Value = "Arrays";
        ws.Range(1, 3, 1, 8).Merge().AddToNamed("Titles");
        ws.Cell(2, 3).InsertData(listOfArr);

        // From a DataTable
        DataTable dataTable = GetTable();
        ws.Cell(6, 1).Value = "DataTable";
        ws.Range(6, 1, 6, 4).Merge().AddToNamed("Titles");
        ws.Cell(7, 1).InsertData(dataTable);

        // From a query
        List<Person> list =
        [
            new()
            {
                Name = "John",
                Age = 30,
                House = "On Elm St.",
            },
            new()
            {
                Name = "Mary",
                Age = 15,
                House = "On Main St.",
            },
            new()
            {
                Name = "Luis",
                Age = 21,
                House = "On 23rd St.",
            },
            new()
            {
                Name = "Henry",
                Age = 45,
                House = "On 5th Ave.",
            },
        ];

        var people =
            from p in list
            where p.Age >= 21
            select new
            {
                p.Name,
                p.House,
                p.Age,
            };

        ws.Cell(6, 6).Value = "Query";
        ws.Range(6, 6, 6, 8).Merge().AddToNamed("Titles");
        ws.Cell(7, 6).InsertData(people);

        // Format all titles in one shot
        wb.DefinedNames.DefinedName("Titles")
            .Ranges.Style.Font.SetBold()
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Fill.SetBackgroundColor(XLColor.Cyan);

        ws.Columns().AdjustToContents();

        wb.SaveAs(filePath);
    }

    class Person
    {
        public String House { get; set; }
        public String Name { get; set; }
        public Int32 Age { get; set; }
    }

    private static DataTable GetTable()
    {
        DataTable table = new();
        table.Columns.Add("Dosage", typeof(int));
        table.Columns.Add("Drug", typeof(string));
        table.Columns.Add("Patient", typeof(string));
        table.Columns.Add("Date", typeof(DateTime));

        table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
        table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
        table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
        table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
        table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
        return table;
    }
}

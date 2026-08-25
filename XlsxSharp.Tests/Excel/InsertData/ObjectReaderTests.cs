using System.Collections;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.InsertData;
using XlsxSharp.Tests.Excel.Tables;

namespace XlsxSharp.Tests.Excel.InsertData;

public class ObjectReaderTests
{
    private static readonly TablesTests.TestObjectWithAttributes[] ObjectWithAttributes =
    [
        new()
        {
            Column1 = "Value 1",
            Column2 = "Value 2",
            UnOrderedColumn = 3,
            MyField = 4,
        },
        new()
        {
            Column1 = "Value 5",
            Column2 = "Value 6",
            UnOrderedColumn = 7,
            MyField = 8,
        },
    ];

    private static readonly TablesTests.TestObjectWithoutAttributes[] ObjectWithoutAttributes =
    [
        new() { Column1 = "Value 9", Column2 = "Value 10" },
        new() { Column1 = "Value 11", Column2 = "Value 12" },
    ];

    private static readonly TestPoint[] Structs =
    [
        new()
        {
            X = 1,
            Y = 2,
            Z = 3,
        },
        new(),
    ];

    private static readonly TestPoint?[] NullableStructs =
    [
        new TestPoint
        {
            X = 1,
            Y = 2,
            Z = 3,
        },
        new TestPoint(),
        null,
    ];

    [Test]
    [MethodDataSource(nameof(ObjectSourceNames))]
    public void CanGetPropertyName(IEnumerable data, int propertyIndex, string expected)
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        ClassicAssert.AreEqual(expected, reader.GetPropertyName(propertyIndex));
    }

    public static IEnumerable<(IEnumerable, int, string)> ObjectSourceNames()
    {
        IEnumerable data = ObjectWithoutAttributes;
        yield return (data, 0, "Column1");
        yield return (data, 1, "Column2");

        data = ObjectWithAttributes;
        yield return (data, 0, "FirstColumn");
        yield return (data, 1, "SecondColumn");
        yield return (data, 2, "SomeFieldNotProperty");
        yield return (data, 3, "UnOrderedColumn");

        data = Structs;
        yield return (data, 0, "X");
        yield return (data, 1, "Y");
        yield return (data, 2, "Z");

        data = NullableStructs;
        yield return (data, 0, "X");
        yield return (data, 1, "Y");
        yield return (data, 2, "Z");
    }

    [Test]
    [MethodDataSource(nameof(PropertyCounts))]
    public void CanGetPropertiesCount(IEnumerable data, int expected)
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        ClassicAssert.AreEqual(expected, reader.GetPropertiesCount());
    }

    public static IEnumerable<(IEnumerable, int)> PropertyCounts()
    {
        yield return (ObjectWithoutAttributes, 2);
        yield return (ObjectWithAttributes, 4);
        yield return (Structs, 3);
        yield return (NullableStructs, 3);
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(
            ObjectWithAttributes
        );
        ClassicAssert.AreEqual(2, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValuesFromObject()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(
            ObjectWithAttributes
        );
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        XLCellValue[] firstRecord = [.. result.First()];
        XLCellValue[] lastRecord = [.. result.Last()];

        ClassicAssert.AreEqual("Value 2", firstRecord[0]);
        ClassicAssert.AreEqual("Value 1", firstRecord[1]);
        ClassicAssert.AreEqual(4, firstRecord[2]);
        ClassicAssert.AreEqual(3, firstRecord[3]);

        ClassicAssert.AreEqual("Value 6", lastRecord[0]);
        ClassicAssert.AreEqual("Value 5", lastRecord[1]);
        ClassicAssert.AreEqual(8, lastRecord[2]);
        ClassicAssert.AreEqual(7, lastRecord[3]);
    }

    [Test]
    public void CanReadValuesFromStruct()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(Structs);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        XLCellValue[] firstRecord = [.. result.First()];
        XLCellValue[] lastRecord = [.. result.Last()];

        ClassicAssert.AreEqual(1, firstRecord[0]);
        ClassicAssert.AreEqual(2, firstRecord[1]);
        ClassicAssert.AreEqual(3, firstRecord[2]);

        ClassicAssert.AreEqual(0, lastRecord[0]);
        ClassicAssert.AreEqual(0, lastRecord[1]);
        ClassicAssert.AreEqual(Blank.Value, lastRecord[2]);
    }

    [Test]
    public void CanReadValuesFromNullableStruct()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(NullableStructs);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        XLCellValue[] firstRecord = [.. result.First()];
        XLCellValue[] lastRecord = [.. result.Last()];

        ClassicAssert.AreEqual(1, firstRecord[0]);
        ClassicAssert.AreEqual(2, firstRecord[1]);
        ClassicAssert.AreEqual(3, firstRecord[2]);

        ClassicAssert.AreEqual(Blank.Value, lastRecord[0]);
        ClassicAssert.AreEqual(Blank.Value, lastRecord[1]);
        ClassicAssert.AreEqual(Blank.Value, lastRecord[2]);
    }

    [Test]
    public void IgnoresIndexers()
    {
        TestClassWithIndexer[] data = [new()];
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);

        ClassicAssert.AreEqual(1, reader.GetPropertiesCount());
        ClassicAssert.AreEqual(nameof(TestClassWithIndexer.Value), reader.GetPropertyName(0));
    }

    private record TestClassWithIndexer
    {
        public static int Value => 0;
        public int this[int i] => 0;
    }

    private struct TestPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double? Z { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.InsertData;

namespace XlsxSharp.Tests.Excel.InsertData;

public class SimpleTypeReaderTests
{
    private readonly int[] data = [1, 2, 3];

    [TestCaseSource(nameof(SimpleSourceNames))]
    public string CanGetPropertyName<T>(IEnumerable<T> data)
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(data);
        return reader.GetPropertyName(0);
    }

    private static IEnumerable<TestCaseData> SimpleSourceNames
    {
        get
        {
            yield return new TestCaseData(new[] { 1, 2, 3 }).Returns("Int32");
            yield return new TestCaseData(new List<double> { 1.0, 2.0, 3.0 }).Returns("Double");
            yield return new TestCaseData(new[] { 1.0m, 2.0m, 3.0m }).Returns("Decimal");
            yield return new TestCaseData(arg: new[] { "A", "B", "C" }).Returns("String");
            yield return new TestCaseData(new[] { 'A', 'B', 'C' }).Returns("Char");
            yield return new TestCaseData(new[] { new DateTime(2020, 1, 1) }).Returns("DateTime");
        }
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(1, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        Assert.AreEqual(3, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValues()
    {
        IInsertDataReader reader = InsertDataReaderFactory.Instance.CreateReader(this.data);
        IEnumerable<IEnumerable<XLCellValue>> result = reader.GetRecords();

        Assert.AreEqual(1, result.First().Single());
        Assert.AreEqual(3, result.Last().Single());
    }
}

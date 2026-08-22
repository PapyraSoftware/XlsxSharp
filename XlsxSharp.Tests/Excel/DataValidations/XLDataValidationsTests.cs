using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests.Excel.DataValidations;

public class XlDataValidationsTests
{
    [Test]
    public void AddedRangesAreTransferredToTargetSheet()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.AddWorksheet();
            IXLWorksheet ws2 = wb.AddWorksheet();

            IXLDataValidation dv1 = ws1.Range("A1:A3").CreateDataValidation();
            dv1.MinValue = "100";

            IXLDataValidation dv2 = ws2.DataValidations.Add(dv1);

            Assert.AreEqual(1, ws1.DataValidations.Count());
            Assert.AreEqual(1, ws2.DataValidations.Count());

            Assert.AreNotSame(dv1, dv2);

            Assert.AreSame(ws1, dv1.Ranges.Single().Worksheet);
            Assert.AreSame(ws2, dv2.Ranges.Single().Worksheet);
        }
    }

    [Test]
    [Description("Ensure one-dv-per-cell invariant")]
    public void AddRangeReplacesIntersectingAreasOfValidation()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLDataValidation dv = ws.Range("A1:C3").CreateDataValidation();
        dv.MinValue = "10";

        dv.AddRange(ws.Range("B1:D4"));

        Assert.AreEqual("A1:A3 B1:D4", ToSpaceList(dv.Ranges));
        Assert.AreEqual("10", dv.MinValue);
    }

    [Test]
    public void AddRangeKeepsValidationSettingsEvenWhenItCompletelyCoversOriginal()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLDataValidation dv = ws.Range("B2:C3").CreateDataValidation();
        dv.MinValue = "10";

        dv.AddRange(ws.Range("A1:D4"));

        Assert.AreEqual("A1:D4", ToSpaceList(dv.Ranges));
        Assert.AreEqual("10", dv.MinValue);
    }

    [Test]
    [Description("Ensure one-dv-per-cell invariant")]
    public void AddRangeReplacesAreaOfOtherValidations()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLDataValidation dv1 = ws.Range("A1:A3").CreateDataValidation();
        dv1.WholeNumber.Between(1, 5);

        IXLDataValidation dv2 = ws.Range("B2:D4").CreateDataValidation();
        dv1.MaxValue = "10";

        // Act
        dv1.AddRange(ws.Range("B1:D3"));

        // Assert
        Assert.AreEqual("A1:A3 B1:D3", ToSpaceList(dv1.Ranges));
        Assert.AreEqual("B4:D4", ToSpaceList(dv2.Ranges));
    }

    [Test]
    public void AddRangeDeletesOtherValidationsThatAreNotUsedByAnyCell()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLDataValidation dv1 = ws.Range("A1:A3").CreateDataValidation();
        dv1.WholeNumber.Between(1, 5);

        IXLDataValidation dv2 = ws.Range("B1:B3").CreateDataValidation();
        dv1.MaxValue = "10";

        // Act
        dv1.AddRange(ws.Range("B1:B3"));

        // Assert
        Assert.AreEqual(1, ws.DataValidations.Count());

        Assert.IsTrue(ws.DataValidations.Contains(dv1));
        Assert.AreEqual("A1:A3 B1:B3", ToSpaceList(dv1.Ranges));

        Assert.IsFalse(ws.DataValidations.Contains(dv2));
        Assert.IsEmpty(ToSpaceList(dv2.Ranges));
    }

    [TestCase("A1:A1", true)]
    [TestCase("A1:A3", true)]
    [TestCase("A1:A4", false)]
    [TestCase("C2:C2", true)]
    [TestCase("C1:C3", true)]
    [TestCase("A1:C3", false)]
    public void CanFindDataValidationForRange(string searchAddress, bool expectedResult)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLDataValidation dv = ws.Range("A1:A3").CreateDataValidation();
            dv.MinValue = "100";
            dv.AddRange(ws.Range("C1:C3"));

            XLRangeAddress address = new(ws as XLWorksheet, searchAddress);

            bool actualResult = ws.DataValidations.TryGet(address, out IXLDataValidation? foundDv);
            Assert.AreEqual(expectedResult, actualResult);
            if (expectedResult)
            {
                Assert.AreSame(dv, foundDv);
            }
            else
            {
                Assert.IsNull(foundDv);
            }
        }
    }

    [TestCase("A1:A1", 1)]
    [TestCase("A1:A3", 1)]
    [TestCase("B1:B4", 0)]
    [TestCase("A1:C3", 1)]
    [TestCase("C2:C3", 1)]
    [TestCase("C2:G6", 2)]
    [TestCase("E2:E3", 0)]
    public void CanGetAllDataValidationsForRange(string searchAddress, int expectedCount)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLDataValidation dv1 = ws.Range("A1:A3").CreateDataValidation();
            dv1.MinValue = "100";
            dv1.AddRange(ws.Range("C1:C3"));

            IXLDataValidation dv2 = ws.Range("E4:G6").CreateDataValidation();
            dv2.MinValue = "200";

            XLRangeAddress address = new(ws as XLWorksheet, searchAddress);

            IEnumerable<IXLDataValidation> actualResult = ws.DataValidations.GetAllInRange(address);

            Assert.AreEqual(expectedCount, actualResult.Count());
        }
    }

    [Test]
    public void AddDataValidationSplitsExistingRanges()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLDataValidation dv1 = ws.Ranges("B2:G7,C11:C13").CreateDataValidation();
            dv1.MinValue = "100";

            IXLDataValidation dv2 = ws.Range("E4:G6").CreateDataValidation();
            dv2.MinValue = "100";

            Assert.AreEqual(4, dv1.Ranges.Count());
            Assert.AreEqual("B2:G3 B4:D6 B7:G7 C11:C13", dv1.Ranges.ToSpaceList());
        }
    }

    [Test]
    public void RemovedRangeExcludedFromIndex()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLDataValidation dv = ws.Range("A1:A3").CreateDataValidation();
            dv.MinValue = "100";
            IXLRange range = ws.Range("C1:C3");
            dv.AddRange(range);

            dv.RemoveRange(range);

            bool actualResult = ws.DataValidations.TryGet(
                range.RangeAddress,
                out IXLDataValidation? foundDv
            );
            Assert.IsFalse(actualResult);
            Assert.IsNull(foundDv);
        }
    }

    [Test]
    public void ConsolidatedDataValidationsAreUnsubscribed()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLDataValidation dv1 = ws.Range("A1:A3").CreateDataValidation();
            dv1.MinValue = "100";
            IXLDataValidation dv2 = ws.Range("B1:B3").CreateDataValidation();
            dv2.MinValue = "100";

            ((XLDataValidations)ws.DataValidations).Consolidate();
            dv1.AddRange(ws.Range("C1:C3"));
            dv2.AddRange(ws.Range("D1:D3"));

            IXLDataValidation consolidatedDv = ws.DataValidations.Single();
            Assert.AreSame(dv1, consolidatedDv);
            Assert.True(ws.Cell("C1").HasDataValidation);
            Assert.False(ws.Cell("D1").HasDataValidation);
        }
    }

    private static string ToSpaceList(IEnumerable<IXLRange> ranges) =>
        string.Join(" ", ranges.Select(x => x.RangeAddress.ToStringRelative()));
}

using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace XlsxSharp.Tests.Excel.PageSetup;

[TestFixture]
public class PrintAreaTests
{
    [Test]
    [TestCase("A1:B2")]
    [TestCase("A1:B2", "D3:D5")]
    public void CanLoadWorksheetWithMultiplePrintAreas(params string[] printAreaRangeAddresses) =>
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                foreach (string printAreaRangeAddress in printAreaRangeAddresses)
                {
                    ws.PageSetup.PrintAreas.Add(printAreaRangeAddress);
                }
            },
            (_, ws) =>
            {
                IEnumerable<string> actualPrintAddresses = ws.PageSetup.PrintAreas.Select(pa =>
                    pa.RangeAddress.ToStringRelative()
                );
                CollectionAssert.AreEqual(printAreaRangeAddresses, actualPrintAddresses);
            }
        );
}

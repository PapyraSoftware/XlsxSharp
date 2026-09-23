using XlsxSharp.Excel.IO;

namespace XlsxSharp.Tests.Excel.IO;

public class SpreadsheetXmlTests
{
    [Test]
    [Arguments("1", true)]
    [Arguments("true", true)]
    [Arguments(" true\n", true)]
    [Arguments("0", false)]
    [Arguments("false", false)]
    public void BooleanAcceptsTheXsdForms(string value, bool expected) =>
        ClassicAssert.AreEqual(expected, SpreadsheetXml.ParseBoolean(value));

    [Test]
    [Arguments("on")]
    [Arguments("off")]
    [Arguments("True")]
    [Arguments("FALSE")]
    [Arguments("-1")]
    [Arguments("")]
    public void BooleanOutsideTheSchemaReadsAsAbsent(string value) =>
        ClassicAssert.IsNull(SpreadsheetXml.ParseBoolean(value));
}

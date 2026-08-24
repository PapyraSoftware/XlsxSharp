namespace XlsxSharp.Tests.Excel.Misc;

public class XlHelperTests
{
    private static void CheckColumnNumber(int column) =>
        ClassicAssert.AreEqual(
            column,
            XLHelper.GetColumnNumberFromLetter(XLHelper.GetColumnLetterFromNumber(column))
        );

    [Test]
    public void InvalidA1Addresses()
    {
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address(""));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("A"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("a"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("-1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("AAAA1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("XFG1"));

        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("@A1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("@AA1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("@AAA1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("[A1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("[AA1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("[AAA1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("{A1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("{AA1"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("{AAA1"));

        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("A1@"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("AA1@"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("AAA1@"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("A1["));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("AA1["));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("AAA1["));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("A1{"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("AA1{"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("AAA1{"));

        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("@A1@"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("@AA1@"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("@AAA1@"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("[A1["));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("[AA1["));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("[AAA1["));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("{A1{"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("{AA1{"));
        ClassicAssert.IsFalse(XLHelper.IsValidA1Address("{AAA1{"));
    }

    [Test]
    public void PlusAa1IsNotAnAddress() => ClassicAssert.IsFalse(XLHelper.IsValidA1Address("+AA1"));

    [Test]
    public void TestConvertColumnLetterToNumberAnd()
    {
        CheckColumnNumber(1);
        CheckColumnNumber(27);
        CheckColumnNumber(28);
        CheckColumnNumber(52);
        CheckColumnNumber(53);
        CheckColumnNumber(1000);
        CheckColumnNumber(1353);
    }

    [Test]
    public void ValidA1Addresses()
    {
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("A1"));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("A" + XLHelper.MaxRowNumber));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("Z1"));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("Z" + XLHelper.MaxRowNumber));

        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("AA1"));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("AA" + XLHelper.MaxRowNumber));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("ZZ1"));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("ZZ" + XLHelper.MaxRowNumber));

        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("AAA1"));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address("AAA" + XLHelper.MaxRowNumber));
        ClassicAssert.IsTrue(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + "1"));
        ClassicAssert.IsTrue(
            XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + XLHelper.MaxRowNumber)
        );
    }

    [Test]
    public void TestColumnLetterLookup()
    {
        List<string> columnLetters = [];
        for (int c = 1; c <= XLHelper.MaxColumnNumber; c++)
        {
            string columnLetter = NaiveGetColumnLetterFromNumber(c);
            columnLetters.Add(columnLetter);

            ClassicAssert.AreEqual(columnLetter, XLHelper.GetColumnLetterFromNumber(c));
        }

        foreach (string cl in columnLetters)
        {
            int columnNumber = NaiveGetColumnNumberFromLetter(cl);
            ClassicAssert.AreEqual(columnNumber, XLHelper.GetColumnNumberFromLetter(cl));
        }
    }

    [Test]
    [Arguments("R")]
    [Arguments("C")]
    [Arguments("RC")]
    [Arguments("R111C222")]
    [Arguments("R[]C")]
    [Arguments("RC[]")]
    [Arguments("R[]C[]")]
    [Arguments("R[111]C222")]
    [Arguments("R111C[222]")]
    [Arguments("R[111]C[222]")]
    [Arguments("R[-111]C[-222]")]
    public void ValidRcAddresses(string address) =>
        ClassicAssert.IsTrue(XLHelper.IsValidRCAddress(address));

    [Test]
    [Arguments("RD")]
    [Arguments("CC")]
    [Arguments("R[-]C222")]
    [Arguments("R[]C[-]")]
    [Arguments("_R111C222")]
    public void InvalidRcAddresses(string address) =>
        ClassicAssert.IsFalse(XLHelper.IsValidRCAddress(address));

    #region Old XLHelper methods

    private static readonly string[] Letters =
    [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "V",
        "W",
        "X",
        "Y",
        "Z",
    ];

    /// <summary>
    /// These used to be the methods in XLHelper, but were later changed
    /// We now use them as a check against the new methods
    /// Gets the column number of a given column letter.
    /// </summary>
    /// <param name="columnLetter"> The column letter to translate into a column number. </param>
    private static int NaiveGetColumnNumberFromLetter(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter))
        {
            throw new ArgumentNullException("columnLetter");
        }

        int retVal;
        columnLetter = columnLetter.ToUpper();

        //Extra check because we allow users to pass row col positions in as strings
        if (columnLetter[0] <= '9')
        {
            retVal = int.Parse(columnLetter, XLHelper.NumberStyle, XLHelper.ParseCulture);
            return retVal;
        }

        int sum = 0;

        for (int i = 0; i < columnLetter.Length; i++)
        {
            sum *= 26;
            sum += (columnLetter[i] - 'A' + 1);
        }

        return sum;
    }

    /// <summary>
    /// Gets the column letter of a given column number.
    /// </summary>
    /// <param name="columnNumber">The column number to translate into a column letter.</param>
    /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
    private static string NaiveGetColumnLetterFromNumber(
        int columnNumber,
        bool trimToAllowed = false
    )
    {
        if (trimToAllowed)
        {
            columnNumber = XLHelper.TrimColumnNumber(columnNumber);
        }

        columnNumber--; // Adjust for start on column 1
        if (columnNumber <= 25)
        {
            return Letters[columnNumber];
        }
        int firstPart = (columnNumber) / 26;
        int remainder = ((columnNumber) % 26) + 1;
        return NaiveGetColumnLetterFromNumber(firstPart)
            + NaiveGetColumnLetterFromNumber(remainder);
    }

    #endregion Old XLHelper methods
}

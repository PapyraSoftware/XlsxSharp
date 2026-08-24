using System.Globalization;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class TextTests
{
    [Test]
    [Arguments(@"ABCDEF123", @"ABCDEF123")]
    [Arguments(@"ァィゥェォッャュョヮ", @"ｧｨｩｪｫｯｬｭｮヮ")] // Small katakana, there is no half wa variant
    [Arguments(
        @"アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン",
        @"ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ"
    )]
    [Arguments(
        "！＂＃\uff04％＆＇（）＊\uff0b，－．／０１２３４５６７８９：；\uff1c\uff1d\uff1e？＠",
        @"!""#$%&'()*+,-./0123456789:;<=>?@"
    )]
    [Arguments(
        @"ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ",
        @"ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    )]
    [Arguments(
        "［＼］\uff3e＿\uff40ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛\uff5c｝\uff5e",
        @"[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    )]
    [Arguments(@"―‘’”、。「」゛゜・ー￥", @"ｰ`'""､｡｢｣ﾞﾟ･ｰ\")]
    public void AscConvertsFullwidthCharactersToHalfwidthCharacters(
        string input,
        string expected
    ) => ClassicAssert.AreEqual(expected, XLWorkbook.EvaluateExpr($"ASC(\"{input}\")"));

    [Test]
    public void CharReturnsErrorOnEmptyString() =>
        // Calc engine tries to coerce it to number and fails. It never even reaches the functions.
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"CHAR("""")"));

    [Test]
    [Arguments(0)]
    [Arguments(256)]
    [Arguments(9797)]
    public void CharNumberMustBeBetween1And255(int number) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"CHAR({number})")
        );

    [Test]
    [Arguments(48, '0')]
    [Arguments(97, 'a')]
    [Arguments(128, '€')]
    [Arguments(138, 'Š')]
    [Arguments(169, '©')]
    [Arguments(182, '¶')]
    [Arguments(230, 'æ')]
    [Arguments(255, 'ÿ')]
    [Arguments(255.9, 'ÿ')]
    public void CharInterpretsNumberAsWin1252(double number, char expected)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"CHAR({number})");
        ClassicAssert.AreEqual(expected.ToString(), actual);
    }

    [Test]
    public void CleanEmptyStringIsEmptyString() =>
        ClassicAssert.AreEqual("", XLWorkbook.EvaluateExpr(@"CLEAN("""")"));

    [Test]
    public void CleanRemovesControlCharacters()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"CLEAN(CHAR(9)&""Monthly report""&CHAR(10))");
        ClassicAssert.AreEqual("Monthly report", actual);

        actual = XLWorkbook.EvaluateExpr(@"CLEAN(""   "")");
        ClassicAssert.AreEqual("   ", actual);
    }

    [Test]
    public void CodeReturnsErrorOnEmptyString() =>
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"CODE("""")"));

    [Test]
    [Arguments("A", 65)]
    [Arguments("BCD", 66)]
    [Arguments("€", 128)]
    [Arguments("ÿ", 255)]
    public void CodeReturnsWin1252CodepointOfFirstCharacter(string text, int expected)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"CODE(\"{text}\")");
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    public void CodeIsInverseToChar()
    {
        for (int i = 1; i < 256; ++i)
        {
            ClassicAssert.AreEqual(i, XLWorkbook.EvaluateExpr($"CODE(CHAR({i}))"));
        }
    }

    [Test]
    [Arguments("π")]
    [Arguments("ب")]
    [Arguments("😃")]
    [Arguments("♫")]
    [Arguments("ひ")]
    public void CodeReturnsQuestionMarkCodeOnNonWin1252Chars(string text)
    {
        XLCellValue expected = XLWorkbook.EvaluateExpr("CODE(\"?\")");
        XLCellValue actual = XLWorkbook.EvaluateExpr($"CODE(\"{text}\")");
        ClassicAssert.AreEqual(63, expected);
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    [Culture("cs-CZ")]
    public void ConcatConcatenatesScalarValues()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellValue actual = ws.Evaluate(@"CONCAT(""ABC"",123,TRUE,IF(TRUE,),1.25)");
        ClassicAssert.AreEqual("ABC123TRUE1,25", actual);

        actual = ws.Evaluate(@"CONCAT("""",""123"")");
        ClassicAssert.AreEqual("123", actual);

        ws.FirstCell()
            .SetValue(20.5)
            .CellBelow()
            .SetValue("AB")
            .CellBelow()
            .SetFormulaA1("DATE(2019,1,1)")
            .CellBelow()
            .SetFormulaA1("CONCAT(A1:A3)");

        actual = ws.Cell("A4").Value;
        ClassicAssert.AreEqual("20,5AB43466", actual);
    }

    [Test]
    public void ConcatConcatenatesArrayValues() =>
        ClassicAssert.AreEqual(
            "ABC0123456789Z",
            XLWorkbook.EvaluateExpr(@"CONCAT({""A"",""B"",""C""},{0,1},{2;3},{4,5,6;7,8,9},""Z"")")
        );

    [Test]
    public void ConcatConcatenatesReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("C2").InsertData(new object[] { ("A", "B", "C"), (1, 2, 3, 4), (5, 6, 7, 8) });
        ClassicAssert.AreEqual("ABC12345678AZ", ws.Evaluate("CONCAT(C2:E2,C3:F4,C2,\"Z\")"));
    }

    [Test]
    public void ConcatHasLimitOf32767Characters() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("CONCAT(REPT(\"A\",32768))")
        );

    [Test]
    public void ConcatAcceptsOnlyAreaReferences()
    {
        // Only areas are accepted, not unions
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            ws.Evaluate("CONCAT((C2:E2,C3:F4),C2,\"Z\")")
        );
    }

    [Test]
    public void ConcatPropagatesErrorValues()
    {
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",#DIV/0!,5)")
        );
        ClassicAssert.AreEqual(
            XLError.DivisionByZero,
            XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",{""D"",#DIV/0!,7},5)")
        );

        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B5").SetValue(XLError.DivisionByZero).CellBelow().SetValue(5);
        ClassicAssert.AreEqual(XLError.DivisionByZero, ws.Evaluate("CONCAT(\"ABC\",B5:B6)"));
    }

    [Test]
    public void ConcatTreatsBlanksAsEmptyString() =>
        ClassicAssert.AreEqual("ABC123", XLWorkbook.EvaluateExpr(@"CONCAT(""ABC"",,""123"",)"));

    [Test]
    [Culture("cs-CZ")]
    public void ConcatenateConcatenatesScalarValues()
    {
        using XLWorkbook wb = new();
        XLCellValue actual = wb.Evaluate(@"CONCATENATE(""ABC"",123,4.56,IF(TRUE,),TRUE)");
        ClassicAssert.AreEqual("ABC1234,56TRUE", actual);

        actual = wb.Evaluate(@"CONCATENATE("""",""123"")");
        ClassicAssert.AreEqual("123", actual);
    }

    [Test]
    public void ConcatenateWithReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Value = "Hello";
        ws.Cell("B1").Value = "World";
        ws.Cell("C1").FormulaA1 = "CONCATENATE(A1:A2,\" \",B1:B2)";
        ws.Cell("A3").FormulaA1 = "CONCATENATE(A1:A2,\" \",B1:B2)";

        ClassicAssert.AreEqual("Hello World", ws.Evaluate(@"CONCATENATE(A1,"" "",B1)"));

        // The result on C1 is on the same row (only one intersected cell) means implicit intersection
        // results in a one value per intersection and thus correct value. The A3 intersects two cells
        // and thus results in #VALUE! error.
        ClassicAssert.AreEqual("Hello World", ws.Cell("C1").Value);
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell("A3").Value);
    }

    [Test]
    public void ConcatenateHasLimitOf32767Characters()
    {
        ClassicAssert.AreNotEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("CONCATENATE(REPT(\"A\",32767))")
        );
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("CONCATENATE(REPT(\"A\",32768))")
        );
    }

    [Test]
    public void ConcatenateUsesImplicitIntersectionOnReferences()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.FirstCell()
            .SetValue(20)
            .CellBelow()
            .SetValue("AB")
            .CellBelow()
            .SetFormulaA1("DATE(2019,1,1)");

        // Calling cell is 1st row, so formula should return A1
        ws.Cell("B1").SetFormulaA1("CONCATENATE(A1:A3)");
        ClassicAssert.AreEqual("20", ws.Cell("B1").Value);

        // Calling cell is 2nd row, so formula should return A2
        ws.Cell("B2").SetFormulaA1("CONCATENATE(A1:A3)");
        ClassicAssert.AreEqual("AB", ws.Cell("B2").Value);

        // Calling cell is 3rd row, so formula should return A3's textual representation
        ws.Cell("B3").SetFormulaA1("CONCATENATE(A1:A3)");
        ClassicAssert.AreEqual("43466", ws.Cell("B3").Value);

        // Calling cell doesn't share row with any cell in parameter range.
        ws.Cell("A4").SetFormulaA1("CONCATENATE(A1:A3)");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell("A4").Value);
    }

    [Test]
    public void DollarCoercion() =>
        // Empty string is not coercible to number
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("DOLLAR(\"\", 3)")
        );

    // en-US culture differs between .NET Fx and Core for negative currency -> no test for negative
    [Test]
    [Arguments(123.54, 3, "$123.540")]
    [Arguments(123.54, 3.9, "$123.540")]
    [Arguments(1234.567, 2, "$1,234.57")]
    [Arguments(1250, -2, "$1,300")]
    [Arguments(1, -1E+100, "$0")]
    public void DollarEn(double number, double decimals, string expected)
    {
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual(expected, wb.Evaluate($"DOLLAR({number},{decimals})").GetText());
    }

    [Test]
    [Culture("cs-CZ")]
    [Arguments(123.54, 3, "123,540 Kč")]
    [Arguments(-1234.567, 4, "-1 234,5670 Kč")]
    [Arguments(-1250, -2, "-1 300 Kč")]
    public void DollarCs(double number, double decimals, string expected)
    {
        using XLWorkbook wb = new();
        string formula =
            $"DOLLAR({number.ToString(CultureInfo.InvariantCulture)},{decimals.ToString(CultureInfo.InvariantCulture)})";
        ClassicAssert.AreEqual(expected, wb.Evaluate(formula).GetText());
    }

    [Test]
    [Culture("de-DE")]
    [Arguments(1234.567, 2, "1.234,57 €")]
    [Arguments(1234.567, -2, "1.200 €")]
    [Arguments(-1234.567, 4, "-1.234,5670 €")]
    public void DollarDe(double number, double decimals, string expected)
    {
        using XLWorkbook wb = new();
        string formula =
            $"DOLLAR({number.ToString(CultureInfo.InvariantCulture)},{decimals.ToString(CultureInfo.InvariantCulture)})";
        ClassicAssert.AreEqual(expected, wb.Evaluate(formula).GetText());
    }

    [Test]
    public void DollarUsesTwoDecimalPlacesByDefault()
    {
        using XLWorkbook wb = new();
        XLCellValue actual = wb.Evaluate("DOLLAR(123.543)");
        ClassicAssert.AreEqual("$123.54", actual);
    }

    [Test]
    public void DollarCanHaveAtMost127DecimalPlaces()
    {
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual("$1." + new string('0', 99), wb.Evaluate("DOLLAR(1,99)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("DOLLAR(1,128)"));
    }

    [Test]
    public void ExactEmptyInputString()
    {
        object actual = XLWorkbook.EvaluateExpr(@"Exact("""", """")");
        ClassicAssert.AreEqual(true, actual);
    }

    [Test]
    public void ExactValue()
    {
        object actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""asdf"")");
        ClassicAssert.AreEqual(true, actual);

        actual = XLWorkbook.EvaluateExpr(@"Exact(""asdf"", ""ASDF"")");
        ClassicAssert.AreEqual(false, actual);

        actual = XLWorkbook.EvaluateExpr(@"Exact(123, 123)");
        ClassicAssert.AreEqual(true, actual);

        actual = XLWorkbook.EvaluateExpr(@"Exact(321, 123)");
        ClassicAssert.AreEqual(false, actual);
    }

    [Test]
    public void FindEmptyPatternAndEmptyText()
    {
        // Different behavior from SEARCH
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr(@"FIND("""", """")"));

        ClassicAssert.AreEqual(2, XLWorkbook.EvaluateExpr(@"FIND("""", ""a"", 2)"));
    }

    [Test]
    public void FindEmptySearchPatternReturnsStartOfText() =>
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr(@"FIND("""", ""asdf"")"));

    [Test]
    public void FindLooksOnlyFromStartPositionOnward() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"FIND(""This"", ""This is some text"", 2)")
        );

    [Test]
    public void FindStartPositionTooLarge() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"FIND(""abc"", ""abcdef"", 10)")
        );

    [Test]
    public void FindStartPositionTooSmall() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"FIND(""text"", ""This is some text"", 0)")
        );

    [Test]
    public void FindEmptySearchedTextReturnsError() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"FIND(""abc"", """")")
        );

    [Test]
    public void FindStringNotFound() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"FIND(""123"", ""asdf"")")
        );

    [Test]
    public void FindCaseSensitiveStringNotFound() =>
        // Find is case-sensitive
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"FIND(""excel"", ""Microsoft Excel 2010"")")
        );

    [Test]
    public void FindValue()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"FIND(""Tuesday"", ""Today is Tuesday"")");
        ClassicAssert.AreEqual(10, actual);

        // Doesnt support wildcards
        actual = XLWorkbook.EvaluateExpr(@"FIND(""T*y"", ""Today is Tuesday"")");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, actual);
    }

    [Test]
    public void FindArgumentsAreConvertedToExpectedTypes()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"FIND(1.2, ""A1.2B"")");
        ClassicAssert.AreEqual(2, actual);

        actual = XLWorkbook.EvaluateExpr(@"FIND(TRUE, ""ATRUE"")");
        ClassicAssert.AreEqual(2, actual);

        actual = XLWorkbook.EvaluateExpr(@"FIND(23, 1.2345)");
        ClassicAssert.AreEqual(3, actual);

        actual = XLWorkbook.EvaluateExpr(@"FIND(""a"", ""aaaaa"", ""2 1/2"")");
        ClassicAssert.AreEqual(2, actual);
    }

    [Test]
    public void FindErrorArgumentsReturnTheError()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"FIND(#N/A, ""a"")");
        ClassicAssert.AreEqual(XLError.NoValueAvailable, actual);

        actual = XLWorkbook.EvaluateExpr(@"FIND("""", #N/A)");
        ClassicAssert.AreEqual(XLError.NoValueAvailable, actual);

        actual = XLWorkbook.EvaluateExpr(@"FIND(""a"", ""a"", #N/A)");
        ClassicAssert.AreEqual(XLError.NoValueAvailable, actual);
    }

    [Test]
    public void FixedCoercion()
    {
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("""FIXED("asdf")"""));
        ClassicAssert.AreEqual("1234.0", wb.Evaluate("""FIXED(1234,1,"TRUE")"""));
        ClassicAssert.AreEqual("1,234.0", wb.Evaluate("""FIXED(1234,1,"FALSE")"""));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("""FIXED(1234,1,"0")"""));
    }

    [Test]
    public void FixedExamples()
    {
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual("1,234,567.00", wb.Evaluate("FIXED(1234567)"));
        ClassicAssert.AreEqual("1234567.5556", wb.Evaluate("FIXED(1234567.555555,4,TRUE)"));
        ClassicAssert.AreEqual("0.5555550000", wb.Evaluate("FIXED(.555555,10)"));
        ClassicAssert.AreEqual("1,235,000", wb.Evaluate("FIXED(1234567,-3)"));
    }

    [Test]
    public void FixedEn()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("FIXED(17300.67,4)");
        ClassicAssert.AreEqual("17,300.6700", actual);

        actual = XLWorkbook.EvaluateExpr("FIXED(17300.67,2,TRUE)");
        ClassicAssert.AreEqual("17300.67", actual);

        actual = XLWorkbook.EvaluateExpr("FIXED(17300.67)");
        ClassicAssert.AreEqual("17,300.67", actual);

        actual = XLWorkbook.EvaluateExpr("FIXED(1,-1E+300)");
        ClassicAssert.AreEqual("0", actual);
    }

    [Test]
    [Culture("cs-CZ")]
    public void FixedCs()
    {
        using XLWorkbook wb = new();
        XLCellValue actual = wb.Evaluate("FIXED(17300.67,4)");
        ClassicAssert.AreEqual("17 300,6700", actual);

        actual = wb.Evaluate("FIXED(17300.67,2,TRUE)");
        ClassicAssert.AreEqual("17300,67", actual);

        actual = wb.Evaluate("FIXED(17300.67)");
        ClassicAssert.AreEqual("17 300,67", actual);
    }

    [Test]
    public void FixedCanHaveAtMost127DecimalPlaces()
    {
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual("1." + new string('0', 99), wb.Evaluate("FIXED(1,99)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, wb.Evaluate("FIXED(1,128)"));
    }

    [Test]
    public void LeftReturnsWholeTextWhenRequestedLengthIsGreaterThanTextLength()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"LEFT(""ABC"", 5)");
        ClassicAssert.AreEqual("ABC", actual);
    }

    [Test]
    public void LeftTakesOneCharacterByDefault()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""LEFT("ABC")""");
        ClassicAssert.AreEqual("A", actual);
    }

    [Test]
    public void LeftReturnsErrorOnNegativeNumberOfChars() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""LEFT("ABC", -1)""")
        );

    [Test]
    public void LeftReturnsEmptyStringOnEmptyInput()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""LEFT("")""");
        ClassicAssert.AreEqual("", actual);
    }

    [Test]
    [Arguments("ABC", 2, "AB")]
    [Arguments("ABC", 2.9, "AB")]
    [Arguments("ABC", 3, "ABC")]
    [Arguments("\uD83D\uDC69Z", 1, "\uD83D\uDC69")] // Paired surrogate
    [Arguments("\uD83D\uDC69Z", 2, "\uD83D\uDC69Z")] // Paired surrogate
    public void LeftTakesSpecifiedNumberOfCharacters(string text, double numChars, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"""LEFT("{text}", {numChars})""").GetText()
        );
    }

    [Test]
    [Arguments("", 0)]
    [Arguments("word", 4)]
    [Arguments("A\r\n", 3)]
    [Arguments("H", 1)]
    [Arguments("\ud83d\ude0a", 2)] // Smile emoji
    [Arguments("Smile: \ud83d\ude0a!", 10)] // Smile emoji
    public void LenReturnsNumberOfCodeUnits(string text, double expected)
    {
        ClassicAssert.AreEqual(expected, XLWorkbook.EvaluateExpr($"""LEN("{text}")""").GetNumber());
    }

    [Test]
    [Arguments("", "")]
    [Arguments("ABC", "abc")]
    [Arguments("Intelligence 2.0!", "intelligence 2.0!")]
    [Arguments("ͶꝎＫǢ", "ͷꝏｋǣ")] // Converts even non-latin chars
    [Arguments("Σ SUM Σ end Σ", "σ sum σ end ς")] // Bug for bug behavior of Excel. Σ at the end is turned to ς
    public void LowerEn(string text, string expected)
    {
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual(expected, wb.Evaluate($"""LOWER("{text}")""").GetText());
    }

    [Test]
    [Culture("tr-TR")]
    [Arguments("INTELLIGENCE 2.0!", "ıntellıgence 2.0!")] // Turkey converts I to i without dot
    [Arguments("ΣΣΣΣ", "σσσς")]
    public void LowerTr(string text, string expected)
    {
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual(expected, wb.Evaluate($"""LOWER("{text}")""").GetText());
    }

    [Test]
    public void MidReturnsRestOfTextWhenEndIsOutOfTextBounds()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""MID("ABC",1,5)""");
        ClassicAssert.AreEqual("ABC", actual);
    }

    [Test]
    public void MidWhenStartIsAfterEndOfTextReturnEmptyString()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""MID("ABC",5,5)""");
        ClassicAssert.AreEqual("", actual);
    }

    [Test]
    [Arguments(0.9)]
    [Arguments(0)]
    [Arguments(-5)]
    [Arguments(int.MaxValue + 1d)]
    [Arguments(int.MaxValue + 5d)]
    public void MidStartMustBeAtLeastOneAndAtMostMaxInt(double start)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"""MID("ABC",{start},1)""");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, actual);
    }

    [Test]
    [Arguments(-0.1)]
    [Arguments(-5)]
    [Arguments(int.MaxValue + 1d)]
    [Arguments(int.MaxValue + 5d)]
    public void MidLengthMustBeAtLeastZeroAndAtMostMaxInt(double length)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"""MID("ABC",1,{length})""");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, actual);
    }

    [Test]
    [Arguments("", 1, 1, "")]
    [Arguments("ABC", 2, 2, "BC")]
    [Arguments("ABC", 2, 0, "")]
    [Arguments("ABC", 3, 5, "")]
    [Arguments(@"abcdef", 3, 2, "cd")]
    [Arguments(@"abcdef", 4, 5, "def")]
    public void MidReturnsSubstring(string text, double start, double length, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"""MID("{text}",{start},{length})""").GetText()
        );
    }

    [Test]
    public void MidUsesCodeUnits()
    {
        // MID returns unpaired surrogates
        ClassicAssert.AreEqual("😊\uD83D", XLWorkbook.EvaluateExpr("""MID("😊😊😊",1,3)"""));
        ClassicAssert.AreEqual("😊😊", XLWorkbook.EvaluateExpr("""MID("😊😊😊",1,4)"""));
        ClassicAssert.AreEqual("\uDE0A😊\uD83D", XLWorkbook.EvaluateExpr("""MID("😊😊😊",2,4)"""));
        ClassicAssert.AreEqual(3, XLWorkbook.EvaluateExpr("""LEN(MID("😊😊😊",1,3))"""));
    }

    [Test]
    [Arguments("", 0d)]
    [Arguments("+ 1", 1d)]
    [Arguments("+1", 1d)]
    [Arguments("+1.23", 1.23)]
    [Arguments("- 1.23", -1.23)]
    [Arguments(" - 0 1 2 . 3 4 ", -12.34)]
    [Arguments(" - 0 \t1\t2\r .\n3 4 ", -12.34)]
    [Arguments(".1", 0.1)]
    [Arguments("-.1", -0.1)]
    [Arguments("1.234567890E+307", 1.234567890E+307)]
    [Arguments("1.234567890E-307", 1.234567890E-307d)]
    [Arguments("1.234567890E-309", 0d)]
    [Arguments("-1.234567890E-307", -1.234567890E-307d)]
    [Arguments(".99999999999999", 0.99999999999999)]
    [Arguments("1,23,4", 1234)]
    [Arguments("1,234,56", 123456)]
    [Arguments("1e-308", 0)]
    [Arguments("-1e-308", 0)]
    [Arguments("75825%", 758.25)]
    [Arguments("75825%%", 7.5825)]
    [Arguments("(56.4)", -56.4)]
    [Arguments("(128)%", -1.28)]
    public void NumberValueConvertsTextToNumber(string text, double expectedResult)
    {
        double actual = (double)XLWorkbook.EvaluateExprCurrent($"NUMBERVALUE(\"{text}\")");
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    [Culture("de-DE")]
    public void NumberValueTakesSeparatorsFromCurrentCulture()
    {
        double actual = (double)XLWorkbook.EvaluateExprCurrent("NUMBERVALUE(\"10.0.00.0,25\")");
        ClassicAssert.AreEqual(100000.25, actual);
    }

    [Test]
    [Arguments("1,234.56", ".", ",", 1234.56d)]
    [Arguments("1.234,56", ",", ".", 1234.56d)]
    [Arguments("1.234,56", ",ABC", ".DEF", 1234.56d)] // Only first char of separators is used
    public void NumberValueOptionalParametersCanSetDecimalAndGroupSeparators(
        string text,
        string @decimal,
        string group,
        double expectedResult
    )
    {
        double actual = (double)
            XLWorkbook.EvaluateExpr($"NUMBERVALUE(\"{text}\",\"{@decimal}\",\"{group}\")");
        ClassicAssert.AreEqual(expectedResult, actual);
    }

    [Test]
    [Arguments("NUMBERVALUE(\"123.45\", \".\", \".\")")] // Group separator same as decimal separator
    [Arguments("NUMBERVALUE(\"1.234.5\")")] // Two decimal separators
    [Arguments("NUMBERVALUE(\"1.234,5\")")] // Decimal separator before group separator
    [Arguments("NUMBERVALUE(\"12;34\")")] // Illegal character
    [Arguments("NUMBERVALUE(\"--1\")")] // Two minuses
    [Arguments("NUMBERVALUE(\"1.234567890E+308\")")] // Too large
    [Arguments("NUMBERVALUE(\"-1.234567890E+308\")")] // Too large (negative)
    [Arguments("NUMBERVALUE(\"1.234567890E-310\")")] // Too tiny
    [Arguments("NUMBERVALUE(\"-1.234567890E-310\")")] // Too tiny (negative)
    [Arguments("NUMBERVALUE(\"1\",\".\",\"\")")] // Empty group separator
    [Arguments("NUMBERVALUE(\"1\",\"\",\",\")")] // Empty decimal separators
    public void NumberValueReturnsErrorOnUnparsableTextsOutOfRange(string expression) =>
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(expression));

    [Test]
    [Arguments("", "")]
    [Arguments("12aBC d123aD#$%sd^", "12Abc D123Ad#$%Sd^")]
    [Arguments("this is a TITLE", "This Is A Title")]
    [Arguments("2-way street", "2-Way Street")]
    [Arguments("76BudGet", "76Budget")]
    [Arguments("my name is francois botha", "My Name Is Francois Botha")]
    [Arguments("\ud83a\udd32", "\ud83a\udd32")] // U+1E932 has uppercase variant, but nothing changes, because PROPER uses code units
    public void ProperUpperCasesFirstLetterAndLowerCasesNextLetters(string text, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"""PROPER("{text}")""").GetText()
        );
    }

    [Test]
    [Arguments(1, 1)]
    [Arguments(1, 0)]
    [Arguments(1, 10)]
    [Arguments(10, 1)]
    [Arguments(10, 10)]
    public void ReplaceBeyondLimitAppendsReplacement(int startPos, int length)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            $"""REPLACE("",{startPos},{length},"new text")"""
        );
        ClassicAssert.AreEqual("new text", actual);
    }

    [Test]
    [Arguments(
        "Here is some obsolete text to replace.",
        14,
        13,
        "new text",
        "Here is some new text to replace."
    )]
    [Arguments("ABC", 1, 2, "D", "DC")]
    [Arguments("ABC", 3, 1, "D", "ABD")]
    [Arguments("ABC", 3, 0, "D", @"ABDC")]
    [Arguments("ABC", 4, 1, "D", @"ABCD")]
    [Arguments("ABC", 4, 0, "D", @"ABCD")]
    [Arguments("ABC", 1, 3, "D", "D")]
    [Arguments("ABC", 2, 2, "D", "AD")]
    [Arguments("ABC", 2, 0, "D", @"ADBC")]
    [Arguments("ABC", 2, 3, "D", "AD")]
    [Arguments(@"abcdefghijk", 3, 4, "XY", @"abXYghijk")]
    [Arguments(@"abcdefghijk", 3, 1, "12345", @"ab12345defghijk")]
    [Arguments(@"abcdefghijk", 15, 4, "XY", @"abcdefghijkXY")]
    public void ReplaceReplacesValue(
        string text,
        double startPos,
        int length,
        string replacement,
        string expected
    )
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook
                .EvaluateExpr($"""REPLACE("{text}",{startPos},{length},"{replacement}")""")
                .GetText()
        );
    }

    [Test]
    public void ReplaceStartPositionMustBeFrom1To32767()
    {
        ClassicAssert.AreEqual(@"DABC", XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,0,"D")"""));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""REPLACE("ABC",0.9,0,"D")""")
        );
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""REPLACE("ABC",-1,0,"D")""")
        );
        ClassicAssert.AreEqual("D", XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,32767.9,"D")"""));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,32768,"D")""")
        );
    }

    [Test]
    public void ReplaceLengthMustBeFrom0To32767()
    {
        ClassicAssert.AreEqual("ABC", XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,0,"")"""));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""REPLACE("ABC",1,-0.1,"D")""")
        );
        ClassicAssert.AreEqual("D", XLWorkbook.EvaluateExpr("""REPLACE("ABC",1, 32767.9,"D")"""));
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""REPLACE("ABC",1, 32768,"D")""")
        );
    }

    [Test]
    public void ReptReturnsEmptyStringWhenTextIsEmptyString()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""REPT("",3)""");
        ClassicAssert.AreEqual("", actual);
    }

    [Test]
    [Arguments(-1)]
    [Arguments(-0.1)]
    [Arguments(2147483648)]
    public void ReptReturnsErrorWhenCountIsNegativeOrGreaterThanMaxInt(double count) =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr($"""REPT("",{count})""")
        );

    [Test]
    public void ReptLimitsOutputTextLengthTo32767()
    {
        ClassicAssert.AreEqual(
            new string('A', 32767),
            XLWorkbook.EvaluateExpr("""REPT("A",32767)""")
        );
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""REPT("A",32768)""")
        );
    }

    [Test]
    [Arguments("ABC", 3, @"ABCABCABC")]
    [Arguments("123", 2.5, "123123")]
    [Arguments("Francois", 0, "")]
    [Arguments("Francois Botha,", 3, "Francois Botha,Francois Botha,Francois Botha,")]
    public void ReptValue(string text, double count, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"""REPT("{text}",{count})""").GetText()
        );
    }

    [Test]
    [Arguments(5)]
    [Arguments(3)]
    public void RightReturnsWholeTextWhenRequestedLengthIsGreaterThanTextLength(int length)
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr($"""RIGHT("ABC",{length})""");
        ClassicAssert.AreEqual("ABC", actual);
    }

    [Test]
    public void RightTakesOneCharacterByDefault()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""RIGHT("ABC")""");
        ClassicAssert.AreEqual("C", actual);
    }

    [Test]
    public void RightReturnsErrorOnNegativeNumberOfChars() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr("""RIGHT("ABC",-1)""")
        );

    [Test]
    public void RightReturnsEmptyStringOnEmptyInput()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""RIGHT("")""");
        ClassicAssert.AreEqual("", actual);
    }

    [Test]
    [Arguments("ABC", 0, "")]
    [Arguments("ABC", 1, "C")]
    [Arguments("ABC", 2, "BC")]
    [Arguments("ABC", 3, "ABC")]
    [Arguments("ABC", 4, "ABC")]
    [Arguments("ABC", 2.9, "BC")]
    [Arguments("Z\uD83D\uDC69", 1, "\uD83D\uDC69")] // Smiley emoji
    [Arguments("\uD83D\uDC69Z", 2, "\uD83D\uDC69Z")]
    [Arguments("\uD83D\uDC69Z", 3, "\uD83D\uDC69Z")]
    public void RightTakesSpecifiedNumberOfCharacters(string text, double numChars, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"""RIGHT("{text}",{numChars})""").GetText()
        );
    }

    [Test]
    public void SearchEmptyPatternAndEmptyText() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"SEARCH("""", """")")
        );

    [Test]
    public void SearchEmptySearchPatternReturnsStartOfText()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SEARCH("""", ""asdf"")");
        ClassicAssert.AreEqual(1, actual);
    }

    [Test]
    public void SearchLooksOnlyFromStartPositionOnward() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"SEARCH(""This"", ""This is some text"", 2)")
        );

    [Test]
    public void SearchStartPositionTooLarge() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"SEARCH(""abc"", ""abcdef"", 10)")
        );

    [Test]
    public void SearchStartPositionTooSmall() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"SEARCH(""text"", ""This is some text"", 0)")
        );

    [Test]
    public void SearchEmptySearchedTextReturnsError() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"SEARCH(""abc"", """")")
        );

    [Test]
    public void SearchTextNotFound() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"SEARCH(""123"", ""asdf"")")
        );

    [Test]
    public void SearchWildcardStringNotFound() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"SEARCH(""soft?2010"", ""Microsoft Excel 2010"")")
        );

    // http://www.excel-easy.com/examples/find-vs-search.html
    [Test]
    public void SearchValue()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SEARCH(""Tuesday"", ""Today is Tuesday"")");
        ClassicAssert.AreEqual(10, actual);

        // The search is case-insensitive
        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""excel"", ""Microsoft Excel 2010"")");
        ClassicAssert.AreEqual(11, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""soft*2010"", ""Microsoft Excel 2010"")");
        ClassicAssert.AreEqual(6, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""Excel 20??"", ""Microsoft Excel 2010"")");
        ClassicAssert.AreEqual(11, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""text"", ""This is some text"", 14)");
        ClassicAssert.AreEqual(14, actual);
    }

    [Test]
    public void SearchTildeEscapesNextChar()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SEARCH(""~a~b~"", ""ab"")");
        ClassicAssert.AreEqual(1, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~*"", ""a*"")");
        ClassicAssert.AreEqual(1, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~*"", ""ab"")");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~?"", ""a?"")");
        ClassicAssert.AreEqual(1, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a~?"", ""ab"")");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, actual);
    }

    [Test]
    public void SearchArgumentsAreConvertedToExpectedTypes()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SEARCH(1.2, ""A1.2B"")");
        ClassicAssert.AreEqual(2, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(TRUE, ""ATRUE"")");
        ClassicAssert.AreEqual(2, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(23, 1.2345)");
        ClassicAssert.AreEqual(3, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a"", ""aaaaa"", ""2 1/2"")");
        ClassicAssert.AreEqual(2, actual);
    }

    [Test]
    public void SearchErrorArgumentsReturnTheError()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SEARCH(#N/A, ""a"")");
        ClassicAssert.AreEqual(XLError.NoValueAvailable, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH("""", #N/A)");
        ClassicAssert.AreEqual(XLError.NoValueAvailable, actual);

        actual = XLWorkbook.EvaluateExpr(@"SEARCH(""a"", ""a"", #N/A)");
        ClassicAssert.AreEqual(XLError.NoValueAvailable, actual);
    }

    [Test]
    public void SubstituteReplacesNThOccurence()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(
            @"SUBSTITUTE(""This is a Tuesday."", ""Tuesday"", ""Monday"")"
        );
        ClassicAssert.AreEqual("This is a Monday.", actual);

        actual = XLWorkbook.EvaluateExpr(
            @"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 1)"
        );
        ClassicAssert.AreEqual("This is a Monday. Next week also has a Tuesday.", actual);

        actual = XLWorkbook.EvaluateExpr(
            @"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", ""Monday"", 2)"
        );
        ClassicAssert.AreEqual("This is a Tuesday. Next week also has a Monday.", actual);

        actual = XLWorkbook.EvaluateExpr(
            @"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", """", ""Monday"")"
        );
        ClassicAssert.AreEqual("This is a Tuesday. Next week also has a Tuesday.", actual);

        actual = XLWorkbook.EvaluateExpr(
            @"SUBSTITUTE(""This is a Tuesday. Next week also has a Tuesday."", ""Tuesday"", """")"
        );
        ClassicAssert.AreEqual("This is a . Next week also has a .", actual);
    }

    [Test]
    public void SubstituteOnEmptyStringReturnsEmptyString()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE("""","""",""Monday"")");
        ClassicAssert.AreEqual("", actual);
    }

    [Test]
    public void SubstituteIsCaseSensitive()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""SUBSTITUTE("A","a","Z")""");
        ClassicAssert.AreEqual("A", actual);
    }

    [Test]
    public void SubstituteReturnsOriginalStringWhenOccurrenceIsNotFound()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABCABC"",""A"",""Z"",3)");
        ClassicAssert.AreEqual(@"ABCABC", actual);
    }

    [Test]
    public void SubstituteSearchesForEveryOccurence()
    {
        // AA is matches at every character, it doesn't skip
        XLCellValue actual = XLWorkbook.EvaluateExpr("""SUBSTITUTE("AAAAAAAA","AA","ZZ",3)""");
        ClassicAssert.AreEqual(@"AAZZAAAA", actual);
    }

    [Test]
    public void SubstituteOccurenceMustBeBetweenOneAndMaxInt()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"",0.9)");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, actual);

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"", 2147483646.9)");
        ClassicAssert.AreEqual("ABC", actual);

        actual = XLWorkbook.EvaluateExpr(@"SUBSTITUTE(""ABC"",""B"",""ZZ"", 2147483647)");
        ClassicAssert.AreEqual(XLError.IncompatibleValue, actual);
    }

    [Test]
    public void TReturnsEmptyStringOnNonText()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("T(TODAY())");
        ClassicAssert.AreEqual("", actual);

        actual = XLWorkbook.EvaluateExpr("T(IF(TRUE,,))");
        ClassicAssert.AreEqual("", actual);

        actual = XLWorkbook.EvaluateExpr("T(TRUE)");
        ClassicAssert.AreEqual("", actual);

        actual = XLWorkbook.EvaluateExpr("T(123)");
        ClassicAssert.AreEqual("", actual);
    }

    [Test]
    public void TPropagatesError() =>
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("T(#DIV/0!)"));

    [Test]
    public void TReturnsTextWhenValueIsText()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""T("asdf")""");
        ClassicAssert.AreEqual("asdf", actual);

        actual = XLWorkbook.EvaluateExpr("""T("")""");
        ClassicAssert.AreEqual("", actual);
    }

    [Test]
    public void TReturnsArrayOfResultsWhenArgumentIsArray()
    {
        const string formula = """T({"A",5,"B"})""";
        ClassicAssert.AreEqual(3, XLWorkbook.EvaluateExpr($"""COLUMNS({formula})"""));
        ClassicAssert.AreEqual(1, XLWorkbook.EvaluateExpr($"""ROWS({formula})"""));
        ClassicAssert.AreEqual("A", XLWorkbook.EvaluateExpr($"""INDEX({formula},1,1)"""));
        ClassicAssert.AreEqual("", XLWorkbook.EvaluateExpr($"""INDEX({formula},1,2)"""));
        ClassicAssert.AreEqual("B", XLWorkbook.EvaluateExpr($"""INDEX({formula},1,3)"""));

        // Array doesn't propagate single error, but returns errors in the array
        ClassicAssert.AreEqual("A", XLWorkbook.EvaluateExpr("""INDEX(T({"A",#REF!}),1,1)"""));
        ClassicAssert.AreEqual(
            XLError.CellReference,
            XLWorkbook.EvaluateExpr("""INDEX(T({"A",#REF!}),1,2)""")
        );
    }

    [Test]
    public void TReturnsTextOfFirstCellInReference()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("B3").Value = "ABC";
        ws.Cell("B4").Value = 10;
        ws.Cell("B5").Value = XLError.NoValueAvailable;

        ClassicAssert.AreEqual("ABC", ws.Evaluate("T(B3:B4)"));
        ClassicAssert.AreEqual(2, ws.Evaluate("TYPE(T(B3:B4))")); // Is text, not array

        ClassicAssert.AreEqual(string.Empty, ws.Evaluate("T(B4:C4)"));

        ClassicAssert.AreEqual(XLError.NoValueAvailable, ws.Evaluate("T(B5:C5)"));
    }

    [Test]
    public void TextReturnsEmptyStringOnEmptyString()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr(@"TEXT(1913415.93,"""")");
        ClassicAssert.AreEqual(string.Empty, actual);
    }

    [Test]
    [Arguments("DATE(2010, 1, 1)", "yyyy-MM-dd", "2010-01-01")]
    [Arguments("1469.07", "0,000,000.00", "0,001,469.07")]
    [Arguments("1913415.93", "#,000.00", "1,913,415.93")]
    [Arguments("2800", "$0.00", "$2800.00")]
    [Arguments("0.4", "0%", "40%")]
    [Arguments("DATE(2010, 1, 1)", "MMMM yyyy", "January 2010")]
    [Arguments("DATE(2010, 1, 1)", "M/d/y", "1/1/10")]
    [Arguments("1234.567", "$0.00", "$1234.57")]
    [Arguments(".125", "$0.0%", "$12.5%")]
    [Arguments("1234.567", "YYYY-MM-DD HH:MM:SS", "1903-05-18 13:36:28")] // Excel is one second off (29), but that is in the library
    [Arguments("\"0.0245\"", "00%", "02%")]
    public void TextFormatsNumber(string numberArg, string format, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($"TEXT({numberArg},\"{format}\")").GetText()
        );
    }

    [Test]
    [Arguments("\"211x\"", "211x")]
    [Arguments("true", "TRUE")]
    public void TextReturnsStringRepresentationOfNonNumbers(string valueArg, string expected)
    {
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr($@"TEXT({valueArg},""#00"")").GetText()
        );
    }

    [Test]
    [Arguments(2020, 11, 1, 9, 23, 11, "m/d/yyyy h:mm:ss", "11/1/2020 9:23:11")]
    [Arguments(2023, 7, 14, 2, 12, 3, "m/d/yyyy h:mm:ss", "7/14/2023 2:12:03")]
    [Arguments(2025, 10, 14, 2, 48, 55, "m/d/yyyy h:mm:ss", "10/14/2025 2:48:55")]
    [Arguments(2023, 2, 19, 22, 1, 38, "m/d/yyyy h:mm:ss", "2/19/2023 22:01:38")]
    [Arguments(2025, 12, 19, 19, 43, 58, "m/d/yyyy h:mm:ss", "12/19/2025 19:43:58")]
    [Arguments(2034, 11, 16, 1, 48, 9, "m/d/yyyy h:mm:ss", "11/16/2034 1:48:09")]
    [Arguments(2018, 12, 10, 11, 22, 42, "m/d/yyyy h:mm:ss", "12/10/2018 11:22:42")]
    public void TextFormatsSerialDates(
        int year,
        int months,
        int days,
        int hour,
        int minutes,
        int seconds,
        string format,
        string expected
    ) =>
        ClassicAssert.AreEqual(
            expected,
            XLWorkbook.EvaluateExpr(
                $@"TEXT(DATE({year},{months},{days}) + TIME({hour},{minutes},{seconds}),""{format}"")"
            )
        );

    [Test]
    public void TextPropagatesErrors() =>
        ClassicAssert.AreEqual(
            XLError.CellReference,
            XLWorkbook.EvaluateExpr(@"TEXT(#REF!,""#00"")")
        );

    [Test]
    [Arguments("TEXTJOIN(\",\",TRUE,A1:B2)", "A,B,D")]
    [Arguments("TEXTJOIN(\",\",FALSE,A1:B2)", "A,,B,D")]
    [Arguments("TEXTJOIN(\",\",FALSE,A1,A2,B1,B2)", "A,B,,D")]
    [Arguments("TEXTJOIN(\",\",FALSE,1)", "1")]
    [Arguments("TEXTJOIN(\",\", TRUE, A:A, B:B)", "A,B,D")]
    [Arguments("TEXTJOIN(\",\", TRUE, D1:E2)", "")]
    [Arguments("TEXTJOIN(\",\", FALSE, D1:E2)", ",,,")]
    [Arguments(
        "TEXTJOIN(\",\", FALSE, D1:D32768)",
        ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
    )]
    [Arguments("TEXTJOIN(0, FALSE, A1:B2)", "A00B0D")]
    [Arguments("TEXTJOIN(false, FALSE, A1:B2)", @"AFALSEFALSEBFALSED")]
    [Arguments("TEXTJOIN(\",\", 0, A1:B2)", "A,,B,D")]
    [Arguments("TEXTJOIN(\",\", 100, A1:B2)", "A,B,D")]
    [Arguments("TEXTJOIN(B2, FALSE, A1:B2)", @"ADDBDD")]
    [Arguments("TEXTJOIN(\",\", FALSE, 12345.67, DATE(2018, 10, 30))", "12345.67,43403")]
    [Arguments("TEXTJOIN(\",\", \"FALSE\", A1:B2)", "A,,B,D")]
    public void TextJoinJoinsArgumentsWithSpecifiedDelimiter(string formula, string expectedOutput)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "A";
        ws.Cell("A2").Value = "B";
        ws.Cell("B1").Value = "";
        ws.Cell("B2").Value = "D";

        ws.Cell("C1").FormulaA1 = formula;
        XLCellValue a = ws.Cell("C1").Value;

        ClassicAssert.AreEqual(expectedOutput, a);
    }

    [Test]
    [Arguments("TEXTJOIN(\",\", FALSE, D1:D32769)")]
    public void TextJoinOutputCanBeAtMost32767(string formula)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("C1").FormulaA1 = formula;

        // Excel actually returns #CALC!, but we don't have that error, mostly
        // because parser doesn't recognize it.
        ClassicAssert.AreEqual(XLError.IncompatibleValue, ws.Cell("C1").Value);
    }

    [Test]
    [Arguments("TEXTJOIN(\",\", \"Invalid\", \"Hello\", \"World\")")]
    public void TextJoinCoercion(string formula) =>
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(formula));

    [Test]
    [Arguments("", "")]
    [Arguments(" ", "")]
    [Arguments("    ", "")]
    [Arguments(" Break\r\n   Line   ", "Break\r\n Line")]
    [Arguments("non-whitespace-text", "non-whitespace-text")]
    [Arguments("white space text", "white space text")]
    [Arguments(" some text with padding   ", "some text with padding")]
    [Arguments(" \t  A  \t ", "\t A \t")]
    public void TrimTrimsSpacesAndRemovesMultiSpacesFromInsideText(string text, string expected)
    {
        ClassicAssert.AreEqual(expected, XLWorkbook.EvaluateExpr($"""TRIM("{text}")""").GetText());
    }

    [Test]
    public void UpperEmptyStringReturnsEmptyString() =>
        ClassicAssert.AreEqual("", XLWorkbook.EvaluateExpr("""UPPER("")"""));

    [Test]
    public void UpperConvertsTextToUpperCase()
    {
        XLCellValue actual = XLWorkbook.EvaluateExpr("""UPPER("AbCdEfG")""");
        ClassicAssert.AreEqual(@"ABCDEFG", actual);
    }

    [Culture("tr-TR")]
    [Test]
    public void UpperUsesWorkbookCulture()
    {
        // Türkiye converts i to İ, not I.
        using XLWorkbook wb = new();
        ClassicAssert.AreEqual("İNTELLİGENCE 2.0!", wb.Evaluate("""UPPER("intelligence 2.0!")"""));
    }

    [Test]
    public void ValueInputStringIsNotANumber() =>
        ClassicAssert.AreEqual(
            XLError.IncompatibleValue,
            XLWorkbook.EvaluateExpr(@"VALUE(""asdf"")")
        );

    [Test]
    public void ValueFromBlankIsZero()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(0d, ws.Evaluate("VALUE(A1)"));
    }

    [Test]
    public void ValueFromEmptyStringIsError() =>
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("VALUE(\"\")"));

    [Test]
    public void ValuePassingUnexpectedTypes()
    {
        ClassicAssert.AreEqual(14d, XLWorkbook.EvaluateExpr(@"VALUE(14)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"VALUE(TRUE)"));
        ClassicAssert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr(@"VALUE(FALSE)"));
        ClassicAssert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr(@"VALUE(#DIV/0!)"));
    }

    [Test]
    public void ValueValue()
    {
        using XLWorkbook wb = new();

        // Examples from spec
        ClassicAssert.AreEqual(123.456d, wb.Evaluate("VALUE(\"123.456\")"));
        ClassicAssert.AreEqual(1000d, wb.Evaluate("VALUE(\"$1,000\")"));
        ClassicAssert.AreEqual(
            new DateTime(2002, 3, 23).ToSerialDateTime(),
            wb.Evaluate("VALUE(\"23-Mar-2002\")")
        );
        ClassicAssert.AreEqual(
            0.188056d,
            (double)wb.Evaluate("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")"),
            0.000001d
        );
    }

    [Test]
    [Culture("cs-CZ")]
    public void ValueNonEnglish()
    {
        using XLWorkbook wb = new();

        // Examples from spec
        ClassicAssert.AreEqual(123.456d, wb.Evaluate("VALUE(\"123,456\")"));
        ClassicAssert.AreEqual(1000d, wb.Evaluate("VALUE(\"1 000 Kč\")"));
        ClassicAssert.AreEqual(37338d, wb.Evaluate("VALUE(\"23-bře-2002\")"));
        ClassicAssert.AreEqual(
            0.188056d,
            (double)wb.Evaluate("VALUE(\"16:48:00\")-VALUE(\"12:17:12\")"),
            0.000001d
        );

        // Various number/currency formats
        ClassicAssert.AreEqual(-1d, wb.Evaluate("VALUE(\"(1)\")"));
        ClassicAssert.AreEqual(-1d, wb.Evaluate("VALUE(\"(100%)\")"));
        ClassicAssert.AreEqual(-1d, wb.Evaluate("VALUE(\"(100%)\")"));
        ClassicAssert.AreEqual(-15d, wb.Evaluate("VALUE(\"(1,5e1 Kč)\")"));
        ClassicAssert.AreEqual(-15d, wb.Evaluate("VALUE(\"(1,5e3%)\")"));
        ClassicAssert.AreEqual(-15d, wb.Evaluate("VALUE(\"(1,5e3)%\")"));

        double expectedSerialDate = new DateTime(2022, 3, 5).ToSerialDateTime();
        ClassicAssert.AreEqual(expectedSerialDate, wb.Evaluate("VALUE(\"5-březen-22\")"));
        ClassicAssert.AreEqual(expectedSerialDate, wb.Evaluate("VALUE(\"05.03.2022\")"));
        ClassicAssert.AreEqual(
            new DateTime(DateTime.Now.Year, 3, 5).ToSerialDateTime(),
            wb.Evaluate("VALUE(\"5-březen\")")
        );
    }
}

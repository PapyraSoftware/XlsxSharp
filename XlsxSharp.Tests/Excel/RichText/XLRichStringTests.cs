using XlsxSharp.Excel;
using XlsxSharp.Excel.RichText;

namespace XlsxSharp.Tests.Excel.RichText;

/// <summary>
///     This is a test class for XLRichStringTests and is intended
///     to contain all XLRichStringTests Unit Tests
/// </summary>
public class XlRichStringTests
{
    [Test]
    public void AccessRichTextTest1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        cell.CreateRichText().AddText("12");

        IXLRichText richText = cell.GetRichText();

        ClassicAssert.AreEqual("12", richText.ToString());

        richText.AddText("34");

        ClassicAssert.AreEqual("1234", cell.GetText());
    }

    /// <summary>
    ///     A test for AddText
    /// </summary>
    [Test]
    public void AddTextTest1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        IXLRichText richString = cell.CreateRichText();

        string text = "Hello";
        richString.AddText(text).SetBold().SetFontColor(XLColor.Red);

        ClassicAssert.AreEqual(cell.GetText(), text);
        ClassicAssert.AreEqual(cell.GetRichText().First().Bold, true);
        ClassicAssert.AreEqual(cell.GetRichText().First().FontColor, XLColor.Red);

        ClassicAssert.AreEqual(1, richString.Count);

        richString.AddText("World");
        ClassicAssert.AreEqual(
            richString.First().Text,
            text,
            "Item in collection is not the same as the one returned"
        );
    }

    [Test]
    public void AddTextTest2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        int number = 123;

        cell.SetValue(number).Style.Font.SetBold().Font.SetFontColor(XLColor.Red);

        string text = number.ToString();

        ClassicAssert.AreEqual(cell.GetRichText().ToString(), text);
        ClassicAssert.AreEqual(cell.GetRichText().First().Bold, true);
        ClassicAssert.AreEqual(cell.GetRichText().First().FontColor, XLColor.Red);

        ClassicAssert.AreEqual(1, cell.GetRichText().Count);

        cell.GetRichText().AddText("World");
        ClassicAssert.AreEqual(
            cell.GetRichText().First().Text,
            text,
            "Item in collection is not the same as the one returned"
        );
    }

    [Test]
    public void AddTextTest3()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        int number = 123;
        cell.Value = number;
        cell.Style.Font.SetBold().Font.SetFontColor(XLColor.Red);

        string text = number.ToString();

        ClassicAssert.AreEqual(cell.GetRichText().ToString(), text);
        ClassicAssert.AreEqual(cell.GetRichText().First().Bold, true);
        ClassicAssert.AreEqual(cell.GetRichText().First().FontColor, XLColor.Red);

        ClassicAssert.AreEqual(1, cell.GetRichText().Count);

        cell.GetRichText().AddText("World");
        ClassicAssert.AreEqual(
            cell.GetRichText().First().Text,
            text,
            "Item in collection is not the same as the one returned"
        );
    }

    /// <summary>
    ///     A test for Clear
    /// </summary>
    [Test]
    public void ClearTest()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText(" ");
        richString.AddText("World!");

        richString.ClearText();
        string expected = string.Empty;
        string actual = richString.ToString();
        ClassicAssert.AreEqual(expected, actual);

        ClassicAssert.AreEqual(0, richString.Count);
    }

    [Test]
    public void CountTest()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText(" ");
        richString.AddText("World!");

        ClassicAssert.AreEqual(3, richString.Count);
    }

    [Test]
    public void HasRichTextTest1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLCell cell = ws.Cell(1, 1);
        cell.GetRichText().AddText("123");

        ClassicAssert.AreEqual(true, cell.HasRichText);

        cell.Value = "123";

        ClassicAssert.AreEqual(false, cell.HasRichText);

        cell.GetRichText().AddText("123");

        ClassicAssert.AreEqual(true, cell.HasRichText);

        cell.Value = 123;

        ClassicAssert.AreEqual(false, cell.HasRichText);

        cell.GetRichText().AddText("123");

        ClassicAssert.AreEqual(true, cell.HasRichText);

        cell.SetValue("123");

        ClassicAssert.AreEqual(false, cell.HasRichText);
    }

    /// <summary>
    ///     A test for Characters
    /// </summary>
    [Test]
    public void SubstringAllFromOneString()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        IXLFormattedText<IXLRichText> actual = richString.Substring(0);

        ClassicAssert.AreEqual(richString.First(), actual.First());

        ClassicAssert.AreEqual(1, actual.Count);

        actual.First().SetBold();

        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().First().Bold);
    }

    [Test]
    public void SubstringAllFromThreeStrings()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        IXLFormattedText<IXLRichText> actual = richString.Substring(0);

        ClassicAssert.AreEqual(richString.ElementAt(0), actual.ElementAt(0));
        ClassicAssert.AreEqual(richString.ElementAt(1), actual.ElementAt(1));
        ClassicAssert.AreEqual(richString.ElementAt(2), actual.ElementAt(2));

        ClassicAssert.AreEqual(3, actual.Count);
        ClassicAssert.AreEqual(3, richString.Count);

        actual.First().SetBold();

        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().First().Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().Last().Bold);
    }

    [Test]
    public void SubstringFromOneStringEnd()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        IXLFormattedText<IXLRichText> actual = richString.Substring(2);

        ClassicAssert.AreEqual(1, actual.Count); // substring was in one piece

        ClassicAssert.AreEqual(2, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("llo", actual.First().Text);

        ClassicAssert.AreEqual("He", richString.First().Text);
        ClassicAssert.AreEqual("llo", richString.Last().Text);

        actual.First().SetBold();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Bold);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Bold);

        richString.Last().SetItalic();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Italic);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Italic);

        ClassicAssert.AreEqual(true, actual.First().Italic);

        richString.SetFontSize(20);

        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().First().FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().Last().FontSize);

        ClassicAssert.AreEqual(20, actual.First().FontSize);
    }

    [Test]
    public void SubstringFromOneStringMiddle()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        IXLFormattedText<IXLRichText> actual = richString.Substring(2, 2);

        ClassicAssert.AreEqual(1, actual.Count); // substring was in one piece

        ClassicAssert.AreEqual(3, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("ll", actual.First().Text);

        ClassicAssert.AreEqual("He", richString.First().Text);
        ClassicAssert.AreEqual("ll", richString.ElementAt(1).Text);
        ClassicAssert.AreEqual("o", richString.Last().Text);

        actual.First().SetBold();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Bold);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().Last().Bold);

        richString.Last().SetItalic();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Italic);

        ClassicAssert.AreEqual(false, actual.First().Italic);

        richString.SetFontSize(20);

        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().First().FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().Last().FontSize);

        ClassicAssert.AreEqual(20, actual.First().FontSize);
    }

    [Test]
    public void SubstringFromOneStringStart()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        IXLFormattedText<IXLRichText> actual = richString.Substring(0, 2);

        ClassicAssert.AreEqual(1, actual.Count); // substring was in one piece

        ClassicAssert.AreEqual(2, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("He", actual.First().Text);

        ClassicAssert.AreEqual("He", richString.First().Text);
        ClassicAssert.AreEqual("llo", richString.Last().Text);

        actual.First().SetBold();

        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().First().Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().Last().Bold);

        richString.Last().SetItalic();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().First().Italic);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().Last().Italic);

        ClassicAssert.AreEqual(false, actual.First().Italic);

        richString.SetFontSize(20);

        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().First().FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().Last().FontSize);

        ClassicAssert.AreEqual(20, actual.First().FontSize);
    }

    [Test]
    public void SubstringFromThreeStringsEnd1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        IXLFormattedText<IXLRichText> actual = richString.Substring(21);

        ClassicAssert.AreEqual(1, actual.Count); // substring was in one piece

        ClassicAssert.AreEqual(4, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("bors!", actual.First().Text);

        ClassicAssert.AreEqual("Good Morning", richString.ElementAt(0).Text);
        ClassicAssert.AreEqual(" my ", richString.ElementAt(1).Text);
        ClassicAssert.AreEqual("neigh", richString.ElementAt(2).Text);
        ClassicAssert.AreEqual("bors!", richString.ElementAt(3).Text);

        actual.First().SetBold();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

        richString.Last().SetItalic();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

        ClassicAssert.AreEqual(true, actual.First().Italic);

        richString.SetFontSize(20);

        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

        ClassicAssert.AreEqual(20, actual.First().FontSize);
    }

    [Test]
    public void SubstringFromThreeStringsEnd2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        IXLFormattedText<IXLRichText> actual = richString.Substring(13);

        ClassicAssert.AreEqual(2, actual.Count);

        ClassicAssert.AreEqual(4, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("my ", actual.ElementAt(0).Text);
        ClassicAssert.AreEqual("neighbors!", actual.ElementAt(1).Text);

        ClassicAssert.AreEqual("Good Morning", richString.ElementAt(0).Text);
        ClassicAssert.AreEqual(" ", richString.ElementAt(1).Text);
        ClassicAssert.AreEqual("my ", richString.ElementAt(2).Text);
        ClassicAssert.AreEqual("neighbors!", richString.ElementAt(3).Text);

        actual.ElementAt(1).SetBold();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

        richString.Last().SetItalic();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

        ClassicAssert.AreEqual(false, actual.ElementAt(0).Italic);
        ClassicAssert.AreEqual(true, actual.ElementAt(1).Italic);

        richString.SetFontSize(20);

        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

        ClassicAssert.AreEqual(20, actual.ElementAt(0).FontSize);
        ClassicAssert.AreEqual(20, actual.ElementAt(1).FontSize);
    }

    [Test]
    public void SubstringFromThreeStringsMid1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        IXLFormattedText<IXLRichText> actual = richString.Substring(5, 10);

        ClassicAssert.AreEqual(2, actual.Count);

        ClassicAssert.AreEqual(5, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("Morning", actual.ElementAt(0).Text);
        ClassicAssert.AreEqual(" my", actual.ElementAt(1).Text);

        ClassicAssert.AreEqual("Good ", richString.ElementAt(0).Text);
        ClassicAssert.AreEqual("Morning", richString.ElementAt(1).Text);
        ClassicAssert.AreEqual(" my", richString.ElementAt(2).Text);
        ClassicAssert.AreEqual(" ", richString.ElementAt(3).Text);
        ClassicAssert.AreEqual("neighbors!", richString.ElementAt(4).Text);
    }

    [Test]
    public void SubstringFromThreeStringsMid2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        IXLFormattedText<IXLRichText> actual = richString.Substring(5, 15);

        ClassicAssert.AreEqual(3, actual.Count);

        ClassicAssert.AreEqual(5, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("Morning", actual.ElementAt(0).Text);
        ClassicAssert.AreEqual(" my ", actual.ElementAt(1).Text);
        ClassicAssert.AreEqual("neig", actual.ElementAt(2).Text);

        ClassicAssert.AreEqual("Good ", richString.ElementAt(0).Text);
        ClassicAssert.AreEqual("Morning", richString.ElementAt(1).Text);
        ClassicAssert.AreEqual(" my ", richString.ElementAt(2).Text);
        ClassicAssert.AreEqual("neig", richString.ElementAt(3).Text);
        ClassicAssert.AreEqual("hbors!", richString.ElementAt(4).Text);
    }

    [Test]
    public void SubstringFromThreeStringsStart1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        IXLFormattedText<IXLRichText> actual = richString.Substring(0, 4);

        ClassicAssert.AreEqual(1, actual.Count); // substring was in one piece

        ClassicAssert.AreEqual(4, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("Good", actual.First().Text);

        ClassicAssert.AreEqual("Good", richString.ElementAt(0).Text);
        ClassicAssert.AreEqual(" Morning", richString.ElementAt(1).Text);
        ClassicAssert.AreEqual(" my ", richString.ElementAt(2).Text);
        ClassicAssert.AreEqual("neighbors!", richString.ElementAt(3).Text);

        actual.First().SetBold();

        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

        richString.First().SetItalic();

        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

        ClassicAssert.AreEqual(true, actual.First().Italic);

        richString.SetFontSize(20);

        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

        ClassicAssert.AreEqual(20, actual.First().FontSize);
    }

    [Test]
    public void SubstringFromThreeStringsStart2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        IXLFormattedText<IXLRichText> actual = richString.Substring(0, 15);

        ClassicAssert.AreEqual(2, actual.Count);

        ClassicAssert.AreEqual(4, richString.Count); // The text was split because of the substring

        ClassicAssert.AreEqual("Good Morning", actual.ElementAt(0).Text);
        ClassicAssert.AreEqual(" my", actual.ElementAt(1).Text);

        ClassicAssert.AreEqual("Good Morning", richString.ElementAt(0).Text);
        ClassicAssert.AreEqual(" my", richString.ElementAt(1).Text);
        ClassicAssert.AreEqual(" ", richString.ElementAt(2).Text);
        ClassicAssert.AreEqual("neighbors!", richString.ElementAt(3).Text);

        actual.ElementAt(1).SetBold();

        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(0).Bold);
        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(1).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Bold);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Bold);

        richString.First().SetItalic();

        ClassicAssert.AreEqual(true, ws.Cell(1, 1).GetRichText().ElementAt(0).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(1).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(2).Italic);
        ClassicAssert.AreEqual(false, ws.Cell(1, 1).GetRichText().ElementAt(3).Italic);

        ClassicAssert.AreEqual(true, actual.ElementAt(0).Italic);
        ClassicAssert.AreEqual(false, actual.ElementAt(1).Italic);

        richString.SetFontSize(20);

        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize);
        ClassicAssert.AreEqual(20, ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize);

        ClassicAssert.AreEqual(20, actual.ElementAt(0).FontSize);
        ClassicAssert.AreEqual(20, actual.ElementAt(1).FontSize);
    }

    [Test]
    public void SubstringIndexOutsideRange1()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        ClassicAssert.Throws<IndexOutOfRangeException>(() => richString.Substring(50));
    }

    [Test]
    public void SubstringIndexOutsideRange2()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText("World");

        ClassicAssert.Throws<IndexOutOfRangeException>(() => richString.Substring(50));
    }

    [Test]
    public void SubstringIndexOutsideRange3()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        ClassicAssert.Throws<IndexOutOfRangeException>(() => richString.Substring(1, 10));
    }

    [Test]
    public void SubstringIndexOutsideRange4()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText("World");

        ClassicAssert.Throws<IndexOutOfRangeException>(() => richString.Substring(5, 20));
    }

    [Test]
    public void CopyFromDoesCopy()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRichText original = ws.Cell(1, 1).GetRichText();
        original
            .AddText("Hello")
            .SetFontSize(15)
            .SetFontColor(XLColor.Red)
            .AddText("World")
            .SetFontSize(7)
            .SetFontColor(XLColor.Blue);

        IXLCell otherCell = ws.Cell(1, 2);
        IXLRichText otherRichText = otherCell.GetRichText();
        otherRichText.CopyFrom(original);

        ClassicAssert.AreEqual("HelloWorld", otherCell.Value);
        ClassicAssert.AreEqual(2, otherRichText.Count);
        ClassicAssert.AreEqual(XLColor.Red, otherRichText.First().FontColor);
        ClassicAssert.AreEqual(XLColor.Blue, otherRichText.Last().FontColor);
    }

    /// <summary>
    ///     A test for ToString
    /// </summary>
    [Test]
    public void ToStringTest()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRichText richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText(" ");
        richString.AddText("World");
        string expected = "Hello World";
        string actual = richString.ToString();
        ClassicAssert.AreEqual(expected, actual);

        richString.AddText("!");
        expected = "Hello World!";
        actual = richString.ToString();
        ClassicAssert.AreEqual(expected, actual);

        richString.ClearText();
        expected = string.Empty;
        actual = richString.ToString();
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    [Property("Description", "See #1361")]
    public void CanClearInlinedRichText()
    {
        using (MemoryStream outputStream = new())
        {
            using (
                Stream inputStream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(
                        @"Other\InlinedRichText\ChangeRichText\inputfile.xlsx"
                    )
                )
            )
            using (XLWorkbook workbook = new(inputStream))
            {
                workbook.Worksheets.First().Cell("A1").Value = "";
                workbook.SaveAs(outputStream);
            }

            using (XLWorkbook wb = new(outputStream))
            {
                ClassicAssert.AreEqual("", wb.Worksheets.First().Cell("A1").Value);
            }
        }
    }

    [Test]
    public void CanChangeInlinedRichText()
    {
        static void AssertRichText(IXLRichText richText)
        {
            ClassicAssert.IsNotNull(richText);
            ClassicAssert.IsTrue(richText.Any());
            ClassicAssert.AreEqual("3", richText.ElementAt(2).Text);
            ClassicAssert.AreEqual(XLColor.Red, richText.ElementAt(2).FontColor);
        }

        using (MemoryStream outputStream = new())
        {
            using (
                Stream inputStream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(
                        @"Other\InlinedRichText\ChangeRichText\inputfile.xlsx"
                    )
                )
            )
            using (XLWorkbook workbook = new(inputStream))
            {
                IXLRichText richText = workbook.Worksheets.First().Cell("A1").GetRichText();
                AssertRichText(richText);
                richText.AddText(" - changed");
                workbook.SaveAs(outputStream);
            }

            using (XLWorkbook wb = new(outputStream))
            {
                IXLCell cell = wb.Worksheets.First().Cell("A1");
                ClassicAssert.IsFalse(cell.ShareString);
                ClassicAssert.IsTrue(cell.HasRichText);
                IXLRichText rt = cell.GetRichText();
                ClassicAssert.AreEqual("Year (range: 3 yrs) - changed", rt.ToString());
                AssertRichText(rt);
            }
        }
    }

    [Test]
    public void ClearInlineRichTextWhenRelevant()
    {
        using MemoryStream ms = new();
        TestHelper.CreateAndCompare(
            () =>
            {
                using (XLWorkbook wb = new())
                {
                    IXLWorksheet ws = wb.AddWorksheet();
                    IXLCell cell = ws.FirstCell();

                    cell.GetRichText()
                        .AddText("Bold")
                        .SetBold()
                        .AddText(" and red")
                        .SetBold()
                        .SetFontColor(XLColor.Red);
                    cell.ShareString = false;

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                XLWorkbook wb2 = new(ms);
                {
                    IXLWorksheet ws = wb2.Worksheets.First();
                    IXLCell cell = ws.FirstCell();

                    cell.FormulaA1 = "1 + 2";
                    wb2.SaveAs(ms);
                }

                return wb2;
            },
            @"Other\InlinedRichText\ChangeRichTextToFormula\output.xlsx"
        );
    }

    [Test]
    public void RichTextChangesContentOfItsCell()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLCell cell = ws.Cell(1, 1);
        IXLRichText richText = cell.GetRichText();

        ClassicAssert.AreEqual(cell.Value, richText.Text);

        richText.AddText("Hello");
        ClassicAssert.AreEqual(cell.Value, "Hello");

        IXLRichString world = richText.AddText(" World");
        ClassicAssert.AreEqual(cell.Value, "Hello World");

        world.Text = " World!";
        ClassicAssert.AreEqual(cell.Value, "Hello World!");
        ClassicAssert.AreEqual(cell.GetRichText().Text, "Hello World!");

        richText.ClearText();
        ClassicAssert.AreEqual(cell.Value, string.Empty);
    }

    [Test]
    public void RemovedRichTextFromCellCantBeChanged()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLCell cell = ws.Cell(1, 1);
        IXLRichText richText = cell.GetRichText();
        cell.Value = 4;

        ClassicAssert.Throws<InvalidOperationException>(
            () => richText.AddText("Hello"),
            "The rich text isn't a content of a cell."
        );
    }

    [Test]
    public void MaintainWhitespaces()
    {
        const string textWithSpaces = "  元  気  ";
        const string phoneticsWithSpace = "  げ  ん  ";
        TestHelper.CreateSaveLoadAssert(
            wb =>
            {
                IXLWorksheet ws = wb.AddWorksheet();
                IXLCell richTextCell = ws.Cell(1, 1);
                IXLRichText richText = richTextCell.GetRichText();
                richText.AddText(textWithSpaces);
                richText.Phonetics.Add(phoneticsWithSpace, 2, 3);
            },
            wb =>
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLRichText richText = ws.Cell(1, 1).GetRichText();
                ClassicAssert.AreEqual(textWithSpaces, richText.First().Text);
                ClassicAssert.AreEqual(phoneticsWithSpace, richText.Phonetics.First().Text);
            }
        );
    }

    [Test]
    public void PreserveEndOfLineInXml() =>
        // When text run in a rich text contains end of line (regardless if CR, LF or CRLF),
        // the written element must be marked with xml:space="preserve". Excel would process
        // text differently (trim ect, see XML spec) and that means there would be a data
        // loss (trimmed ends of line). Another problem would be phonetic runs. They use indexes
        // to the text run, but if text would be trimmed, they might suddenly have out-of-bounds
        // values and Excel would try to repair the workbook.
        // The source files contains a text run with end of line at the start and end. It also
        // contains phonetic run for the kanji in the text that would be out-of-bounds if space
        // attribute there. The input is from Excel, output is by XlsxSharp. Output must contain
        // the space attribute.
        TestHelper.LoadSaveAndCompare(
            @"Other\RichText\kanji-with-new-line-input.xlsx",
            @"Other\RichText\kanji-with-new-line-output.xlsx"
        );
}

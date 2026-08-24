#nullable disable

using System.Text;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.PageSetup;

internal class XLHFItem : IXLHFItem
{
    internal readonly XLHeaderFooter HeaderFooter;

    public XLHFItem(XLHeaderFooter headerFooter) => this.HeaderFooter = headerFooter;

    public XLHFItem(XLHFItem defaultHFItem, XLHeaderFooter headerFooter)
        : this(headerFooter) => defaultHFItem.texts.ForEach(kp => this.texts.Add(kp.Key, kp.Value));

    private readonly Dictionary<XLHFOccurrence, List<XLHFText>> texts = new();

    public string GetText(XLHFOccurrence occurrence)
    {
        StringBuilder sb = new();
        if (this.texts.TryGetValue(occurrence, out List<XLHFText> hfTexts))
        {
            foreach (XLHFText hfText in hfTexts)
            {
                sb.Append(hfText.GetHFText(sb.ToString()));
            }
        }

        return sb.ToString();
    }

    public IXLRichString AddText(string text) => this.AddText(text, XLHFOccurrence.AllPages);

    public IXLRichString AddText(XLHFPredefinedText predefinedText) =>
        this.AddText(predefinedText, XLHFOccurrence.AllPages);

    public IXLRichString AddText(string text, XLHFOccurrence occurrence)
    {
        // TODO Styles: This doesn't update source when API object changes
        XLRichString richText = new(
            text,
            this.HeaderFooter.Worksheet.GetFormat().Font,
            this,
            this.HeaderFooter.Worksheet.Workbook.Styles,
            null
        );

        XLHFText hfText = new(richText, this);
        if (occurrence == XLHFOccurrence.AllPages)
        {
            this.AddTextToOccurrence(hfText, XLHFOccurrence.EvenPages);
            this.AddTextToOccurrence(hfText, XLHFOccurrence.FirstPage);
            this.AddTextToOccurrence(hfText, XLHFOccurrence.OddPages);
        }
        else
        {
            this.AddTextToOccurrence(hfText, occurrence);
        }

        return richText;
    }

    public IXLRichString AddNewLine() => this.AddText(Environment.NewLine);

    public IXLRichString AddImage(
        string imagePath,
        XLHFOccurrence occurrence = XLHFOccurrence.AllPages
    ) => throw new NotImplementedException();

    private void AddTextToOccurrence(XLHFText hfText, XLHFOccurrence occurrence)
    {
        if (this.texts.TryGetValue(occurrence, out List<XLHFText> hfTexts))
        {
            hfTexts.Add(hfText);
        }
        else
        {
            this.texts.Add(occurrence, [hfText]);
        }

        this.HeaderFooter.Changed = true;
    }

    public IXLRichString AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence)
    {
        string hfText;
        switch (predefinedText)
        {
            case XLHFPredefinedText.PageNumber:
                hfText = "&P";
                break;
            case XLHFPredefinedText.NumberOfPages:
                hfText = "&N";
                break;
            case XLHFPredefinedText.Date:
                hfText = "&D";
                break;
            case XLHFPredefinedText.Time:
                hfText = "&T";
                break;
            case XLHFPredefinedText.Path:
                hfText = "&Z";
                break;
            case XLHFPredefinedText.File:
                hfText = "&F";
                break;
            case XLHFPredefinedText.SheetName:
                hfText = "&A";
                break;
            case XLHFPredefinedText.FullPath:
                hfText = "&Z&F";
                break;
            default:
                throw new NotImplementedException();
        }
        return this.AddText(hfText, occurrence);
    }

    public void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
    {
        if (occurrence == XLHFOccurrence.AllPages)
        {
            this.ClearOccurrence(XLHFOccurrence.EvenPages);
            this.ClearOccurrence(XLHFOccurrence.FirstPage);
            this.ClearOccurrence(XLHFOccurrence.OddPages);
        }
        else
        {
            this.ClearOccurrence(occurrence);
        }
    }

    private void ClearOccurrence(XLHFOccurrence occurrence)
    {
        if (this.texts.ContainsKey(occurrence))
        {
            this.texts.Remove(occurrence);
        }
    }
}

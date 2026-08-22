#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.PageSetup;

internal class XLHeaderFooter : IXLHeaderFooter
{
    public XLHeaderFooter(XLWorksheet worksheet)
    {
        this.Worksheet = worksheet;
        this.Left = new XLHFItem(this);
        this.Right = new XLHFItem(this);
        this.Center = new XLHFItem(this);
        this.SetAsInitial();
    }

    public XLHeaderFooter(XLHeaderFooter defaultHF, XLWorksheet worksheet)
    {
        this.Worksheet = worksheet;
        defaultHF.innerTexts.ForEach(kp => this.innerTexts.Add(kp.Key, kp.Value));
        this.Left = new XLHFItem((XLHFItem)defaultHF.Left, this);
        this.Center = new XLHFItem((XLHFItem)defaultHF.Center, this);
        this.Right = new XLHFItem((XLHFItem)defaultHF.Right, this);
        this.SetAsInitial();
    }

    internal readonly XLWorksheet Worksheet;

    public IXLHFItem Left { get; private set; }
    public IXLHFItem Center { get; private set; }
    public IXLHFItem Right { get; private set; }

    public String GetText(XLHFOccurrence occurrence)
    {
        //if (innerTexts.ContainsKey(occurrence)) return innerTexts[occurrence];

        string retVal = String.Empty;
        string leftText = this.Left.GetText(occurrence);
        string centerText = this.Center.GetText(occurrence);
        string rightText = this.Right.GetText(occurrence);
        retVal += leftText.Length > 0 ? "&L" + leftText : String.Empty;
        retVal += centerText.Length > 0 ? "&C" + centerText : String.Empty;
        retVal += rightText.Length > 0 ? "&R" + rightText : String.Empty;
        if (retVal.Length > 255)
        {
            throw new ArgumentOutOfRangeException(
                "Headers and Footers cannot be longer than 255 characters (including style markups)"
            );
        }

        return retVal;
    }

    private Dictionary<XLHFOccurrence, String> innerTexts = new();

    internal void SetInnerText(XLHFOccurrence occurrence, String text)
    {
        List<ParsedHeaderFooterElement> parsedElements = ParseFormattedHeaderFooterText(text);

        if (parsedElements.Any(e => e.Position == 'L'))
        {
            this.Left.AddText(
                string.Join(
                    "\r\n",
                    parsedElements.Where(e => e.Position == 'L').Select(e => e.Text).ToArray()
                ),
                occurrence
            );
        }

        if (parsedElements.Any(e => e.Position == 'C'))
        {
            this.Center.AddText(
                string.Join(
                    "\r\n",
                    parsedElements.Where(e => e.Position == 'C').Select(e => e.Text).ToArray()
                ),
                occurrence
            );
        }

        if (parsedElements.Any(e => e.Position == 'R'))
        {
            this.Right.AddText(
                string.Join(
                    "\r\n",
                    parsedElements.Where(e => e.Position == 'R').Select(e => e.Text).ToArray()
                ),
                occurrence
            );
        }

        this.innerTexts[occurrence] = text;
    }

    private struct ParsedHeaderFooterElement
    {
        public char Position;
        public string Text;
    }

    private static List<ParsedHeaderFooterElement> ParseFormattedHeaderFooterText(string text)
    {
        Func<int, bool> IsAtPositionIndicator = i =>
            i < text.Length - 1
            && text[i] == '&'
            && (Enumerable.Contains(new char[] { 'L', 'C', 'R' }, text[i + 1]));

        List<ParsedHeaderFooterElement> parsedElements = [];
        char currentPosition = 'L'; // default is LEFT
        string hfElement = "";

        for (int i = 0; i < text.Length; i++)
        {
            if (IsAtPositionIndicator(i))
            {
                if (hfElement.Length > 0)
                {
                    parsedElements.Add(
                        new ParsedHeaderFooterElement()
                        {
                            Position = currentPosition,
                            Text = hfElement,
                        }
                    );
                }

                currentPosition = text[i + 1];
                i += 2;
                hfElement = "";
            }

            if (i < text.Length)
            {
                if (IsAtPositionIndicator(i))
                {
                    i--;
                }
                else
                {
                    hfElement += text[i];
                }
            }
        }

        if (hfElement.Length > 0)
        {
            parsedElements.Add(
                new ParsedHeaderFooterElement() { Position = currentPosition, Text = hfElement }
            );
        }

        return parsedElements;
    }

    private Dictionary<XLHFOccurrence, String> _initialTexts;

    private Boolean _changed;
    internal Boolean Changed
    {
        get => this._changed || this._initialTexts.Any(it => this.GetText(it.Key) != it.Value);
        set => this._changed = value;
    }

    internal void SetAsInitial()
    {
        this._initialTexts = new Dictionary<XLHFOccurrence, string>();
        foreach (XLHFOccurrence o in Enum.GetValues(typeof(XLHFOccurrence)).Cast<XLHFOccurrence>())
        {
            this._initialTexts.Add(o, this.GetText(o));
        }
    }

    public IXLHeaderFooter Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
    {
        this.Left.Clear(occurrence);
        this.Right.Clear(occurrence);
        this.Center.Clear(occurrence);
        return this;
    }
}

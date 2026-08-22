using System;

namespace XlsxSharp.Excel.RichText;

internal class XLPhonetic : IXLPhonetic
{
    public XLPhonetic(string text, int start, int end)
    {
        this.Text = text;
        this.Start = start;
        this.End = end;
    }

    public string Text { get; }
    public int Start { get; }
    public int End { get; }

    public bool Equals(IXLPhonetic? other)
    {
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return this.Text == other.Text && this.Start == other.Start && this.End == other.End;
    }
}

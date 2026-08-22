using System;

namespace XlsxSharp.Excel.RichText;

internal class XLPhonetic : IXLPhonetic
{
    public XLPhonetic(String text, Int32 start, Int32 end)
    {
        this.Text = text;
        this.Start = start;
        this.End = end;
    }

    public String Text { get; }
    public Int32 Start { get; }
    public Int32 End { get; }

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

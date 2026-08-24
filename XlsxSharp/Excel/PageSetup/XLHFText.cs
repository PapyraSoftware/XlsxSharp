#nullable disable

using System.Text;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.PageSetup;

internal class XLHFText
{
    private readonly XLHFItem _hfItem;

    public XLHFText(XLRichString richText, XLHFItem hfItem)
    {
        this.RichText = richText;
        this._hfItem = hfItem;
    }

    public XLRichString RichText { get; private set; }

    public string GetHFText(string prevText)
    {
        IXLFont wsFont = this._hfItem.HeaderFooter.Worksheet.Style.Font;

        bool isRichText =
            this.RichText.FontName != null && this.RichText.FontName != wsFont.FontName
            || this.RichText.Bold != wsFont.Bold
            || this.RichText.Italic != wsFont.Italic
            || this.RichText.Strikethrough != wsFont.Strikethrough
            || this.RichText.FontSize > 0
                && Math.Abs(this.RichText.FontSize - wsFont.FontSize) > XlsxSharp.XLHelper.Epsilon
            || this.RichText.VerticalAlignment != wsFont.VerticalAlignment
            || this.RichText.Underline != wsFont.Underline
            || !this.RichText.FontColor.Equals(wsFont.FontColor);

        if (!isRichText)
        {
            return this.RichText.Text;
        }

        StringBuilder sb = new();

        if (this.RichText.FontName != null && this.RichText.FontName != wsFont.FontName)
        {
            sb.Append("&\"" + this.RichText.FontName);
        }
        else
        {
            sb.Append("&\"-");
        }

        if (this.RichText.Bold && this.RichText.Italic)
        {
            sb.Append(",Bold Italic\"");
        }
        else if (this.RichText.Bold)
        {
            sb.Append(",Bold\"");
        }
        else if (this.RichText.Italic)
        {
            sb.Append(",Italic\"");
        }
        else
        {
            sb.Append(",Regular\"");
        }

        if (
            this.RichText.FontSize > 0
            && Math.Abs(this.RichText.FontSize - wsFont.FontSize) > XlsxSharp.XLHelper.Epsilon
        )
        {
            sb.Append("&" + this.RichText.FontSize);
        }

        if (this.RichText.Strikethrough && !wsFont.Strikethrough)
        {
            sb.Append("&S");
        }

        if (this.RichText.VerticalAlignment != wsFont.VerticalAlignment)
        {
            if (this.RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript)
            {
                sb.Append("&Y");
            }
            else if (
                this.RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript
            )
            {
                sb.Append("&X");
            }
        }

        if (this.RichText.Underline != wsFont.Underline)
        {
            if (this.RichText.Underline == XLFontUnderlineValues.Single)
            {
                sb.Append("&U");
            }
            else if (this.RichText.Underline == XLFontUnderlineValues.Double)
            {
                sb.Append("&E");
            }
        }

        int lastColorPosition = prevText.LastIndexOf("&K");

        if (
            (
                lastColorPosition >= 0
                && !this.RichText.FontColor.Equals(
                    XLColor.FromHtml("#" + prevText.Substring(lastColorPosition + 2, 6))
                )
            ) || (lastColorPosition == -1 && !this.RichText.FontColor.Equals(wsFont.FontColor))
        )
        {
            sb.Append("&K" + this.RichText.FontColor.Color.ToHex().Substring(2));
        }

        sb.Append(this.RichText.Text);

        if (this.RichText.Underline != wsFont.Underline)
        {
            if (this.RichText.Underline == XLFontUnderlineValues.Single)
            {
                sb.Append("&U");
            }
            else if (this.RichText.Underline == XLFontUnderlineValues.Double)
            {
                sb.Append("&E");
            }
        }

        if (this.RichText.VerticalAlignment != wsFont.VerticalAlignment)
        {
            if (this.RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript)
            {
                sb.Append("&Y");
            }
            else if (
                this.RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript
            )
            {
                sb.Append("&X");
            }
        }

        if (this.RichText.Strikethrough && !wsFont.Strikethrough)
        {
            sb.Append("&S");
        }

        return sb.ToString();
    }
}

using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel.RichText;

/// <summary>
/// An API object for manipulating rich text. Every time it is changed, it calls
/// <see cref="OnContentChanged"/> to project changes back to the <see cref="SharedStringTable"/>.
/// </summary>
internal class XLRichText : XLFormattedText<IXLRichText>, IXLRichText
{
    // Should be set as the last thing in ctor to prevent firing changes to immutable rich text during ctor
    private readonly XLCell? _cell;

    /// <summary>
    /// Copy ctor to return user modifiable rich text from an immutable rich text stored
    /// in the shared string table.
    /// </summary>
    internal XLRichText(XLCell cell, XLFontFormatValue defaultFont, XLImmutableRichText original)
        : base(defaultFont, cell.Worksheet.Workbook.Styles)
    {
        foreach (XLImmutableRichText.RichTextRun originalRun in original.Runs)
        {
            string runText = original.GetRunText(originalRun);
            this.AddText(
                new XLRichString(
                    runText,
                    originalRun.Font,
                    this,
                    this.Styles,
                    this.OnContentChanged
                )
            );
        }

        bool hasPhonetics =
            original.PhoneticRuns.Count > 0 || original.PhoneticsProperties.HasValue;
        if (hasPhonetics)
        {
            XLPhonetics phonetics;
            if (original.PhoneticsProperties.HasValue)
            {
                XLImmutableRichText.PhoneticProperties originalProps = original
                    .PhoneticsProperties
                    .Value;
                phonetics = new XLPhonetics(
                    originalProps.Font,
                    defaultFont,
                    this.Styles,
                    this.OnContentChanged
                )
                {
                    Type = originalProps.Type,
                    Alignment = originalProps.Alignment,
                };
            }
            else
            {
                phonetics = new XLPhonetics(
                    defaultFont,
                    defaultFont,
                    this.Styles,
                    this.OnContentChanged
                );
            }

            foreach (XLImmutableRichText.PhoneticRun phoneticRun in original.PhoneticRuns)
            {
                phonetics.Add(phoneticRun.Text, phoneticRun.StartIndex, phoneticRun.EndIndex);
            }

            this.Phonetics = phonetics;
        }

        // TODO Styles: Convert to a factory method. The cell is set at the end to avoid false change trigger. Refactor so it's not needed anymore
        this.Container = this;
        this._cell = cell;
    }

    internal XLRichText(XLCell cell, XLFontFormatValue defaultFont, String text)
        : this(cell, defaultFont) =>
        this.AddText(new XLRichString(text, defaultFont, this, this.Styles, this.OnContentChanged));

    internal XLRichText(XLCell cell, XLFontFormatValue defaultFont)
        : base(defaultFont, cell.Worksheet.Workbook.Styles)
    {
        this.Container = this;
        this._cell = cell;
    }

    protected override void OnContentChanged()
    {
        // The rich text is still being created
        if (this._cell is null)
        {
            return;
        }

        if (this._cell.DataType != XLDataType.Text || !this._cell.HasRichText)
        {
            throw new InvalidOperationException("The rich text isn't a content of a cell.");
        }

        this._cell.SetOnlyValue(this.Text);
        Point point = this._cell.Point;
        XLImmutableRichText richText = XLImmutableRichText.Create(this);
        this._cell.Worksheet.Internals.CellsCollection.ValueSlice.SetRichText(point, richText);
    }
}

using System.Diagnostics;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel;

/// <summary>
/// A container for styles and formatting records in a workbook.
/// </summary>
internal class XLWorkbookStyles
{
    /// <summary>
    /// First user-defined numFmtId.
    /// </summary>
    public const int FirstUserDefinedNumberFormatIndex = 164;

    private readonly BiDictionary<int, XLNumberFormat> _numberFormats;

    private readonly BiDictionary<int, XLFontFormatValue> _fontFormats;

    private readonly BiDictionary<int, XLFillFormatValue> _fillFormats;

    private readonly BiDictionary<int, XLBorderFormatValue> _borderFormats;

    private readonly BiDictionary<int, XLAlignmentFormatValue> _alignmentFormats;

    private readonly BiDictionary<int, XLProtectionFormatValue> _protectionFormats;

    /// <summary>
    /// The key is XfId, the value is cell format.
    /// </summary>
    private readonly BiDictionary<int, XLCellFormatValue> _cellFormats;

    /// <summary>
    /// The key is cellStyleXfId, the value is cell style.
    /// </summary>
    private readonly BiDictionary<StyleId, XLCellStyleValue> _cellStyles;

    private readonly BiDictionary<int, XLDxfValue> _differentialFormats;

    /// <summary>
    /// Key is a table style name, value is a table style.
    /// </summary>
    private readonly Dictionary<string, XLTableTheme> _tableStyles;

    /// <summary>
    /// Key is a pivot table style name, value is a pivot table style.
    /// </summary>
    private readonly Dictionary<string, XLPivotTableStyle> _pivotStyles;

    private List<uint>? _indexedColorsArgb;

    private List<XLColor> _mruColors = [];

    /// <summary>
    /// A normal style that is used for newly create workbooks or loaded workbooks without a normal style.
    /// </summary>
    internal readonly XLCellStyleValue DefaultNormalStyle = new()
    {
        Name = "Normal",
        BuiltInStyle = BuiltInStyleValues.Normal,
        Hidden = false,
        Alignment = new XLAlignmentFormatValue
        {
            Horizontal = XLAlignmentHorizontalValues.General,
            Vertical = XLAlignmentVerticalValues.Bottom,
            TextRotation = TextRotation.None,
            WrapText = false,
            Indent = 0,
            RelativeIndent = 0,
            JustifyLastLine = false,
            ShrinkToFit = false,
            ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent,
        },
        Protection = new XLProtectionFormatValue { Locked = true, Hidden = false },
        NumberFormat = XLPredefinedFormat.FormatCodes[XLPredefinedFormat.General],
        Font = new XLFontFormatValue
        {
            Name = "Calibri",
            Charset = XLFontCharSet.Ansi,
            Family = XLFontFamilyNumberingValues.Swiss,
            Bold = false,
            Italic = false,
            Strikethrough = false,
            Outline = false,
            Shadow = false,
            Condense = false,
            Extend = false,
            Color = XLColor.Black,
            Size = XLFontSize.FromPoints(11),
            Underline = XLFontUnderlineValues.None,
            VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
            Scheme = XLFontScheme.None,
        },
        Fill = XLFillFormatValue.None,
        Border = XLBorderFormatValue.None,
        IncludedComponents = CellFormatComponents.All,
    };

    /// <summary>
    /// A cell format that is used when an element doesn't explicitly define formatting. This
    /// format must be saved at index 0 in a file. The likely reason is that when an element
    /// (e.g., a cell) in a XML doesn't explicitly define index of a format, the default value
    /// is 0 = this format.
    /// </summary>
    internal XLCellFormatValue DefaultCellFormat => this._cellFormats[0];

    internal XLWorkbookStyles()
    {
        this._numberFormats = new BiDictionary<int, XLNumberFormat>();
        this._fontFormats = new BiDictionary<int, XLFontFormatValue>();
        this._fillFormats = new BiDictionary<int, XLFillFormatValue>();
        this._borderFormats = new BiDictionary<int, XLBorderFormatValue>();
        this._alignmentFormats = new BiDictionary<int, XLAlignmentFormatValue>();
        this._protectionFormats = new BiDictionary<int, XLProtectionFormatValue>();
        this._cellFormats = new BiDictionary<int, XLCellFormatValue>();
        this._cellStyles = new BiDictionary<StyleId, XLCellStyleValue>();
        this._differentialFormats = new BiDictionary<int, XLDxfValue>();
        this._tableStyles = new Dictionary<string, XLTableTheme>(XlsxSharp.XLHelper.NameComparer);
        this._pivotStyles = new Dictionary<string, XLPivotTableStyle>(
            XlsxSharp.XLHelper.NameComparer
        );
    }

    internal IReadOnlyBiDictionary<int, XLNumberFormat> NumberFormats => this._numberFormats;

    internal IReadOnlyBiDictionary<int, XLFontFormatValue> Fonts => this._fontFormats;

    internal IReadOnlyBiDictionary<int, XLFillFormatValue> Fills => this._fillFormats;

    internal IReadOnlyBiDictionary<int, XLBorderFormatValue> Borders => this._borderFormats;

    internal IReadOnlyBiDictionary<int, XLCellFormatValue> CellFormats => this._cellFormats;

    internal IReadOnlyBiDictionary<StyleId, XLCellStyleValue> CellStyles => this._cellStyles;

    internal IReadOnlyBiDictionary<int, XLDxfValue> DifferentialFormats =>
        this._differentialFormats;

    internal IReadOnlyDictionary<string, XLTableTheme> TableStyles => this._tableStyles;

    internal IReadOnlyDictionary<string, XLPivotTableStyle> PivotStyles => this._pivotStyles;

    /// <summary>
    /// Name of a table style that should be used for newly added tables. It's not used for tables
    /// without a specified style.
    /// </summary>
    internal string? DefaultTableStyle { get; set; }

    /// <summary>
    /// Name of a pivot style that should be used for newly added pivot tables. It's not used for
    /// pivot tables without a specified style.
    /// </summary>
    internal string? DefaultPivotStyle { get; set; }

    /// <summary>
    /// Some workbooks use indexed colors that are not in the standard <see cref="XLColor.IndexedColors"/>,
    /// but have their own list of indexed colors. Legacy feature, do not expose. If the value is null, use
    /// predefined indexed colors.
    /// </summary>
    internal IReadOnlyList<uint>? IndexedColorsArgb => this._indexedColorsArgb;

    /// <summary>
    /// Most recently used colors are colors <em>hand-picked</em> by a user that are displayed in
    /// the color picker dialogue. The standard and theme colors are always offered in the color
    /// picker, so they are (in general) not added to the MRU color list.
    /// </summary>
    internal IReadOnlyList<XLColor> MruColors => this._mruColors;

    /// <summary>
    /// A default format values used when format doesn't have a value in a property. All props in
    /// the default format must have a value. The default is set on load and not changed later.
    /// Nearly all props are equivalent of "zero", except things that can't be like that, e.g. font
    /// name or font size.
    /// </summary>
    // TODO: Make private and use GetDefaultFormat
    internal XLCellFormatValue DefaultFormat { get; set; } =
        new()
        {
            Font = new XLFontFormatValue
            {
                Name = "Calibri",
                Charset = XLFontCharSet.Ansi,
                Family = XLFontFamilyNumberingValues.NotApplicable,
                Bold = false,
                Italic = false,
                Strikethrough = false,
                Outline = false,
                Shadow = false,
                Condense = false,
                Extend = false,
                Color = XLColor.FromArgb(0x00000000),
                Size = XLFontSize.FromPoints(11),
                Underline = XLFontUnderlineValues.None,
                VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
                Scheme = XLFontScheme.None,
            },
            NumberFormat = XLPredefinedFormat.FormatCodes[XLPredefinedFormat.General],
            Alignment = new XLAlignmentFormatValue()
            {
                Horizontal = XLAlignmentHorizontalValues.General,
                Vertical = XLAlignmentVerticalValues.Bottom,
                TextRotation = TextRotation.None,
                WrapText = false,
                Indent = 0,
                RelativeIndent = 0,
                JustifyLastLine = false,
                ShrinkToFit = false,
                ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent,
            },
            Protection = new XLProtectionFormatValue { Locked = true, Hidden = false },
            Fill = XLFillFormatValue.None,
            Border = XLBorderFormatValue.None,
            CellStyleId = null,
            IncludeQuotePrefix = false,
            PivotButton = false,
            CustomFormat = CellFormatComponents.None,
        };

    internal void AddNumberFormat(int numFmtId, XLNumberFormat format) =>
        this._numberFormats.Add(numFmtId, format);

    internal void AddUserDefinedNumberFormat(XLNumberFormat numberFormat)
    {
        int numFmtId = FirstUserDefinedNumberFormatIndex;
        if (this._numberFormats.Count > 0)
        {
            numFmtId = Math.Max(this._numberFormats.Keys.Max() + 1, numFmtId);
        }

        this._numberFormats.Add(numFmtId, numberFormat);
    }

    internal void AddFontFormat(XLFontFormatValue fontFormat) =>
        this._fontFormats.Add(this._fontFormats.Count, fontFormat);

    internal void AddFillFormat(XLFillFormatValue fillFormat) =>
        this._fillFormats.Add(this._fillFormats.Count, fillFormat);

    internal void AddBorderFormat(XLBorderFormatValue borderFormat) =>
        this._borderFormats.Add(this._borderFormats.Count, borderFormat);

    internal void AddFormat(XLCellFormatValue cellFormat)
    {
        int xfId = this._cellFormats.Count;
        this._cellFormats.Add(xfId, cellFormat);
    }

    internal void AddCellStyle(int cellStyleXfId, XLCellStyleValue cellStyle) =>
        this._cellStyles.Add(cellStyleXfId, cellStyle);

    internal void AddDifferentialFormat(XLDxfValue dxf) =>
        this._differentialFormats.Add(this._differentialFormats.Count, dxf);

    internal void AddTableStyle(XLTableTheme tableStyle) =>
        this._tableStyles.Add(tableStyle.Name, tableStyle);

    internal void AddPivotStyle(XLPivotTableStyle pivotStyle) =>
        this._pivotStyles.Add(pivotStyle.Name, pivotStyle);

    internal void SetIndexedColors(List<uint> indexedColors) =>
        this._indexedColorsArgb = indexedColors;

    internal void SetMruColors(List<XLColor> mruColors) => this._mruColors = mruColors;

    internal XLNumberFormat RegisterNumberFormat(XLNumberFormat numberFormat)
    {
        if (this._numberFormats.TryGetValue(numberFormat, out XLNumberFormat existingFormat))
        {
            return existingFormat;
        }

        this.AddUserDefinedNumberFormat(numberFormat);
        return numberFormat;
    }

    private XLAlignmentFormatValue GetRegisteredAlignmentFormat(
        XLAlignmentFormatValue original,
        Func<XLAlignmentFormatValue, XLAlignmentFormatValue> modify
    ) => this.RegisterAlignmentFormat(modify(original));

    internal XLAlignmentFormatValue RegisterAlignmentFormat(XLAlignmentFormatValue alignment)
    {
        if (
            this._alignmentFormats.TryGetValue(
                alignment,
                out XLAlignmentFormatValue? existingAlignment
            )
        )
        {
            return existingAlignment;
        }

        this._alignmentFormats.Add(this._alignmentFormats.Count, alignment);
        return alignment;
    }

    internal XLProtectionFormatValue RegisterProtectionFormat(XLProtectionFormatValue protection)
    {
        if (
            this._protectionFormats.TryGetValue(
                protection,
                out XLProtectionFormatValue? existingProtection
            )
        )
        {
            return existingProtection;
        }

        this._protectionFormats.Add(this._protectionFormats.Count, protection);
        return protection;
    }

    /// <summary>
    /// Get a font format that is stored in the internal structures of the styles class. The font
    /// format is created by modification of existing font format. This is essential for saving,
    /// all formats must be registered in the styles class.
    /// </summary>
    internal XLFontFormatValue GetRegisteredFontFormat(
        XLFontFormatValue original,
        Func<XLFontFormatValue, XLFontFormatValue> modify
    )
    {
        XLFontFormatValue modified = modify(original);
        return this.RegisterFontFormat(modified);
    }

    internal XLFontFormatValue RegisterFontFormat(XLFontFormatValue font)
    {
        if (this._fontFormats.TryGetValue(font, out XLFontFormatValue? existingFont))
        {
            return existingFont;
        }

        this.AddFontFormat(font);
        return font;
    }

    internal XLCellFormatValue GetModifiedFormat(
        XLCellFormatValue originalFormat,
        XLNumberFormat numberFormat
    )
    {
        XLNumberFormat modifiedNumberFormat = this.RegisterNumberFormat(numberFormat);
        XLCellFormatValue modifiedFormat = this.GetRegisteredCellFormat(
            originalFormat,
            format => format with { NumberFormat = modifiedNumberFormat }
        );
        return modifiedFormat;
    }

    internal XLCellFormatValue GetModifiedFormat(
        XLCellFormatValue originalFormat,
        Func<XLAlignmentFormatValue, XLAlignmentFormatValue> modify
    )
    {
        XLAlignmentFormatValue modifiedAlignment = this.GetRegisteredAlignmentFormat(
            originalFormat.Alignment,
            modify
        );
        XLCellFormatValue modifiedFormat = this.GetRegisteredCellFormat(
            originalFormat,
            format => format with { Alignment = modifiedAlignment }
        );
        return modifiedFormat;
    }

    internal XLCellFormatValue GetModifiedFormat(
        XLCellFormatValue originalFormat,
        Func<XLFontFormatValue, XLFontFormatValue> modify
    )
    {
        XLFontFormatValue modifiedFont = this.GetRegisteredFontFormat(originalFormat.Font, modify);
        XLCellFormatValue modifiedFormat = this.GetRegisteredCellFormat(
            originalFormat,
            format => format with { Font = modifiedFont }
        );
        return modifiedFormat;
    }

    internal XLCellFormatValue GetModifiedFormat(
        XLCellFormatValue originalFormat,
        Func<XLBorderFormatValue, XLBorderFormatValue> modify
    )
    {
        XLBorderFormatValue modifiedBorder = this.GetRegisteredBorderFormat(
            originalFormat.Border,
            modify
        );
        XLCellFormatValue modifiedFormat = this.GetRegisteredCellFormat(
            originalFormat,
            format => format with { Border = modifiedBorder }
        );
        return modifiedFormat;
    }

    internal XLFillFormatValue GetRegisteredFillFormat(
        XLFillFormatValue original,
        Func<XLFillFormatValue, XLFillFormatValue> modify
    )
    {
        XLFillFormatValue modified = modify(original);
        return this.RegisterFillFormat(modified);
    }

    private XLFillFormatValue RegisterFillFormat(XLFillFormatValue fill)
    {
        if (this._fillFormats.TryGetValue(fill, out XLFillFormatValue? existingFill))
        {
            return existingFill;
        }

        this.AddFillFormat(fill);
        return fill;
    }

    internal XLBorderFormatValue GetRegisteredBorderFormat(
        XLBorderFormatValue original,
        Func<XLBorderFormatValue, XLBorderFormatValue> modify
    )
    {
        XLBorderFormatValue modified = modify(original);
        return this.RegisterBorderFormat(modified);
    }

    private XLBorderFormatValue RegisterBorderFormat(XLBorderFormatValue border)
    {
        if (this._borderFormats.TryGetValue(border, out XLBorderFormatValue? existingBorder))
        {
            return existingBorder;
        }

        this.AddBorderFormat(border);
        return border;
    }

    internal XLCellFormatValue GetRegisteredCellFormat(
        XLCellFormatValue original,
        Func<XLCellFormatValue, XLCellFormatValue> modify
    )
    {
        XLCellFormatValue modified = modify(original);
        return this.RegisterCellFormat(modified);
    }

    internal XLCellFormatValue RegisterCellFormat(XLCellFormatValue cellFormat)
    {
        if (this._cellFormats.TryGetValue(cellFormat, out XLCellFormatValue? existing))
        {
            return existing;
        }

        Debug.Assert(this._numberFormats.ContainsValue(cellFormat.NumberFormat));
        Debug.Assert(
            this._alignmentFormats.TryGetValue(
                cellFormat.Alignment,
                out XLAlignmentFormatValue? registeredAlignment
            ) && ReferenceEquals(cellFormat.Alignment, registeredAlignment)
        );
        Debug.Assert(
            this._protectionFormats.TryGetValue(
                cellFormat.Protection,
                out XLProtectionFormatValue? registeredProtection
            ) && ReferenceEquals(cellFormat.Protection, registeredProtection)
        );
        Debug.Assert(
            this._fontFormats.TryGetValue(cellFormat.Font, out XLFontFormatValue? registeredFont)
                && ReferenceEquals(cellFormat.Font, registeredFont)
        );
        Debug.Assert(
            this._fillFormats.TryGetValue(cellFormat.Fill, out XLFillFormatValue? registeredFill)
                && ReferenceEquals(cellFormat.Fill, registeredFill)
        );
        Debug.Assert(
            this._borderFormats.TryGetValue(
                cellFormat.Border,
                out XLBorderFormatValue? registeredBorder
            ) && ReferenceEquals(cellFormat.Border, registeredBorder)
        );
        this.AddFormat(cellFormat);
        return cellFormat;
    }

    /// <summary>
    /// Get registered format equal to <paramref name="format"/> from the styles. Generally for copying formats other workbooks.
    /// </summary>
    internal XLCellFormatValue GetRegisteredCellFormat(XLCellFormatValue format)
    {
        // If format is already registered, all its components must be registered too.
        if (this._cellFormats.TryGetValue(format, out XLCellFormatValue? existing))
        {
            return existing;
        }

        // TODO: If format is from different workbook, we should copy style if from different workbook and this workbook doesn't already have a style with same name
        StyleId? formatStyleId =
            format.CellStyleId is { } cellStyleId && this._cellStyles.ContainsKey(cellStyleId)
                ? cellStyleId
                : null;

        // We have to create new one, because some components may already exist here
        XLCellFormatValue cellFormat = new()
        {
            NumberFormat = this.RegisterNumberFormat(format.NumberFormat),
            Alignment = this.RegisterAlignmentFormat(format.Alignment),
            Protection = this.RegisterProtectionFormat(format.Protection),
            Font = this.RegisterFontFormat(format.Font),
            Fill = this.RegisterFillFormat(format.Fill),
            Border = this.RegisterBorderFormat(format.Border),
            CellStyleId = formatStyleId,
            IncludeQuotePrefix = format.IncludeQuotePrefix,
            PivotButton = format.PivotButton,
            CustomFormat = format.CustomFormat,
        };

        this.AddFormat(cellFormat);
        return cellFormat;
    }

    /// <summary>
    /// Get a differential format that is stored in the internal structures of the styles class.
    /// The differential format is created by modification of existing dxf format. This is
    /// essential for saving, all formats must be registered in the styles class.
    /// </summary>
    internal XLDxfValue GetRegisteredDxFormat(
        XLDxfValue original,
        Func<XLDxfValue, XLDxfValue> modify
    )
    {
        XLDxfValue modified = modify(original);
        return this.RegisterDxFormat(modified);
    }

    /// <summary>
    /// Register dxf from potentially different workbook into this workbook.
    /// </summary>
    /// <returns>Registered instance.</returns>
    internal XLDxfValue RegisterDxFormat(XLDxfValue dxf)
    {
        if (this._differentialFormats.TryGetValue(dxf, out XLDxfValue? existingDxf))
        {
            return existingDxf;
        }

        this.AddDifferentialFormat(dxf);
        return dxf;
    }

    /// <summary>
    /// Create a workbook styles component suitable for a new workbook.
    /// </summary>
    internal static XLWorkbookStyles CreateInitialized()
    {
        XLWorkbookStyles styles = new()
        {
            DefaultTableStyle = XLTableTheme.TableStyleMedium2.ToString(),
            DefaultPivotStyle = nameof(XLPivotTableTheme.PivotStyleLight16),
        };

        foreach ((int numFmtId, XLNumberFormat formatCode) in XLPredefinedFormat.FormatCodes)
        {
            styles.AddNumberFormat(numFmtId, formatCode);
        }

        XLCellStyleValue normalStyle = styles.DefaultNormalStyle;
        styles.AddFontFormat(normalStyle.Font);
        styles.AddFillFormat(XLFillFormatValue.None);
        styles.AddFillFormat(XLFillFormatValue.Gray125);
        styles.AddBorderFormat(XLBorderFormatValue.None);
        styles.AddCellStyle(0, normalStyle);

        XLCellFormatValue defaultFormat = XLCellFormatValue.FromStyle(0, normalStyle);
        styles.DefaultFormat = styles.GetRegisteredCellFormat(defaultFormat);

        return styles;
    }
}

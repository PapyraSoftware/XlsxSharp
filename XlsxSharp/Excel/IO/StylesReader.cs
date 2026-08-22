using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Tables;
using XlsxSharp.IO;
using XlsxSharp.Utils;
using PTS = XlsxSharp.Excel.Formatting.XLPivotStyleRegionValues;
using TS = XlsxSharp.Excel.Formatting.XLTableStyleRegionValues;

namespace XlsxSharp.Excel.IO;

internal partial class StylesReader
{
    private readonly XmlTreeReader _reader;
    private readonly XLWorkbookStyles _styles;
    private readonly string _ns = OpenXmlConst.Main2006SsNs;
    private readonly SequentialNameGenerator _styleNameGenerator = new("Style ", 1);

    // Format components to use when not specified in xf record
    private readonly XLNumberFormat _defaultNumberFormat;
    private readonly XLFillFormatValue _defaultFillFormat;
    private readonly XLBorderFormatValue _defaultBorderFormat;
    private readonly XLAlignmentFormatValue _defaultAlignmentFormat;
    private readonly XLProtectionFormatValue _defaultProtectionFormat;
    private XLFontFormatValue _defaultFontFormat;

    // Currently read CT_TableStyle element
    private Dictionary<TS, (XLDxfValue Dxf, int BandSize)> _currentTableStyle = new();
    private Dictionary<PTS, (XLDxfValue Dxf, int BandSize)> _currentPivotStyle = new();

    /// <summary>
    /// Style formats from <c>cellStyleXfs</c>.
    /// </summary>
    private List<XLCellFormatValue> _styleFormats = [];

    public StylesReader(XmlTreeReader reader, XLWorkbookStyles styles)
    {
        this._reader = reader;
        this._styles = styles;

        // Set initial fallback values if part is empty
        this._defaultNumberFormat = styles.DefaultNormalStyle.NumberFormat;
        this._defaultFontFormat = styles.DefaultNormalStyle.Font;
        this._defaultFillFormat = styles.DefaultNormalStyle.Fill;
        this._defaultBorderFormat = styles.DefaultNormalStyle.Border;
        this._defaultAlignmentFormat = styles.DefaultNormalStyle.Alignment;
        this._defaultProtectionFormat = styles.DefaultNormalStyle.Protection;
    }

    internal void Load()
    {
        this._reader.Open("styleSheet", this._ns);
        this.ParseStylesheet("styleSheet");

        this.LoadDefaultFormat();
    }

    private void LoadDefaultFormat()
    {
        // Normal style is technically optional, but it's basically a requirement for any sane work.
        (StyleId styleId, XLCellStyleValue? normalStyle) = this._styles.CellStyles.SingleOrDefault(
            x => x.Value.BuiltInStyle == BuiltInStyleValues.Normal
        );
        if (normalStyle is null)
        {
            styleId = this._styles.CellStyles.Count;
            normalStyle = this._styles.DefaultNormalStyle;

            // Number format collection already contains all predefined numFmts, so it can only be user-defined
            if (!this._styles.NumberFormats.ContainsValue(normalStyle.NumberFormat))
            {
                this._styles.AddUserDefinedNumberFormat(normalStyle.NumberFormat);
            }

            if (!this._styles.Fonts.ContainsValue(normalStyle.Font))
            {
                this._styles.AddFontFormat(normalStyle.Font);
            }

            if (!this._styles.Fills.ContainsValue(normalStyle.Fill))
            {
                this._styles.AddFillFormat(normalStyle.Fill);
            }

            if (!this._styles.Borders.ContainsValue(normalStyle.Border))
            {
                this._styles.AddBorderFormat(normalStyle.Border);
            }

            this._styles.AddCellStyle(styleId.Value, normalStyle);
        }

        // Ensure there is a default format.
        if (!this._styles.CellFormats.TryGetValue(0, out XLCellFormatValue? defaultFormat))
        {
            defaultFormat = XLCellFormatValue.FromStyle(styleId, normalStyle);
            this._styles.AddFormat(defaultFormat);
        }

        this._styles.DefaultFormat = defaultFormat;
    }

    private void ParseStylesheet(string elementName)
    {
        ParseNumFmts("numFmts", this._ns);

        // The spec says that the predefined formats have "formatCode value [..] implied rather
        // than explicitly saved in the file."... so if there was something saved, it should have
        // be used. If something was saved, it's explicit and explicit (generally) has preference
        // over implicit. It needs to be added after numFmts, but before cellStyleXfs/cellXfs.
        this.AddImpliedNumberFormats();

        ParseFonts("fonts", this._ns);

        if (this._styles.Fonts.Count == 0)
        {
            this._styles.AddFontFormat(this._defaultFontFormat);
        }
        else
        {
            this._defaultFontFormat = this._styles.Fonts[0];
        }

        ParseFills("fills", this._ns);

        // Default fill is always none, should be at index 0.
        if (!this._styles.Fills.ContainsValue(this._defaultFillFormat))
        {
            this._styles.AddFillFormat(this._defaultFillFormat);
        }

        ParseBorders("borders", this._ns);

        if (!this._styles.Borders.ContainsValue(this._defaultBorderFormat))
        {
            this._styles.AddBorderFormat(this._defaultBorderFormat);
        }

        ParseCellStyleXfs("cellStyleXfs", this._ns);

        List<(XLCellFormatValue Format, int? CellStyleXfId)> cellFormats = [];
        if (ParseCellXfs("cellXfs", this._ns) is { IsSuccess: true } cellXfsResult)
        {
            cellFormats = cellXfsResult.Value;
        }

        Dictionary<int, XLCellStyleValue> cellStyles = new();
        if (ParseCellStyles("cellStyles", this._ns) is { IsSuccess: true } cellStylesResult)
        {
            cellStyles = cellStylesResult.Value;
        }

        this.RepairMissingStyles(cellStyles);
        this.AddCellStyles(cellStyles);
        this.AddFormats(cellFormats, cellStyles);

        ParseDxfs("dxfs", this._ns);
        ParseTableStyles("tableStyles", this._ns);
        ParseColors("colors", this._ns);
        this.ParseExtensionList("extLst", this._ns);
        this._reader.Close(elementName, this._ns);
    }

    private void AddImpliedNumberFormats()
    {
        // Add predefined formats, so we can treat predefined number formats and
        // user defined number formats the same way.
        foreach ((int numFmtId, XLNumberFormat formatCode) in XLPredefinedFormat.FormatCodes)
        {
            if (!this._styles.NumberFormats.ContainsKey(numFmtId))
            {
                this._styles.AddNumberFormat(numFmtId, formatCode);
            }
        }
    }

    private void RepairMissingStyles(Dictionary<int, XLCellStyleValue> cellStyles)
    {
        // Because cellStyleXfs might be referenced from cell formats, each one must be converted
        // to a cell style. If the cellStyles didn't contain a record for any cellStyleXf, add it.
        for (int cellStyleXfId = 0; cellStyleXfId < this._styleFormats.Count; ++cellStyleXfId)
        {
            if (!cellStyles.ContainsKey(cellStyleXfId))
            {
                XLCellFormatValue format = this._styleFormats[cellStyleXfId];
                string generatedName = this._styleNameGenerator.NextUnusedStyleName();
                cellStyles.Add(
                    cellStyleXfId,
                    new XLCellStyleValue
                    {
                        Name = generatedName,
                        BuiltInStyle = null,
                        Hidden = false,
                        NumberFormat = format.NumberFormat,
                        Alignment = format.Alignment,
                        Protection = format.Protection,
                        Font = format.Font,
                        Fill = format.Fill,
                        Border = format.Border,
                        IncludedComponents = CellFormatComponents.All,
                    }
                );
            }
        }
    }

    private void AddCellStyles(Dictionary<int, XLCellStyleValue> cellStyles)
    {
        foreach ((int cellStyleXfId, XLCellStyleValue cellStyle) in cellStyles)
        {
            this._styles.AddCellStyle(cellStyleXfId, cellStyle);
        }
    }

    private void AddFormats(
        List<(XLCellFormatValue Format, int? CellStyleXfId)> cellFormats,
        Dictionary<int, XLCellStyleValue> cellStyles
    )
    {
        // At the time when cellXf were parsed, cell styles weren't resolved. Resolve them now.
        for (int xfId = 0; xfId < cellFormats.Count; ++xfId)
        {
            if (cellFormats[xfId].CellStyleXfId is { } cellStyleXfId)
            {
                // Make sure a the styleId is valid. The styleId for cell formats is read from file
                // and thus could be invalid. We don't want to crash later.
                if (!cellStyles.ContainsKey(cellStyleXfId))
                {
                    throw PartStructureException.InvalidAttributeValue();
                }

                XLCellFormatValue cellFormat = cellFormats[xfId].Format with
                {
                    CellStyleId = cellStyleXfId,
                };
                cellFormats[xfId] = (cellFormat, cellStyleXfId);
            }
        }

        foreach ((XLCellFormatValue cellFormat, int? _) in cellFormats)
        {
            this._styles.AddFormat(cellFormat);
        }
    }

    private static (int NumFmtId, XLNumberFormat FormatCode) OnNumFmtParsed(
        uint numFmtId,
        string formatCode
    ) => (checked((int)numFmtId), XLNumberFormat.Parse(formatCode));

    partial void OnNumFmtsParsed(List<(int NumFmtId, XLNumberFormat Format)> numFmt, uint? count)
    {
        foreach ((int numFmtId, XLNumberFormat formatCode) in numFmt)
        {
            // Even if numFmtId is predefined, store the supplied value. Excel accepts predefined
            // numFmtId and uses supplied format instead of predefined format. It fixes situation
            // during save, where such items are saved to user supplied numFmtId range.
            this._styles.AddNumberFormat(numFmtId, formatCode);
        }
    }

    private Xpr<XLDifferentialFontValue> ParseFont(string elementName, string ns)
    {
        if (!this._reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLDifferentialFontValue>();
        }

        // Font is mostly buggy specification. Excel basically chokes on anything but a sequence,
        // but standard requires an unbound choice where elements can repeat.
        XLFontName? fontName = null;
        XLFontCharSet? fontCharset = null;
        XLFontFamilyNumberingValues? fontFamily = null;
        bool? fontBold = null,
            fontItalic = null,
            fontStrikethrough = null,
            fontOutline = null,
            fontShadow = null,
            fontCondense = null,
            fontExtend = null;
        XLColor? fontColor = null;
        XLFontSize? fontSize = null;
        XLFontUnderlineValues? fontUnderline = null;
        XLFontVerticalTextAlignmentValues? fontVerticalAlignment = null;
        XLFontScheme? fontScheme = null;
        while (!this._reader.TryClose(elementName, ns))
        {
            if (this._reader.TryReadXStringValElement("name", this._ns, out string? name))
            {
                fontName = name;
            }
            else if (this._reader.TryReadIntValElement("charset", this._ns, out int charset))
            {
                fontCharset = (XLFontCharSet)charset;
            }
            else if (this._reader.TryReadIntValElement("family", this._ns, out int family))
            {
                // Bug in the spec. Spec says that it has values 0-14 and doesn't specify meaning
                // for the numerical values. It's supposed to refer to the same enum ST_FontFamily
                // as in WordML. The OI-29500 fixes this problem:
                // "Excel restricts the value of this attribute to be at least 0 and at most 5."
                fontFamily = family switch
                {
                    >= 0 and <= 5 => (XLFontFamilyNumberingValues)family,
                    > 5 and <= 14 => XLFontFamilyNumberingValues.NotApplicable,
                    _ => throw PartStructureException.InvalidAttributeFormat(),
                };
            }
            else if (this._reader.TryReadBoolElement("b", this._ns, out bool b))
            {
                fontBold = b;
            }
            else if (this._reader.TryReadBoolElement("i", this._ns, out bool i))
            {
                fontItalic = i;
            }
            else if (this._reader.TryReadBoolElement("strike", this._ns, out bool strike))
            {
                fontStrikethrough = strike;
            }
            else if (this._reader.TryReadBoolElement("outline", this._ns, out bool outline))
            {
                fontOutline = outline;
            }
            else if (this._reader.TryReadBoolElement("shadow", this._ns, out bool shadow))
            {
                fontShadow = shadow;
            }
            else if (this._reader.TryReadBoolElement("condense", this._ns, out bool condense))
            {
                fontCondense = condense;
            }
            else if (this._reader.TryReadBoolElement("extend", this._ns, out bool extend))
            {
                fontExtend = extend;
            }
            else if (this._reader.TryReadColor("color", this._ns, out XLColor? color))
            {
                fontColor = color;
            }
            else if (this._reader.TryOpen("sz", this._ns))
            {
                double fontSizePt = this._reader.GetDouble("val");
                this._reader.Close("sz", this._ns);
                fontSize = XLFontSize.FromPoints(fontSizePt);
            }
            else if (this._reader.TryOpen("u", this._ns))
            {
                XLFontUnderlineValues underline =
                    this._reader.GetOptionalEnum<XLFontUnderlineValues>("val")
                    ?? XLFontUnderlineValues.Single;
                this._reader.Close("u", this._ns);
                fontUnderline = underline;
            }
            else if (
                this._reader.TryReadEnumValElement<XLFontVerticalTextAlignmentValues>(
                    "vertAlign",
                    this._ns,
                    out XLFontVerticalTextAlignmentValues? vertAlign
                )
            )
            {
                fontVerticalAlignment = vertAlign;
            }
            else if (
                this._reader.TryReadEnumValElement<XLFontScheme>(
                    "scheme",
                    this._ns,
                    out XLFontScheme? scheme
                )
            )
            {
                fontScheme = scheme;
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(this._reader);
            }
        }

        XLDifferentialFontValue fontFormat = new()
        {
            Name = fontName,
            Charset = fontCharset,
            Family = fontFamily,
            Bold = fontBold,
            Italic = fontItalic,
            Strikethrough = fontStrikethrough,
            Outline = fontOutline,
            Shadow = fontShadow,
            Condense = fontCondense,
            Extend = fontExtend,
            Color = fontColor,
            Size = fontSize,
            Underline = fontUnderline,
            VerticalAlignment = fontVerticalAlignment,
            Scheme = fontScheme,
        };
        return Xpr.From(fontFormat);
    }

    // ParseFont is shared between <fonts> table and <dxf> elements. Once the <fonts> table is read,
    // register collected fonts to the workbook styles.
    partial void OnFontsParsed(List<XLDifferentialFontValue> font, uint? count)
    {
        XLFontFormatValue defaultFont = XLFontFormatValue.Default;

        // Excel probably screwed up. Default name and size should likely be font of default
        // format, but it is always taken from font zero (if font zero defines name/size).
        if (font.Count > 0)
        {
            XLDifferentialFontValue fontZero = font[0];
            defaultFont = defaultFont with
            {
                Name = fontZero.Name ?? defaultFont.Name,
                Size = fontZero.Size ?? defaultFont.Size,
            };
        }

        foreach (XLDifferentialFontValue fontProps in font)
        {
            XLFontFormatValue fontFormat = new()
            {
                Name = fontProps.Name ?? defaultFont.Name,
                Size = fontProps.Size ?? defaultFont.Size,
                Charset = fontProps.Charset ?? defaultFont.Charset,
                Family = fontProps.Family ?? defaultFont.Family,
                Bold = fontProps.Bold ?? defaultFont.Bold,
                Italic = fontProps.Italic ?? defaultFont.Italic,
                Strikethrough = fontProps.Strikethrough ?? defaultFont.Strikethrough,
                Outline = fontProps.Outline ?? defaultFont.Outline,
                Shadow = fontProps.Shadow ?? defaultFont.Shadow,
                Condense = fontProps.Condense ?? defaultFont.Condense,
                Extend = fontProps.Extend ?? defaultFont.Extend,
                Color = fontProps.Color ?? defaultFont.Color,
                Underline = fontProps.Underline ?? defaultFont.Underline,
                VerticalAlignment = fontProps.VerticalAlignment ?? defaultFont.VerticalAlignment,
                Scheme = fontProps.Scheme ?? defaultFont.Scheme,
            };
            this._styles.AddFontFormat(fontFormat);
        }
    }

    private XLFillFormatValue OnFillParsed(XLFillFormatValue? foundFill)
    {
        XLFillFormatValue fillFormat = foundFill ?? XLFillFormatValue.Empty;
        this._styles.AddFillFormat(fillFormat);
        return fillFormat;
    }

    private XLFillFormatValue OnPatternFillParsed(
        XLColor? fgColor,
        XLColor? bgColor,
        XLFillPatternValues? patternType
    )
    {
        // There is a discrepancy between <fill> interpretation for a solid fill:
        // * cell fill: Pattern color is ignored, the background color is used for fill
        // * dxf fill: Pattern color is the one used for fill, the background color is ignored
        // The GUI in both cases says that the background color is the one that is used. Therefore
        // use background is correct per GUI. The problem is that XlsxSharp historically says
        // the pattern color is the one that is used. This sucks, I have to live with it.
        // The alternative is a breaking change with no benefit.
        //
        // The other difference between dxf and cell fill is the default pattern type. Spec and
        // OI-29500 is silent, but Excel uses solid fill for dxf and none for cell fill.
        if (this._reader.Context[^2] == "dxf")
        {
            XLPatternFill patternFill = new()
            {
                PatternColor = fgColor ?? XLColor.Automatic,
                BackgroundColor = bgColor ?? XLColor.Automatic,
                PatternType = patternType ?? XLFillPatternValues.Solid,
            };
            return new XLFillFormatValue(patternFill);
        }
        else
        {
            XLFillPatternValues pattern = patternType ?? XLFillPatternValues.None;

            // Fix solid pattern discrepancy for cell fill
            if (pattern == XLFillPatternValues.Solid)
            {
                (bgColor, fgColor) = (fgColor, bgColor);
            }

            XLPatternFill patternFill = new()
            {
                PatternColor = fgColor ?? XLColor.Automatic,
                BackgroundColor = bgColor ?? XLColor.Automatic,
                PatternType = pattern,
            };
            return new XLFillFormatValue(patternFill);
        }
    }

    private static XLFillFormatValue OnGradientFillParsed(
        List<(FractionOfOne Value, XLColor Color)> stop,
        XLGradientType type,
        double degree,
        double left,
        double right,
        double top,
        double bottom
    )
    {
        Dictionary<FractionOfOne, XLColor> stops = stop.ToDictionary(x => x.Value, x => x.Color);
        switch (type)
        {
            case XLGradientType.Linear:
                return new XLFillFormatValue(
                    new XLLinearGradientFill { Stops = stops, Degrees = degree }
                );
            case XLGradientType.Path:
                return new XLFillFormatValue(
                    new XLPathGradientFill
                    {
                        Stops = stops,
                        InnerLeft = left,
                        InnerRight = right,
                        InnerTop = top,
                        InnerBottom = bottom,
                    }
                );
            default:
                throw new UnreachableException();
        }
    }

    private static (FractionOfOne Position, XLColor Color) OnGradientStopParsed(
        XLColor color,
        double position
    ) =>
        // Spec requires stop positions to be 0..1, but doesn't have a type for that. Excel repairs workbook when it receives values outside 0..1.
        (position, color);

    private XLDifferentialBorderValue OnBorderParsed(
        XLBorderLine? left,
        XLBorderLine? right,
        XLBorderLine? top,
        XLBorderLine? bottom,
        XLBorderLine? diagonal,
        XLBorderLine? vertical,
        XLBorderLine? horizontal,
        bool? diagonalUp,
        bool? diagonalDown,
        bool outline
    )
    {
        XLDifferentialBorderValue dxfBorder = new()
        {
            Left = left,
            Right = right,
            Top = top,
            Bottom = bottom,
            Diagonal = diagonal,
            Vertical = vertical,
            Horizontal = horizontal,
            DiagonalUp = diagonalUp ?? false,
            DiagonalDown = diagonalDown ?? false,
            Outline = outline,
        };
        if (this._reader.Context[^1] == "borders")
        {
            XLBorderFormatValue cellBorder = XLBorderFormatValue.FromDxf(dxfBorder);
            this._styles.AddBorderFormat(cellBorder);
        }

        return dxfBorder;
    }

    private static XLBorderLine OnBorderPrParsed(XLColor? color, XLBorderStyleValues style) =>
        new(color ?? XLColor.Automatic, style);

    partial void OnCellStyleXfsParsed(
        List<(XLCellFormatValue Format, int? CellStyleXfId)> xf,
        uint? count
    ) => this._styleFormats = [.. xf.Select(x => x.Format)];

    private (XLCellFormatValue Format, int? CellStyleXfId) OnXfParsed(
        XLDifferentialAlignmentValue? alignment,
        XLDifferentialProtectionValue? protection,
        uint? numFmtId,
        uint? fontId,
        uint? fillId,
        uint? borderId,
        uint? xfId,
        bool quotePrefix,
        bool pivotButton,
        bool? applyNumberFormat,
        bool? applyFont,
        bool? applyFill,
        bool? applyBorder,
        bool? applyAlignment,
        bool? applyProtection
    )
    {
        // When xf is parsed, all number formats, fonts, fills and borders should already be read.
        XLNumberFormat numberFormat = this._defaultNumberFormat;
        if (
            numFmtId is not null
            && this._styles.NumberFormats.TryGetValue(
                checked((int)numFmtId),
                out XLNumberFormat numFmt
            )
        )
        {
            numberFormat = numFmt;
        }

        XLFontFormatValue font = fontId is not null
            ? this._styles.Fonts[checked((int)fontId)]
            : this._defaultFontFormat;
        XLFillFormatValue fill = fillId is not null
            ? this._styles.Fills[checked((int)fillId)]
            : this._defaultFillFormat;
        XLBorderFormatValue border = borderId is not null
            ? this._styles.Borders[checked((int)borderId)]
            : this._defaultBorderFormat;

        // Excel doesn't actually use the apply* for xf, but at least it writes as if it did. It
        // actually checks whether the id is same for xf and a style and if it is, the aspect
        // should be from a style. Excel is doesn't use it, other producers might.

        // Cell format has default apply* false (interpreted as "does format has its own custom format for this aspect")
        // Style has default apply* true (interpreted as "does style define this aspect")
        // The apply* attributes have default value true for cellStyleXfs and false for cellXfs.
        bool isStyleXf = this._reader.Context[^1] == "cellStyleXfs";
        bool isCellFormat = !isStyleXf;
        bool defaultApply = !isCellFormat;
        CellFormatComponents components = CellFormatComponents.None;

        if (applyNumberFormat ?? defaultApply)
        {
            components |= CellFormatComponents.NumberFormat;
        }

        if (applyFont ?? defaultApply)
        {
            components |= CellFormatComponents.Font;
        }

        if (applyFill ?? defaultApply)
        {
            components |= CellFormatComponents.Fill;
        }

        if (applyBorder ?? defaultApply)
        {
            components |= CellFormatComponents.Border;
        }

        if (applyAlignment ?? defaultApply)
        {
            components |= CellFormatComponents.Alignment;
        }

        if (applyProtection ?? defaultApply)
        {
            components |= CellFormatComponents.Protection;
        }

        XLAlignmentFormatValue formatAlignment = alignment is not null
            ? new XLAlignmentFormatValue
            {
                Horizontal = alignment.Horizontal ?? XLAlignmentFormatValue.Default.Horizontal,
                Vertical = alignment.Vertical ?? XLAlignmentFormatValue.Default.Vertical,
                TextRotation =
                    alignment.TextRotation ?? XLAlignmentFormatValue.Default.TextRotation,
                WrapText = alignment.WrapText ?? XLAlignmentFormatValue.Default.WrapText,
                Indent = alignment.Indent ?? XLAlignmentFormatValue.Default.Indent,
                RelativeIndent =
                    alignment.RelativeIndent ?? XLAlignmentFormatValue.Default.RelativeIndent,
                JustifyLastLine =
                    alignment.JustifyLastLine ?? XLAlignmentFormatValue.Default.JustifyLastLine,
                ShrinkToFit = alignment.ShrinkToFit ?? XLAlignmentFormatValue.Default.ShrinkToFit,
                ReadingOrder =
                    alignment.ReadingOrder ?? XLAlignmentFormatValue.Default.ReadingOrder,
            }
            : XLAlignmentFormatValue.Default;
        XLProtectionFormatValue formatProtection = protection is not null
            ? new XLProtectionFormatValue
            {
                Locked = protection.Locked ?? XLProtectionFormatValue.Default.Locked,
                Hidden = protection.Hidden ?? XLProtectionFormatValue.Default.Hidden,
            }
            : XLProtectionFormatValue.Default;
        XLCellFormatValue format = new()
        {
            NumberFormat = numberFormat,
            // Alignment is not copied from default format
            Alignment = this._styles.RegisterAlignmentFormat(formatAlignment),
            Protection = this._styles.RegisterProtectionFormat(formatProtection),
            Font = font,
            Fill = fill,
            Border = border,
            CellStyleId = null, // The style is set once cell styles are resolved
            IncludeQuotePrefix = quotePrefix,
            PivotButton = pivotButton,
            CustomFormat = components,
        };
        return (format, checked((int?)xfId));
    }

    private static List<(XLCellFormatValue Format, int? CellStyleXfId)> OnCellXfsParsed(
        List<(XLCellFormatValue Format, int? CellStyleXfId)> xf,
        uint? count
    ) => xf;

    private (int XfId, XLCellStyleValue Style) OnCellStyleParsed(
        string? name,
        uint xfId,
        uint? builtinId,
        uint? iLevel,
        bool? hidden,
        bool? customBuiltin
    )
    {
        // The quotePrefix and pivotButton attributes of the cellStyleXf are not applied, plus
        // the xfId of the cellStyle is ignored. The OI-29500 also requires uniqueness of xfId
        // in the cellStyle elements, although Excel can load such workbook and has several
        // "linked" styles. It's likely the separation into two elements is based on the internal
        // structure inside the Excel.

        // Fill dummy name for the style
        if (string.IsNullOrWhiteSpace(name))
        {
            name = this._styleNameGenerator.NextUnusedStyleName();
        }
        else
        {
            this._styleNameGenerator.AddName(name);
        }

        XLCellFormatValue cellStyleFormat = this._styleFormats[checked((int)xfId)];

        // If the built in style is an outline style, expand it to avoid ugly representation.
        // The spec has only one only builtIn id for all RowLevel* styles (1) and one builtIn
        // id for all ColLevel* styles (2). Since the iLevel/outlineLevel is used only for
        // the RowLevel/ColLevel (OI-29500), expand the builtIn+iLevel for Row/Col level into
        // a separate builtIn styles (101-107 for RowLevel1-7, 201-207 for ColumnLevel1-7).
        if (builtinId is 1 or 2)
        {
            builtinId = builtinId.Value * 100 + 1 + iLevel ?? 0;
        }

        // BuiltIn must be among defined built-in styles ("implementers should restrict the content
        // of this attribute to enumerations present in the list")
        if (
            builtinId is not null
            && !Enum.IsDefined((BuiltInStyleValues)checked((int)builtinId.Value))
        )
        {
            throw PartStructureException.InvalidAttributeFormat();
        }

        // The apply* attributes have default `true` for cellStyleXfs and `false` for cellXfs.
        // We already took care of correct default value during the parsing of <xf>, so we don't
        // have to deal with it here.
        CellFormatComponents styleIncludesComponents = cellStyleFormat.CustomFormat;
        XLCellStyleValue cellStyle = new()
        {
            Name = name,
            BuiltInStyle = builtinId is not null ? (BuiltInStyleValues)builtinId.Value : null,
            Hidden = hidden ?? false,
            NumberFormat = cellStyleFormat.NumberFormat,
            Alignment = cellStyleFormat.Alignment,
            Protection = cellStyleFormat.Protection,
            Font = cellStyleFormat.Font,
            Fill = cellStyleFormat.Fill,
            Border = cellStyleFormat.Border,
            IncludedComponents = styleIncludesComponents,
        };

        return (checked((int)xfId), cellStyle);
    }

    private Dictionary<int, XLCellStyleValue> OnCellStylesParsed(
        List<(int CellStyleXfId, XLCellStyleValue Style)> cellStyle,
        uint? count
    )
    {
        Dictionary<int, XLCellStyleValue> cellStyles = new();
        foreach ((int cellStyleXfId, XLCellStyleValue style) in cellStyle)
        {
            // Multiple cell styles use same style formatting - split them, so each one uses
            // separate formatting. I considered removing duplicates, but it could mean that I
            // also might remove normal style, which is not desirable.
            if (!cellStyles.ContainsKey(cellStyleXfId))
            {
                cellStyles.Add(cellStyleXfId, style);
            }
            else
            {
                this._styleFormats.Add(this._styleFormats[cellStyleXfId]);
                int newCellStyleXfId = this._styleFormats.Count - 1;
                cellStyles.Add(newCellStyleXfId, style);
            }
        }

        return cellStyles;
    }

    private XLDifferentialAlignmentValue OnCellAlignmentParsed(
        XLAlignmentHorizontalValues? horizontal,
        XLAlignmentVerticalValues vertical,
        uint? textRotation,
        bool? wrapText,
        uint? indent,
        int? relativeIndent,
        bool? justifyLastLine,
        bool? shrinkToFit,
        uint? readingOrder
    )
    {
        if (readingOrder is not null && readingOrder is not (0 or 1 or 2))
        {
            throw PartStructureException.InvalidAttributeFormat();
        }

        int normalizedTextRotation = OpenXmlHelper.NormalizeRotation(textRotation ?? 0);
        return new XLDifferentialAlignmentValue
        {
            Horizontal = horizontal,
            Vertical = vertical,
            TextRotation = textRotation is not null
                ? new TextRotation(normalizedTextRotation)
                : null,
            WrapText = wrapText,
            Indent = indent is not null ? checked((int)indent.Value) : null,
            RelativeIndent = relativeIndent,
            JustifyLastLine = justifyLastLine,
            ShrinkToFit = shrinkToFit,
            ReadingOrder = readingOrder is not null
                ? (XLAlignmentReadingOrderValues)readingOrder.Value
                : null,
        };
    }

    partial void OnDxfParsed(
        XLDifferentialFontValue? font,
        (int NumFmtId, XLNumberFormat Format)? numFmt,
        XLFillFormatValue? fill,
        XLDifferentialAlignmentValue? alignment,
        XLDifferentialBorderValue? border,
        XLDifferentialProtectionValue? protection
    )
    {
        XLDxfValue dxf = new()
        {
            NumberFormat = numFmt?.Format,
            Font = font ?? XLDifferentialFontValue.Empty,
            Fill = fill is not null
                ? new XLDifferentialFillValue(fill)
                : XLDifferentialFillValue.Empty,
            Alignment = alignment ?? XLDifferentialAlignmentValue.Empty,
            Border = border ?? XLDifferentialBorderValue.Empty,
            Protection = protection ?? XLDifferentialProtectionValue.Empty,
        };
        this._styles.AddDifferentialFormat(dxf);
    }

    partial void OnTableStyleElementParsed((TS?, PTS?) type, uint size, uint? dxfId)
    {
        // Skip definition without a differential format
        if (dxfId is null)
        {
            return;
        }

        // Excel permits only 0-9
        if (size > 9)
        {
            throw PartStructureException.InvalidAttributeFormat(
                nameof(size),
                size.ToString(),
                this._reader
            );
        }

        // If there is a duplicate definition for a type, last one wins
        XLDxfValue dxf = this._styles.DifferentialFormats[checked((int)dxfId.Value)];

        if (type.Item1 is { } tableStyleRegion)
        {
            this._currentTableStyle[tableStyleRegion] = (dxf, (int)size);
        }

        if (type.Item2 is { } pivotStyleRegion)
        {
            this._currentPivotStyle[pivotStyleRegion] = (dxf, (int)size);
        }
    }

    partial void OnTableStyleParsed(string name, bool pivot, bool table, uint? count)
    {
        // Because of tableStyle element duality, we are filling both styles and
        // only insert types that have set flag.
        if (table)
        {
            XLTableTheme tableStyle = new(name);
            foreach ((TS region, (XLDxfValue dxf, int bandSize)) in this._currentTableStyle)
            {
                tableStyle.SetRegionFormat(region, dxf, bandSize);
            }

            this._styles.AddTableStyle(tableStyle);
        }

        this._currentTableStyle = new Dictionary<TS, (XLDxfValue Dxf, int BandSize)>();

        if (pivot)
        {
            XLPivotTableStyle pivotStyle = new(name);
            foreach ((PTS region, (XLDxfValue dxf, int bandSize)) in this._currentPivotStyle)
            {
                pivotStyle.SetRegionFormat(region, dxf, bandSize);
            }

            this._styles.AddPivotStyle(pivotStyle);
        }

        this._currentPivotStyle = new Dictionary<PTS, (XLDxfValue Dxf, int BandSize)>();
    }

    partial void OnTableStylesParsed(
        uint? count,
        string? defaultTableStyle,
        string? defaultPivotStyle
    )
    {
        if (!string.IsNullOrEmpty(defaultTableStyle))
        {
            this._styles.DefaultTableStyle = defaultTableStyle;
        }

        if (!string.IsNullOrEmpty(defaultPivotStyle))
        {
            this._styles.DefaultPivotStyle = defaultPivotStyle;
        }
    }

    private static uint OnRgbColorParsed(uint? rgb) =>
        // Despite the name, it's ARGB. If not specified, use black (Excel supplies 0x00000000, but
        // Excel plays very fast and loose with transparency).
        rgb ?? 0xFF000000;

    partial void OnIndexedColorsParsed(List<uint> rgbColor) =>
        this._styles.SetIndexedColors(rgbColor);

    partial void OnMRUColorsParsed(List<XLColor> color) => this._styles.SetMruColors(color);

    private Xpr<XLColor> ParseColor(string elementName, string ns)
    {
        if (!this._reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail<XLColor>();
        }

        return Xpr.From(this._reader.ParseColor(elementName, ns));
    }

    private Xpr ParseExtensionList(string elementName, string ns)
    {
        if (!this._reader.TryOpen(elementName, ns))
        {
            return Xpr.Fail();
        }

        this._reader.Skip(elementName);
        return Xpr.Success();
    }

    private static XLDifferentialProtectionValue OnCellProtectionParsed(
        bool? locked,
        bool? hidden
    ) => new() { Locked = locked, Hidden = hidden };

    private static XLFillFormatValue OnFillPatternFillParsed(XLFillFormatValue patternFillValue) =>
        patternFillValue;

    private static XLFillFormatValue OnFillGradientFillParsed(
        XLFillFormatValue gradientFillValue
    ) => gradientFillValue;

    /// <summary>
    /// A mapping of <c>ST_TableStyleType</c>. Custom enum mapping due to table/pivot duality.
    /// </summary>
    private static readonly Dictionary<string, (TS?, PTS?)> TableStyleTypeMap = new()
    {
        { "wholeTable", (TS.WholeTable, PTS.WholeTable) },
        { "headerRow", (TS.HeaderRow, PTS.HeaderRow) },
        { "totalRow", (TS.TotalRow, PTS.GrandTotalRow) },
        { "firstColumn", (TS.FirstColumn, PTS.FirstColumn) },
        { "lastColumn", (TS.LastColumn, PTS.GrandTotalColumn) },
        { "firstRowStripe", (TS.FirstRowStripe, PTS.FirstRowStripe) },
        { "secondRowStripe", (TS.SecondRowStripe, PTS.SecondRowStripe) },
        { "firstColumnStripe", (TS.FirstColumnStripe, PTS.FirstColumnStripe) },
        { "secondColumnStripe", (TS.SecondColumnStripe, PTS.SecondColumnStripe) },
        { "firstHeaderCell", (TS.FirstHeaderCell, PTS.FirstHeaderCell) },
        { "lastHeaderCell", (TS.LastHeaderCell, null) },
        { "firstTotalCell", (TS.FirstTotalCell, null) },
        { "lastTotalCell", (TS.LastTotalCell, null) },
        { "firstSubtotalColumn", (null, PTS.SubtotalColumn1) },
        { "secondSubtotalColumn", (null, PTS.SubtotalColumn2) },
        { "thirdSubtotalColumn", (null, PTS.SubtotalColumn3) },
        { "firstSubtotalRow", (null, PTS.SubtotalRow1) },
        { "secondSubtotalRow", (null, PTS.SubtotalRow2) },
        { "thirdSubtotalRow", (null, PTS.SubtotalRow3) },
        { "blankRow", (null, PTS.BlankRow) },
        { "firstColumnSubheading", (null, PTS.ColumnSubheading1) },
        { "secondColumnSubheading", (null, PTS.ColumnSubheading2) },
        { "thirdColumnSubheading", (null, PTS.ColumnSubheading3) },
        { "firstRowSubheading", (null, PTS.RowSubheading1) },
        { "secondRowSubheading", (null, PTS.RowSubheading2) },
        { "thirdRowSubheading", (null, PTS.RowSubheading3) },
        { "pageFieldLabels", (null, PTS.PageFieldLabels) },
        { "pageFieldValues", (null, PTS.PageFieldValues) },
    };
}

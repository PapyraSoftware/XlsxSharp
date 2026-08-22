namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A differential font format.
/// </summary>
internal record XLDifferentialFontValue
{
    /// <summary>
    /// A value with all properties null.
    /// </summary>
    internal static XLDifferentialFontValue Empty = new()
    {
        Name = null,
        Charset = null,
        Family = null,
        Bold = null,
        Italic = null,
        Strikethrough = null,
        Outline = null,
        Shadow = null,
        Condense = null,
        Extend = null,
        Color = null,
        Size = null,
        Underline = null,
        VerticalAlignment = null,
        Scheme = null,
    };

    internal required XLFontName? Name { get; init; }

    internal required XLFontCharSet? Charset { get; init; }

    internal required XLFontFamilyNumberingValues? Family { get; init; }

    internal required bool? Bold { get; init; }

    internal required bool? Italic { get; init; }

    internal required bool? Strikethrough { get; init; }

    internal required bool? Outline { get; init; }

    internal required bool? Shadow { get; init; }

    internal required bool? Condense { get; init; }

    internal required bool? Extend { get; init; }

    internal required XLColor? Color { get; init; }

    internal required XLFontSize? Size { get; init; }

    internal required XLFontUnderlineValues? Underline { get; init; }

    internal required XLFontVerticalTextAlignmentValues? VerticalAlignment { get; init; }

    internal required XLFontScheme? Scheme { get; init; }

    internal bool IsEmpty() =>
        this.Name is null
        && this.Charset is null
        && this.Family is null
        && this.Bold is null
        && this.Italic is null
        && this.Strikethrough is null
        && this.Outline is null
        && this.Shadow is null
        && this.Condense is null
        && this.Extend is null
        && this.Color is null
        && this.Size is null
        && this.Underline is null
        && this.VerticalAlignment is null
        && this.Scheme is null;
}

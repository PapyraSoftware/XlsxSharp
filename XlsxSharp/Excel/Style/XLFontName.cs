using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// A font name, two font names are equal when they are case insensitive equal. It is a custom
/// class because that way <see cref="XLFontFormatValue"/> and other structures don't have to implement
/// custom hash code and equality methods.
/// </summary>
internal readonly record struct XLFontName : IEquatable<string>
{
    private const StringComparison Comparison = StringComparison.OrdinalIgnoreCase;

    private XLFontName(string text)
    {
        // Spec says at most 31 chars, Excel also tries to repair workbook when value is longer.
        if (string.IsNullOrWhiteSpace(text) || text.Length > 31)
        {
            throw new ArgumentException(
                "Font name can't be empty and must be less than 32 characters long.",
                nameof(text)
            );
        }

        this.Text = text;
    }

    public string Text { get; }

    public bool Equals(string other)
    {
        return string.Equals(this.Text, other, Comparison);
    }

    public override int GetHashCode()
    {
        return this.Text.GetHashCode(Comparison);
    }

    public bool Equals(XLFontName other)
    {
        return this.Equals(other.Text);
    }

    public static implicit operator XLFontName(string text) => new(text);
}

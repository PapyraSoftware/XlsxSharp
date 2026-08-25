using System.Globalization;

namespace XlsxSharp.ExcelNumberFormat;

/// <summary>
/// Parse ECMA-376 number format strings and format values like Excel and other spreadsheet softwares.
/// </summary>
public class NumberFormat
{
    /// <summary>
    /// Initializes a new instance of the <see cref="NumberFormat"/> class.
    /// </summary>
    /// <param name="formatString">The number format string.</param>
    public NumberFormat(string formatString)
    {
        List<Section> sections = Parser.ParseSections(formatString, out bool syntaxError);

        this.IsValid = !syntaxError;
        this.FormatString = formatString;

        if (this.IsValid)
        {
            this.Sections = sections;
            this.IsDateTimeFormat =
                Evaluator.GetFirstSection(this.Sections, SectionType.Date) != null;
            this.IsTimeSpanFormat =
                Evaluator.GetFirstSection(this.Sections, SectionType.Duration) != null;
        }
        else
        {
            this.Sections = new List<Section>();
        }
    }

    /// <summary>
    /// Gets a value indicating whether the number format string is valid.
    /// </summary>
    public bool IsValid { get; }

    /// <summary>
    /// Gets the number format string.
    /// </summary>
    public string FormatString { get; }

    /// <summary>
    /// Gets a value indicating whether the format represents a DateTime
    /// </summary>
    public bool IsDateTimeFormat { get; }

    /// <summary>
    /// Gets a value indicating whether the format represents a TimeSpan
    /// </summary>
    public bool IsTimeSpanFormat { get; }

    internal List<Section> Sections { get; }

    /// <summary>
    /// Formats a value with this number format in a specified culture.
    /// </summary>
    /// <param name="value">The value to format.</param>
    /// <param name="culture">The culture to use for formatting.</param>
    /// <param name="isDate1904">If false, numeric dates start on January 0 1900 and include February 29 1900 - like Excel on PC. If true, numeric dates start on January 1 1904 - like Excel on Mac.</param>
    /// <returns>The formatted string.</returns>
    public string Format(object value, CultureInfo culture, bool isDate1904 = false)
    {
        Section? section = Evaluator.GetSection(this.Sections, value);
        if (section == null)
        {
            return CompatibleConvert.ToString(value, culture);
        }

        try
        {
            return Formatter.Format(value, section, culture, isDate1904);
        }
        catch (InvalidCastException)
        {
            // TimeSpan cast exception
            return CompatibleConvert.ToString(value, culture);
        }
        catch (FormatException)
        {
            // Convert.ToDouble/ToDateTime exceptions
            return CompatibleConvert.ToString(value, culture);
        }
    }
}

namespace XlsxSharp.Excel;

/// <summary>
/// Which kind of SpreadsheetML document a file is, which is what its extension says and what the
/// content type of its workbook part records.
/// </summary>
/// <remarks>
/// The save path maps each of these to its workbook part type in
/// <see cref="XlsxSharp.IO.Packaging.OoxmlPartTypes"/>, which carries the content type the
/// workbook part is declared with.
/// </remarks>
internal enum XLSpreadsheetDocumentType
{
    /// <summary>A <c>.xlsx</c> workbook.</summary>
    Workbook,

    /// <summary>A <c>.xltx</c> template.</summary>
    Template,

    /// <summary>A <c>.xlsm</c> macro enabled workbook.</summary>
    MacroEnabledWorkbook,

    /// <summary>A <c>.xltm</c> macro enabled template.</summary>
    MacroEnabledTemplate,
}

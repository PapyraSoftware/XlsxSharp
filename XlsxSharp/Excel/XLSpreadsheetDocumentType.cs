namespace XlsxSharp.Excel;

/// <summary>
/// Which kind of SpreadsheetML document a file is, which is what its extension says and what the
/// content type of its workbook part records.
/// </summary>
/// <remarks>
/// This replaces the SDK's <c>SpreadsheetDocumentType</c> in the workbook model, so that deciding
/// what a <c>.xlsm</c> is does not need the SDK. The save path still maps it to the SDK's enum
/// where it creates the package; <see cref="XlsxSharp.IO.Packaging.OoxmlPartTypes"/> carries the
/// content type each of these ends up with.
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

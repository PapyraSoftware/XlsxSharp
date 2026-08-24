#nullable disable

using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using DocumentFormat.OpenXml;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.CustomProperties;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.Graphics;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel;

public enum XLCalculateMode
{
    Auto,
    AutoNoTable,
    Manual,
    Default,
};

public enum XLReferenceStyle
{
    R1C1,
    A1,
    Default,
};

public enum XLCellSetValueBehavior
{
    /// <summary>
    ///   Analyze input string and convert value. For avoid analyzing use escape symbol '
    /// </summary>
    Smart = 0,

    /// <summary>
    ///   Direct set value. If value has unsupported type - value will be stored as string returned by <see
    ///    cref = "object.ToString()" />
    /// </summary>
    Simple = 1,
}

public sealed partial class XLWorkbook : IXLWorkbook
{
    private bool _disposed;

    #region Static

    public static double DefaultRowHeight { get; private set; }
    public static double DefaultColumnWidth { get; private set; }

    public static IXLPageSetup DefaultPageOptions
    {
        get
        {
            XLPageSetup defaultPageOptions = new(null, null)
            {
                PageOrientation = XLPageOrientation.Default,
                Scale = 100,
                PaperSize = XLPaperSize.LetterPaper,
                Margins = new XLMargins
                {
                    Top = 0.75,
                    Bottom = 0.5,
                    Left = 0.75,
                    Right = 0.75,
                    Header = 0.5,
                    Footer = 0.75,
                },
                ScaleHFWithDocument = true,
                AlignHFWithMargins = true,
                PrintErrorValue = XLPrintErrorValues.Displayed,
                ShowComments = XLShowCommentsValues.None,
            };
            return defaultPageOptions;
        }
    }

    public static IXLOutline DefaultOutline =>
        new XLOutline(null)
        {
            SummaryHLocation = XLOutlineSummaryHLocation.Right,
            SummaryVLocation = XLOutlineSummaryVLocation.Bottom,
        };

    /// <summary>
    ///   Behavior for <see cref = "IXLCell.set_Value" />
    /// </summary>
    public static XLCellSetValueBehavior CellSetValueBehavior { get; set; }

    public static XLWorkbook OpenFromTemplate(string path) => new(path, asTemplate: true);

    #endregion Static

    internal readonly List<UnsupportedSheet> UnsupportedSheets = [];

    internal IXLGraphicEngine GraphicEngine { get; }

    internal double DpiX { get; }

    internal double DpiY { get; }

    /// <inheritdoc cref="LoadOptions.StrictAttributeParsing"/>
    internal bool StrictAttributeParsing { get; }

    internal CancellationToken CancellationToken { get; }

    internal XLPivotCaches PivotCachesInternal
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        private init;
    }

    internal SharedStringTable SharedStringTable
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
    } = new();

    internal XLWorkbookStyles Styles
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
    }

    #region Nested Type : XLLoadSource

    private enum XLLoadSource
    {
        New,
        File,
        Stream,
    };

    #endregion Nested Type : XLLoadSource

    internal XLWorksheets WorksheetsInternal
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        private init;
    }

    /// <summary>
    ///   Gets an object to manipulate the worksheets.
    /// </summary>
    public IXLWorksheets Worksheets => this.WorksheetsInternal;

    internal XLDefinedNames DefinedNamesInternal
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
    }

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    public IXLDefinedNames NamedRanges => this.DefinedNamesInternal;

    /// <summary>
    ///   Gets an object to manipulate this workbook's named ranges.
    /// </summary>
    public IXLDefinedNames DefinedNames => this.DefinedNamesInternal;

    /// <summary>
    ///   Gets an object to manipulate this workbook's theme.
    /// </summary>
    public IXLTheme Theme
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        private set
        {
            this.ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    /// All pivot caches in the workbook, whether they have a pivot table or not.
    /// </summary>
    public IXLPivotCaches PivotCaches => this.PivotCachesInternal;

    /// <inheritdoc/>
    public IXLStyle Style
    {
        get
        {
            this.ThrowIfDisposed();
            return this.Format;
        }
        set
        {
            this.ThrowIfDisposed();
            this.Format.SetStyle(value);
        }
    }

    /// <summary>
    ///   Gets or sets the default row height for the workbook.
    ///   <para>All new worksheets will use this row height.</para>
    /// </summary>
    public double RowHeight { get; set; }

    /// <summary>
    ///   Gets or sets the default column width for the workbook.
    ///   <para>All new worksheets will use this column width.</para>
    /// </summary>
    public double ColumnWidth { get; set; }

    /// <summary>
    ///   Gets or sets the default page options for the workbook.
    ///   <para>All new worksheets will use these page options.</para>
    /// </summary>
    public IXLPageSetup PageOptions
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        set
        {
            this.ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    ///   Gets or sets the default outline options for the workbook.
    ///   <para>All new worksheets will use these outline options.</para>
    /// </summary>
    public IXLOutline Outline
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        set
        {
            this.ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    ///   Gets or sets the workbook's properties.
    /// </summary>
    public XLWorkbookProperties Properties
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        set
        {
            this.ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    ///   Gets or sets the workbook's calculation mode.
    /// </summary>
    public XLCalculateMode CalculateMode { get; set; }

    public bool CalculationOnSave { get; set; }
    public bool ForceFullCalculation { get; set; }
    public bool FullCalculationOnLoad { get; set; }
    public bool FullPrecision { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's reference style.
    /// </summary>
    public XLReferenceStyle ReferenceStyle { get; set; }

    public IXLCustomProperties CustomProperties
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        private init;
    }

    public bool ShowFormulas { get; set; }
    public bool ShowGridLines { get; set; }
    public bool ShowOutlineSymbols { get; set; }
    public bool ShowRowColHeaders { get; set; }
    public bool ShowRuler { get; set; }
    public bool ShowWhiteSpace { get; set; }
    public bool ShowZeros { get; set; }
    public bool RightToLeft { get; set; }

    public bool DefaultShowFormulas => false;

    public bool DefaultShowGridLines => true;

    public bool DefaultShowOutlineSymbols => true;

    public bool DefaultShowRowColHeaders => true;

    public bool DefaultShowRuler => true;

    public bool DefaultShowWhiteSpace => true;

    public bool DefaultShowZeros => true;

    public IXLFileSharing FileSharing
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
    } = new XLFileSharing();

    public bool DefaultRightToLeft => false;

    private void InitializeTheme() =>
        this.Theme = new XLTheme
        {
            Text1 = XLColor.FromHtml("#FF000000"),
            Background1 = XLColor.FromHtml("#FFFFFFFF"),
            Text2 = XLColor.FromHtml("#FF1F497D"),
            Background2 = XLColor.FromHtml("#FFEEECE1"),
            Accent1 = XLColor.FromHtml("#FF4F81BD"),
            Accent2 = XLColor.FromHtml("#FFC0504D"),
            Accent3 = XLColor.FromHtml("#FF9BBB59"),
            Accent4 = XLColor.FromHtml("#FF8064A2"),
            Accent5 = XLColor.FromHtml("#FF4BACC6"),
            Accent6 = XLColor.FromHtml("#FFF79646"),
            Hyperlink = XLColor.FromHtml("#FF0000FF"),
            FollowedHyperlink = XLColor.FromHtml("#FF800080"),
        };

#nullable enable
    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    public IXLDefinedName? NamedRange(string name) => this.DefinedName(name);

    /// <inheritdoc/>
    public IXLDefinedName? DefinedName(string name)
    {
        this.ThrowIfDisposed();
        if (name.Contains('!'))
        {
            string[] split = name.Split('!');
            string first = split[0];
            string wsName = first.StartsWith('\'') ? first.Substring(1, first.Length - 2) : first;
            string sheetlessName = split[1];
            if (this.TryGetWorksheet(wsName, out XLWorksheet ws))
            {
                if (
                    ws.DefinedNames.TryGetScopedValue(
                        sheetlessName,
                        out XLDefinedName? sheetDefinedName
                    )
                )
                {
                    return sheetDefinedName;
                }
            }

            name = sheetlessName;
        }

        return this.DefinedNamesInternal.TryGetScopedValue(name, out XLDefinedName? definedName)
            ? definedName
            : null;
    }

#nullable disable

    public bool TryGetWorksheet(string name, out IXLWorksheet worksheet)
    {
        this.ThrowIfDisposed();
        if (this.TryGetWorksheet(name, out XLWorksheet foundSheet))
        {
            worksheet = foundSheet;
            return true;
        }

        worksheet = default;
        return false;
    }

    internal bool TryGetWorksheet(string name, [NotNullWhen(true)] out XLWorksheet worksheet)
    {
        this.ThrowIfDisposed();
        return this.WorksheetsInternal.TryGetWorksheet(name, out worksheet);
    }

    public IXLRange RangeFromFullAddress(string rangeAddress, out IXLWorksheet ws)
    {
        this.ThrowIfDisposed();
        if (!rangeAddress.Contains('!'))
        {
            ws = null;
            return null;
        }

        string[] split = rangeAddress.Split('!');
        string wsName = split[0].UnescapeSheetName();
        if (this.TryGetWorksheet(wsName, out XLWorksheet sheet))
        {
            ws = sheet;
            return sheet.Range(split[1]);
        }

        ws = null;
        return null;
    }

    public IXLCell CellFromFullAddress(string cellAddress, out IXLWorksheet ws)
    {
        this.ThrowIfDisposed();
        if (!cellAddress.Contains('!'))
        {
            ws = null;
            return null;
        }

        string[] split = cellAddress.Split('!');
        string wsName = split[0].UnescapeSheetName();
        if (this.TryGetWorksheet(wsName, out XLWorksheet sheet))
        {
            ws = sheet;
            return sheet.Cell(split[1]);
        }

        ws = null;
        return null;
    }

    /// <summary>
    ///   Saves the current workbook.
    /// </summary>
    public void Save()
    {
        this.ThrowIfDisposed();
#if DEBUG
        this.Save(true);
#else
        Save(false, false);
#endif
    }

    /// <summary>
    ///   Saves the current workbook and optionally performs validation
    /// </summary>
    public void Save(bool validate, bool evaluateFormulae = false)
    {
        this.ThrowIfDisposed();
        this.Save(
            new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true,
            }
        );
    }

    public void Save(SaveOptions options)
    {
        this.ThrowIfDisposed();
        this.checkForWorksheetsPresent();
        if (this._loadSource == XLLoadSource.New)
        {
            throw new InvalidOperationException(
                "This is a new file. Please use one of the 'SaveAs' methods."
            );
        }

        if (this._loadSource == XLLoadSource.Stream)
        {
            this.CreatePackage(this._originalStream, false, this._spreadsheetDocumentType, options);
        }
        else
        {
            this.CreatePackage(this._originalFile, this._spreadsheetDocumentType, options);
        }
    }

    /// <summary>
    ///   Saves the current workbook to a file.
    /// </summary>
    public void SaveAs(string file)
    {
        this.ThrowIfDisposed();
#if DEBUG
        this.SaveAs(file, true);
#else
        SaveAs(file, false, false);
#endif
    }

    /// <summary>
    ///   Saves the current workbook to a file and optionally validates it.
    /// </summary>
    public void SaveAs(string file, bool validate, bool evaluateFormulae = false)
    {
        this.ThrowIfDisposed();
        this.SaveAs(
            file,
            new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true,
            }
        );
    }

    public void SaveAs(string file, SaveOptions options)
    {
        this.ThrowIfDisposed();
        this.checkForWorksheetsPresent();

        string directoryName = Path.GetDirectoryName(file);
        if (!string.IsNullOrWhiteSpace(directoryName))
        {
            Directory.CreateDirectory(directoryName);
        }

        if (this._loadSource == XLLoadSource.New)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }

            this.CreatePackage(file, GetSpreadsheetDocumentType(file), options);
        }
        else if (this._loadSource == XLLoadSource.File)
        {
            if (string.Compare(this._originalFile.Trim(), file.Trim(), true) != 0)
            {
                File.Copy(this._originalFile, file, true);
                File.SetAttributes(file, FileAttributes.Normal);
            }

            this.CreatePackage(file, GetSpreadsheetDocumentType(file), options);
        }
        else if (this._loadSource == XLLoadSource.Stream)
        {
            this._originalStream.Position = 0;

            using (FileStream fileStream = File.Create(file))
            {
                CopyStream(this._originalStream, fileStream);
                this.CreatePackage(fileStream, false, this._spreadsheetDocumentType, options);
            }
        }

        this._loadSource = XLLoadSource.File;
        this._originalFile = file;
        this._originalStream = null;
    }

    private static SpreadsheetDocumentType GetSpreadsheetDocumentType(string filePath)
    {
        string extension = Path.GetExtension(filePath);

        if (string.IsNullOrEmpty(extension))
        {
            throw new ArgumentException("Empty extension is not supported.");
        }

        extension = extension.Substring(1).ToLowerInvariant();

        switch (extension)
        {
            case "xlsm":
                return SpreadsheetDocumentType.MacroEnabledWorkbook;

            case "xltm":
                return SpreadsheetDocumentType.MacroEnabledTemplate;

            case "xlsx":
                return SpreadsheetDocumentType.Workbook;

            case "xltx":
                return SpreadsheetDocumentType.Template;

            default:
                throw new ArgumentException(
                    string.Format(
                        "Extension '{0}' is not supported. Supported extensions are '.xlsx', '.xlsm', '.xltx' and '.xltm'.",
                        extension
                    )
                );
        }
    }

    private void checkForWorksheetsPresent()
    {
        if (!this.Worksheets.Any())
        {
            throw new InvalidOperationException("Workbooks need at least one worksheet.");
        }
    }

    /// <summary>
    ///   Saves the current workbook to a stream.
    /// </summary>
    public void SaveAs(Stream stream)
    {
        this.ThrowIfDisposed();
#if DEBUG
        this.SaveAs(stream, true);
#else
        SaveAs(stream, false, false);
#endif
    }

    /// <summary>
    ///   Saves the current workbook to a stream and optionally validates it.
    /// </summary>
    public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false)
    {
        this.ThrowIfDisposed();
        this.SaveAs(
            stream,
            new SaveOptions
            {
                ValidatePackage = validate,
                EvaluateFormulasBeforeSaving = evaluateFormulae,
                GenerateCalculationChain = true,
            }
        );
    }

    public void SaveAs(Stream stream, SaveOptions options)
    {
        this.ThrowIfDisposed();
        this.checkForWorksheetsPresent();
        if (this._loadSource == XLLoadSource.New)
        {
            // dm 20130422, this method or better the method SpreadsheetDocument.Create which is called
            // inside of 'CreatePackage' need a stream which CanSeek & CanRead
            // and an ordinary Response stream of a webserver can't do this
            // so we have to ask and provide a way around this
            if (stream.CanRead && stream.CanSeek && stream.CanWrite)
            {
                // all is fine the package can be created in a direct way
                this.CreatePackage(stream, true, this._spreadsheetDocumentType, options);
            }
            else
            {
                // the harder way
                using (MemoryStream ms = new())
                {
                    this.CreatePackage(ms, true, this._spreadsheetDocumentType, options);
                    // not really necessary, because I changed CopyStream too.
                    // but for better understanding and if somebody in the future
                    // provide an changed version of CopyStream
                    ms.Position = 0;
                    CopyStream(ms, stream);
                }
            }
        }
        else if (this._loadSource == XLLoadSource.File)
        {
            using (FileStream fileStream = new(this._originalFile, FileMode.Open, FileAccess.Read))
            {
                CopyStream(fileStream, stream);
            }
            this.CreatePackage(stream, false, this._spreadsheetDocumentType, options);
        }
        else if (this._loadSource == XLLoadSource.Stream)
        {
            this._originalStream.Position = 0;
            if (this._originalStream != stream)
            {
                CopyStream(this._originalStream, stream);
            }

            this.CreatePackage(stream, false, this._spreadsheetDocumentType, options);
        }

        this._loadSource = XLLoadSource.Stream;
        this._originalStream = stream;
        this._originalFile = null;
    }

    internal static void CopyStream(Stream input, Stream output)
    {
        byte[] buffer = new byte[8 * 1024];
        int len;
        // dm 20130422, it is always a good idea to rewind the input stream, or not?
        if (input.CanSeek)
        {
            input.Seek(0, SeekOrigin.Begin);
        }

        while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            output.Write(buffer, 0, len);
        }

        // dm 20130422, and flushing the output after write
        output.Flush();
    }

    public IXLTable Table(
        string tableName,
        StringComparison comparisonType = StringComparison.OrdinalIgnoreCase
    )
    {
        this.ThrowIfDisposed();
        if (!this.TryGetTable(tableName, out XLTable table, comparisonType))
        {
            throw new ArgumentOutOfRangeException($"Table {tableName} was not found.");
        }

        return table;
    }

    /// <summary>
    /// Try to find a table with <paramref name="tableName"/> in a workbook.
    /// </summary>
    internal bool TryGetTable(
        string tableName,
        out XLTable table,
        StringComparison comparisonType = StringComparison.OrdinalIgnoreCase
    )
    {
        this.ThrowIfDisposed();
        table = this
            .WorksheetsInternal.SelectMany<XLWorksheet, XLTable>(ws => ws.Tables)
            .FirstOrDefault(t => t.Name.Equals(tableName, comparisonType));

        return table is not null;
    }

    /// <summary>
    /// Try to find a table that covers same area as the <paramref name="area"/> in a workbook.
    /// </summary>
    internal bool TryGetTable(SheetArea area, out XLTable foundTable)
    {
        this.ThrowIfDisposed();
        foreach (XLWorksheet sheet in this.WorksheetsInternal)
        {
            if (XlsxSharp.XLHelper.SheetComparer.Equals(sheet.Name, area.Name))
            {
                foreach (XLTable table in sheet.Tables)
                {
                    if (table.Area != area.Area)
                    {
                        continue;
                    }

                    foundTable = table;
                    return true;
                }

                // No other sheet has correct name.
                break;
            }
        }

        foundTable = null;
        return false;
    }

    public IXLWorksheet Worksheet(string name)
    {
        this.ThrowIfDisposed();
        return this.WorksheetsInternal.Worksheet(name);
    }

    public IXLWorksheet Worksheet(int position)
    {
        this.ThrowIfDisposed();
        return this.WorksheetsInternal.Worksheet(position);
    }

    public IXLCustomProperty CustomProperty(string name)
    {
        this.ThrowIfDisposed();
        return this.CustomProperties.CustomProperty(name);
    }

    public IXLCells FindCells(Func<IXLCell, bool> predicate)
    {
        this.ThrowIfDisposed();
        XLCells cells = new(this, false, XLCellsUsedOptions.AllContents);
        foreach (XLWorksheet ws in this.WorksheetsInternal)
        {
            foreach (XLCell cell in ws.CellsUsed(XLCellsUsedOptions.All))
            {
                if (predicate(cell))
                {
                    cells.Add(cell);
                }
            }
        }
        return cells;
    }

    public IXLRows FindRows(Func<IXLRow, bool> predicate)
    {
        this.ThrowIfDisposed();
        XLRows rows = new(this, worksheet: null, defaultStyleSheet: null);
        foreach (XLWorksheet ws in this.WorksheetsInternal)
        {
            foreach (IXLRow row in ws.Rows().Where(predicate))
            {
                rows.Add(row as XLRow);
            }
        }
        return rows;
    }

    public IXLColumns FindColumns(Func<IXLColumn, bool> predicate)
    {
        this.ThrowIfDisposed();
        XLColumns columns = new(this, worksheet: null, defaultStyleSheet: null);
        foreach (XLWorksheet ws in this.WorksheetsInternal)
        {
            foreach (IXLColumn column in ws.Columns().Where(predicate))
            {
                columns.Add(column as XLColumn);
            }
        }
        return columns;
    }

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    public IEnumerable<IXLCell> Search(
        string searchText,
        CompareOptions compareOptions = CompareOptions.Ordinal,
        bool searchFormulae = false
    )
    {
        this.ThrowIfDisposed();
        foreach (XLWorksheet ws in this.WorksheetsInternal)
        {
            foreach (IXLCell cell in ws.Search(searchText, compareOptions, searchFormulae))
            {
                yield return cell;
            }
        }
    }

    #region Fields

    private XLLoadSource _loadSource = XLLoadSource.New;
    private string _originalFile;
    private Stream _originalStream;
    private XLWorkbookProtection _workbookProtection;

    #endregion Fields

    #region Constructor

    /// <summary>
    ///   Creates a new Excel workbook.
    /// </summary>
    public XLWorkbook()
        : this(new LoadOptions()) { }

    internal XLWorkbook(string file, bool asTemplate)
        : this(new LoadOptions())
    {
        this.Styles = new XLWorkbookStyles();
        this.LoadSheetsFromTemplate(file);
    }

    /// <summary>
    ///   Opens an existing workbook from a file.
    /// </summary>
    /// <param name = "file">The file to open.</param>
    public XLWorkbook(string file)
        : this(file, new LoadOptions()) { }

    public XLWorkbook(string file, LoadOptions loadOptions)
        : this(loadOptions)
    {
        this._loadSource = XLLoadSource.File;
        this._originalFile = file;
        this._spreadsheetDocumentType = GetSpreadsheetDocumentType(this._originalFile);
        this.Styles = new XLWorkbookStyles();
        this.Load(file);

        if (loadOptions.RecalculateAllFormulas)
        {
            this.RecalculateAllFormulas();
        }
    }

    /// <summary>
    ///   Opens an existing workbook from a stream.
    /// </summary>
    /// <param name = "stream">The stream to open.</param>
    public XLWorkbook(Stream stream)
        : this(stream, new LoadOptions()) { }

    public XLWorkbook(Stream stream, LoadOptions loadOptions)
        : this(loadOptions)
    {
        this._loadSource = XLLoadSource.Stream;
        this._originalStream = stream;
        this.Styles = new XLWorkbookStyles();
        this.Load(stream);

        if (loadOptions.RecalculateAllFormulas)
        {
            this.RecalculateAllFormulas();
        }
    }

    public XLWorkbook(LoadOptions loadOptions)
    {
        ArgumentNullException.ThrowIfNull(loadOptions);

        this.DpiX = loadOptions.Dpi.X;
        this.DpiY = loadOptions.Dpi.Y;
        this.StrictAttributeParsing = loadOptions.StrictAttributeParsing;
        this.GraphicEngine =
            loadOptions.GraphicEngine
            ?? LoadOptions.DefaultGraphicEngine
            ?? DefaultGraphicEngine.Instance.Value;
        this.CancellationToken = loadOptions.CancellationToken;
        this.Protection = new XLWorkbookProtection(DefaultProtectionAlgorithm);
        DefaultRowHeight = 15;
        DefaultColumnWidth = 8.43;
        this.Styles = XLWorkbookStyles.CreateInitialized();
        this.RowHeight = DefaultRowHeight;
        this.ColumnWidth = DefaultColumnWidth;
        this.PageOptions = DefaultPageOptions;
        this.Outline = DefaultOutline;
        this.Properties = new XLWorkbookProperties();
        this.CalculateMode = XLCalculateMode.Default;
        this.ReferenceStyle = XLReferenceStyle.Default;
        this.InitializeTheme();
        this.ShowFormulas = this.DefaultShowFormulas;
        this.ShowGridLines = this.DefaultShowGridLines;
        this.ShowOutlineSymbols = this.DefaultShowOutlineSymbols;
        this.ShowRowColHeaders = this.DefaultShowRowColHeaders;
        this.ShowRuler = this.DefaultShowRuler;
        this.ShowWhiteSpace = this.DefaultShowWhiteSpace;
        this.ShowZeros = this.DefaultShowZeros;
        this.RightToLeft = this.DefaultRightToLeft;
        this.WorksheetsInternal = new XLWorksheets(this);
        this.DefinedNamesInternal = new XLDefinedNames(this);
        this.PivotCachesInternal = new XLPivotCaches(this);
        this.CustomProperties = new XLCustomProperties(this);
        this.ShapeIdManager = new XLIdManager();
        this.Author = Environment.UserName;
    }

    #endregion Constructor

    #region Nested type: UnsupportedSheet

    internal sealed class UnsupportedSheet
    {
        public bool IsActive;
        public uint SheetId;
        public int Position;
    }

    #endregion Nested type: UnsupportedSheet

    public IXLCell Cell(string namedCell)
    {
        this.ThrowIfDisposed();
        IXLDefinedName namedRange = this.DefinedName(namedCell);
        if (namedRange != null)
        {
            return namedRange.Ranges?.FirstOrDefault()?.FirstCell();
        }
        else
        {
            return this.CellFromFullAddress(namedCell, out _);
        }
    }

    public IXLCells Cells(string namedCells)
    {
        this.ThrowIfDisposed();
        return this.Ranges(namedCells).Cells();
    }

    public IXLRange Range(string range)
    {
        this.ThrowIfDisposed();
        IXLDefinedName namedRange = this.DefinedName(range);
        if (namedRange != null)
        {
            return namedRange.Ranges.FirstOrDefault();
        }
        else
        {
            return this.RangeFromFullAddress(range, out _);
        }
    }

    public IXLRanges Ranges(string ranges)
    {
        this.ThrowIfDisposed();
        XLRanges retVal = new(this);
        string[] rangePairs = ranges.Split(',');
        foreach (
            IXLRange range in rangePairs
                .Select(r => this.Range(r.Trim()))
                .Where(range => range != null)
        )
        {
            retVal.Add(range);
        }
        return retVal;
    }

    internal XLIdManager ShapeIdManager
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        private set
        {
            this.ThrowIfDisposed();
            field = value;
        }
    }

    public void Dispose()
    {
        if (this._disposed)
        {
            return;
        }

        foreach (XLWorksheet worksheet in this.WorksheetsInternal)
        {
            worksheet.Cleanup();
        }

        this._disposed = true;
    }

    public bool Use1904DateSystem { get; set; }

    public XLWorkbook SetUse1904DateSystem() => this.SetUse1904DateSystem(true);

    public XLWorkbook SetUse1904DateSystem(bool value)
    {
        this.Use1904DateSystem = value;
        return this;
    }

    public IXLWorksheet AddWorksheet()
    {
        this.ThrowIfDisposed();
        return this.Worksheets.Add();
    }

    public IXLWorksheet AddWorksheet(int position)
    {
        this.ThrowIfDisposed();
        return this.Worksheets.Add(position);
    }

    public IXLWorksheet AddWorksheet(string sheetName)
    {
        this.ThrowIfDisposed();
        return this.Worksheets.Add(sheetName);
    }

    public IXLWorksheet AddWorksheet(string sheetName, int position)
    {
        this.ThrowIfDisposed();
        return this.Worksheets.Add(sheetName, position);
    }

    public void AddWorksheet(DataSet dataSet)
    {
        this.ThrowIfDisposed();
        this.Worksheets.Add(dataSet);
    }

    public void AddWorksheet(IXLWorksheet worksheet)
    {
        this.ThrowIfDisposed();
        worksheet.CopyTo(this, worksheet.Name);
    }

    public IXLWorksheet AddWorksheet(DataTable dataTable)
    {
        this.ThrowIfDisposed();
        return this.Worksheets.Add(dataTable);
    }

    public IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName)
    {
        this.ThrowIfDisposed();
        return this.Worksheets.Add(dataTable, sheetName);
    }

    public IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName, string tableName)
    {
        this.ThrowIfDisposed();
        return this.Worksheets.Add(dataTable, sheetName, tableName);
    }

    private XLCalcEngine _calcEngine;

    internal XLCalcEngine CalcEngine =>
        this._calcEngine ??= new XLCalcEngine(CultureInfo.CurrentCulture);

    public XLCellValue Evaluate(string expression)
    {
        this.ThrowIfDisposed();
        return this.CalcEngine.EvaluateFormula(expression, this).ToCellValue();
    }

    /// <summary>
    /// Force recalculation of all cell formulas.
    /// </summary>
    public void RecalculateAllFormulas()
    {
        this.ThrowIfDisposed();
        foreach (XLWorksheet sheet in this.WorksheetsInternal)
        {
            sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(Area.Full);
        }

        this.CalcEngine.Recalculate(this, null);
    }

    private static XLCalcEngine _calcEngineExpr;
    private SpreadsheetDocumentType _spreadsheetDocumentType;

    private static XLCalcEngine CalcEngineExpr =>
        _calcEngineExpr ??= new XLCalcEngine(CultureInfo.InvariantCulture);

    /// <summary>
    /// Evaluate a formula and return a value. Formulas with references don't work and culture used for conversion is invariant.
    /// </summary>
    public static XLCellValue EvaluateExpr(string expression) =>
        CalcEngineExpr.EvaluateFormula(expression).ToCellValue();

    /// <summary>
    /// Evaluate a formula and return a value. Use current culture.
    /// </summary>
    internal static XLCellValue EvaluateExprCurrent(string expression) =>
        new XLCalcEngine(CultureInfo.CurrentCulture).EvaluateFormula(expression).ToCellValue();

    public string Author { get; set; }

    public bool LockStructure
    {
        get
        {
            this.ThrowIfDisposed();
            return this.Protection.IsProtected
                && !this.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure);
        }
        set
        {
            this.ThrowIfDisposed();
            if (!this.Protection.IsProtected)
            {
                throw new InvalidOperationException(
                    $"Enable workbook protection before setting the {nameof(this.LockStructure)} property"
                );
            }

            this.Protection.AllowElement(XLWorkbookProtectionElements.Structure, value);
        }
    }

    public XLWorkbook SetLockStructure(bool value)
    {
        this.ThrowIfDisposed();
        this.LockStructure = value;
        return this;
    }

    public bool LockWindows
    {
        get
        {
            this.ThrowIfDisposed();
            return this.Protection.IsProtected
                && !this.Protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows);
        }
        set
        {
            this.ThrowIfDisposed();
            if (!this.Protection.IsProtected)
            {
                throw new InvalidOperationException(
                    $"Enable workbook protection before setting the {nameof(this.LockWindows)} property"
                );
            }

            this.Protection.AllowElement(XLWorkbookProtectionElements.Windows, value);
        }
    }

    public XLWorkbook SetLockWindows(bool value)
    {
        this.ThrowIfDisposed();
        this.LockWindows = value;
        return this;
    }

    public bool IsPasswordProtected => this.Protection.IsPasswordProtected;
    public bool IsProtected => this.Protection.IsProtected;

    IXLWorkbookProtection IXLProtectable<
        IXLWorkbookProtection,
        XLWorkbookProtectionElements
    >.Protection
    {
        get
        {
            this.ThrowIfDisposed();
            return this.Protection;
        }
        set
        {
            this.ThrowIfDisposed();
            this.Protection = value as XLWorkbookProtection;
        }
    }

    internal XLWorkbookProtection Protection
    {
        get
        {
            this.ThrowIfDisposed();
            return this._workbookProtection;
        }
        set
        {
            this.ThrowIfDisposed();
            this._workbookProtection = value.Clone().CastTo<XLWorkbookProtection>();
        }
    }

    public IXLWorkbookProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        this.ThrowIfDisposed();
        return this.Protection.Protect(algorithm);
    }

    public IXLWorkbookProtection Protect(XLWorkbookProtectionElements allowedElements)
    {
        this.ThrowIfDisposed();
        return this.Protection.Protect(allowedElements);
    }

    public IXLWorkbookProtection Protect(
        Algorithm algorithm,
        XLWorkbookProtectionElements allowedElements
    )
    {
        this.ThrowIfDisposed();
        return this.Protection.Protect(algorithm, allowedElements);
    }

    public IXLWorkbookProtection Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm
    )
    {
        this.ThrowIfDisposed();
        return this.Protect(password, algorithm, XLWorkbookProtectionElements.Windows);
    }

    public IXLWorkbookProtection Protect(
        string password,
        Algorithm algorithm,
        XLWorkbookProtectionElements allowedElements
    )
    {
        this.ThrowIfDisposed();
        return this.Protection.Protect(password, algorithm, allowedElements);
    }

    IXLElementProtection IXLProtectable.Protect(Algorithm algorithm)
    {
        this.ThrowIfDisposed();
        return this.Protect(algorithm);
    }

    IXLElementProtection IXLProtectable.Protect(string password, Algorithm algorithm)
    {
        this.ThrowIfDisposed();
        return this.Protect(password, algorithm);
    }

    IXLWorkbookProtection IXLProtectable<
        IXLWorkbookProtection,
        XLWorkbookProtectionElements
    >.Protect(XLWorkbookProtectionElements allowedElements)
    {
        this.ThrowIfDisposed();
        return this.Protect(allowedElements);
    }

    IXLWorkbookProtection IXLProtectable<
        IXLWorkbookProtection,
        XLWorkbookProtectionElements
    >.Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
    {
        this.ThrowIfDisposed();
        return this.Protect(algorithm, allowedElements);
    }

    IXLWorkbookProtection IXLProtectable<
        IXLWorkbookProtection,
        XLWorkbookProtectionElements
    >.Protect(string password, Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
    {
        this.ThrowIfDisposed();
        return this.Protect(password, algorithm, allowedElements);
    }

    public IXLWorkbookProtection Unprotect()
    {
        this.ThrowIfDisposed();
        return this.Protection.Unprotect();
    }

    public IXLWorkbookProtection Unprotect(string password)
    {
        this.ThrowIfDisposed();
        return this.Protection.Unprotect(password);
    }

    IXLElementProtection IXLProtectable.Unprotect()
    {
        this.ThrowIfDisposed();
        return this.Unprotect();
    }

    IXLElementProtection IXLProtectable.Unprotect(string password)
    {
        this.ThrowIfDisposed();
        return this.Unprotect(password);
    }

    /// <summary>
    /// Notify various component of a workbook that sheet has been added.
    /// </summary>
    internal void NotifyWorksheetAdded(XLWorksheet newSheet)
    {
        this.ThrowIfDisposed();
        this._calcEngine.OnAddedSheet(newSheet);
    }

    /// <summary>
    /// Notify various component of a workbook that sheet is about to be removed.
    /// </summary>
    internal void NotifyWorksheetDeleting(XLWorksheet sheet)
    {
        this.ThrowIfDisposed();
        this._calcEngine.OnDeletingSheet(sheet);
    }

    public override string ToString()
    {
        this.ThrowIfDisposed();
        switch (this._loadSource)
        {
            case XLLoadSource.New:
                return "XLWorkbook(new)";

            case XLLoadSource.File:
                return string.Format("XLWorkbook({0})", this._originalFile);

            case XLLoadSource.Stream:
                return string.Format("XLWorkbook({0})", this._originalStream.ToString());

            default:
                throw new NotImplementedException();
        }
    }

    internal XLCellFormat Format
    {
        get
        {
            this.ThrowIfDisposed();
            return XLCellFormat.ForWorkbook(this);
        }
    }

    private void ThrowIfDisposed()
    {
        if (this._disposed)
        {
            throw new ObjectDisposedException(nameof(XLWorkbook));
        }
    }
}

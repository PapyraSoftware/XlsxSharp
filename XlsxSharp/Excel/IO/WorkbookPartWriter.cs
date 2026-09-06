#nullable enable

using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Extensions;
using XlsxSharp.IO.Packaging;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Writes <c>xl/workbook.xml</c>.
/// </summary>
/// <remarks>
/// <para>
/// The part is patched, not rewritten: a workbook from Excel carries a fileVersion, external
/// references and an extLst that XlsxSharp does not model and that have to survive a save.
/// </para>
/// <para>
/// Writing happens in two steps, and they are not next to each other in the save. The
/// relationship ids of the sheets are handed out early, where the previous writer ran, because
/// every part created after that point takes the next id and moving the allocation would
/// renumber all of them. The XML itself is written at the end, once the pivot caches exist,
/// because the list of them belongs in this part.
/// </para>
/// </remarks>
internal class WorkbookPartWriter
{
    private static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace Rel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>Child order on the root, from the schema (ECMA-376 Part 1 §18.2.27).</summary>
    private static readonly string[] ElementOrder =
    [
        "fileVersion",
        "fileSharing",
        "workbookPr",
        "workbookProtection",
        "bookViews",
        "sheets",
        "functionGroups",
        "externalReferences",
        "definedNames",
        "calcPr",
        "oleSize",
        "customWorkbookViews",
        "pivotCaches",
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
        "extLst",
    ];

    /// <summary>
    /// Hands out a relationship id to every sheet that does not have one yet, keeping the ones a
    /// loaded workbook already used. Runs where the old writer ran, so that the ids of the parts
    /// created afterwards do not move.
    /// </summary>
    internal static void AssignSheetRelIds(XLWorkbook xlWorkbook, XLWorkbook.SaveContext context)
    {
        foreach (
            XLWorksheet xlSheet in xlWorkbook.WorksheetsInternal.OrderBy<XLWorksheet, int>(w =>
                w.Position
            )
        )
        {
            if (string.IsNullOrWhiteSpace(xlSheet.RelId))
            {
                // Sheet isn't from loaded file and hasn't been saved yet.
                xlSheet.RelId = context.RelIdGenerator.GetNext(XLWorkbook.RelType.Workbook);
            }
        }
    }

    internal static void GenerateContent(
        OpcPart workbookPart,
        XLWorkbook xlWorkbook,
        SaveOptions options,
        XLWorkbook.SaveContext context
    )
    {
        XDocument document = ReadExisting(workbookPart);
        XElement workbook = document.Root!;

        WriteWorkbookProperties(workbook, xlWorkbook, options);
        WriteFileSharing(workbook, xlWorkbook);
        WriteProtection(workbook, xlWorkbook);
        WriteSheetsAndViews(workbook, xlWorkbook);
        SetElement(workbook, "definedNames", BuildDefinedNames(workbook, xlWorkbook));
        WriteCalculationProperties(workbook, xlWorkbook);
        WritePivotCaches(workbook, xlWorkbook);

        using Stream partStream = workbookPart.GetWriteStream();
        using XmlWriter xml = XmlWriter.Create(
            partStream,
            new XmlWriterSettings { Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );

        document.Save(xml);
    }

    /// <summary>
    /// The part's document with the two prefixes the writer needs on the root, or a fresh one.
    /// </summary>
    private static XDocument ReadExisting(OpcPart part)
    {
        XElement? loaded = null;

        if (part.Length > 0)
        {
            using Stream stream = part.GetReadStream();
            try
            {
                loaded = XDocument.Load(stream).Root;
            }
            catch (XmlException)
            {
                // A workbook.xml we cannot read is one we are about to replace.
            }
        }

        XElement workbook = new(
            Main + "workbook",
            new XAttribute(XNamespace.Xmlns + "r", Rel.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "x", Main.NamespaceName)
        );

        if (loaded is not null)
        {
            foreach (XAttribute attribute in loaded.Attributes())
            {
                if (
                    attribute.IsNamespaceDeclaration
                    && (attribute.Name.LocalName == "xmlns" || IsDeclared(workbook, attribute))
                )
                {
                    continue;
                }

                workbook.Add(new XAttribute(attribute));
            }

            foreach (XElement child in loaded.Elements())
            {
                XElement copy = new(child);
                copy.DescendantsAndSelf()
                    .Attributes()
                    .Where(a => a.IsNamespaceDeclaration && a.Name.LocalName == "xmlns")
                    .ToList()
                    .ForEach(a => a.Remove());

                workbook.Add(copy);
            }

            HoistDeclarations(workbook);
        }

        return new XDocument(workbook);
    }

    private static bool IsDeclared(XElement root, XAttribute declaration) =>
        root.Attributes().Any(a => a.IsNamespaceDeclaration && a.Value == declaration.Value);

    /// <summary>
    /// Copies the namespace declarations of the descendants onto the root, which is what the SDK
    /// did when it re-serialised a part.
    /// </summary>
    private static void HoistDeclarations(XElement root)
    {
        foreach (XElement descendant in root.Descendants())
        {
            foreach (XAttribute attribute in descendant.Attributes().ToList())
            {
                if (
                    attribute.IsNamespaceDeclaration
                    && attribute.Name.LocalName != "xmlns"
                    && !root.Attributes()
                        .Any(a => a.IsNamespaceDeclaration && a.Name == attribute.Name)
                )
                {
                    root.Add(new XAttribute(attribute));
                }
            }
        }
    }

    private static void WriteWorkbookProperties(
        XElement workbook,
        XLWorkbook xlWorkbook,
        SaveOptions options
    )
    {
        XElement properties = Ensure(workbook, "workbookPr");

        if (properties.Attribute("codeName") is null)
        {
            properties.SetAttributeValue("codeName", "ThisWorkbook");
        }

        SetOptionalBool(properties, "date1904", xlWorkbook.Use1904DateSystem, false);
        if (options.FilterPrivacy.HasValue)
        {
            SetOptionalBool(properties, "filterPrivacy", options.FilterPrivacy.Value, false);
        }
    }

    private static void WriteFileSharing(XElement workbook, XLWorkbook xlWorkbook)
    {
        XElement fileSharing = Ensure(workbook, "fileSharing");

        SetOptionalBool(
            fileSharing,
            "readOnlyRecommended",
            xlWorkbook.FileSharing.ReadOnlyRecommended,
            false
        );

        fileSharing.SetAttributeValue(
            "userName",
            string.IsNullOrWhiteSpace(xlWorkbook.FileSharing.UserName)
                ? null
                : xlWorkbook.FileSharing.UserName
        );

        // An element with nothing in it says nothing, and the previous writer dropped it.
        if (!fileSharing.HasAttributes && !fileSharing.HasElements)
        {
            fileSharing.Remove();
        }
    }

    private static void WriteProtection(XElement workbook, XLWorkbook xlWorkbook)
    {
        if (!xlWorkbook.Protection.IsProtected)
        {
            workbook.Element(Main + "workbookProtection")?.Remove();
            return;
        }

        XElement protectionElement = Ensure(workbook, "workbookProtection");
        XLWorkbookProtection protection = xlWorkbook.Protection;

        // Whichever of the two ways of storing the password was used before, only one of them
        // goes back out.
        foreach (
            string attribute in new[]
            {
                "workbookPassword",
                "workbookAlgorithmName",
                "workbookHashValue",
                "workbookSpinCount",
                "workbookSaltValue",
            }
        )
        {
            protectionElement.Attribute(attribute)?.Remove();
        }

        if (protection.Algorithm == XLProtectionAlgorithm.Algorithm.SimpleHash)
        {
            if (!string.IsNullOrWhiteSpace(protection.PasswordHash))
            {
                protectionElement.SetAttributeValue("workbookPassword", protection.PasswordHash);
            }
        }
        else
        {
            protectionElement.SetAttributeValue(
                "workbookAlgorithmName",
                DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.ToDescription(
                    protection.Algorithm
                )
            );

            protectionElement.SetAttributeValue("workbookHashValue", protection.PasswordHash);
            protectionElement.SetAttributeValue(
                "workbookSpinCount",
                protection.SpinCount.ToInvariantString()
            );

            protectionElement.SetAttributeValue("workbookSaltValue", protection.Base64EncodedSalt);
        }

        SetOptionalBool(
            protectionElement,
            "lockStructure",
            !protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure),
            false
        );

        SetOptionalBool(
            protectionElement,
            "lockWindows",
            !protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows),
            false
        );
    }

    /// <summary>
    /// Rebuilds the sheet list in its final order and the book view that points into it.
    /// </summary>
    private static void WriteSheetsAndViews(XElement workbook, XLWorkbook xlWorkbook)
    {
        XElement sheets = Ensure(workbook, "sheets");
        XLWorksheets worksheets = xlWorkbook.WorksheetsInternal;

        // Drop the sheets that were deleted from the workbook since it was loaded.
        sheets
            .Elements(Main + "sheet")
            .Where(s => worksheets.Deleted.Contains(s.Attribute(Rel + "id")?.Value ?? string.Empty))
            .ToList()
            .ForEach(s => s.Remove());

        // A sheet that survived may have been renamed, and the model is the authority on that.
        Dictionary<string, XElement> byRelId = new(StringComparer.Ordinal);
        foreach (XElement sheet in sheets.Elements(Main + "sheet"))
        {
            if (sheet.Attribute(Rel + "id")?.Value is { } relId)
            {
                byRelId[relId] = sheet;
            }
        }

        List<XElement> supported = [];
        foreach (XLWorksheet xlSheet in worksheets.OrderBy<XLWorksheet, int>(w => w.Position))
        {
            // AssignSheetRelIds ran earlier in the save, so every sheet has one by now.
            string relId = xlSheet.RelId ?? string.Empty;
            if (!byRelId.TryGetValue(relId, out XElement? sheet))
            {
                sheet = new XElement(Main + "sheet");
                sheet.SetAttributeValue(Rel + "id", relId);
            }

            sheet.SetAttributeValue("name", xlSheet.Name);
            sheet.SetAttributeValue("sheetId", xlSheet.SheetId.ToInvariantString());
            sheet.SetAttributeValue(
                "state",
                xlSheet.Visibility == XLWorksheetVisibility.Visible
                    ? null
                    : SheetState(xlSheet.Visibility)
            );

            supported.Add(sheet);
        }

        // The unsupported sheets keep the elements they were loaded with; they are only put back
        // in the right place among the supported ones.
        Dictionary<uint, XElement> unsupportedById = new();
        foreach (XElement sheet in sheets.Elements(Main + "sheet"))
        {
            if (
                uint.TryParse(sheet.Attribute("sheetId")?.Value, out uint sheetId)
                && xlWorkbook.UnsupportedSheets.Any(us => us.SheetId == sheetId)
            )
            {
                unsupportedById[sheetId] = sheet;
            }
        }

        int totalSheets = supported.Count + xlWorkbook.UnsupportedSheets.Count;
        List<XElement> ordered = [];
        uint firstSheetVisible = 0;
        bool foundVisible = false;
        int nextSupported = 0;

        for (int p = 1; p <= totalSheets; p++)
        {
            XLWorkbook.UnsupportedSheet? unsupported = xlWorkbook.UnsupportedSheets.FirstOrDefault(
                us => us.Position == p
            );

            if (unsupported is not null)
            {
                if (unsupportedById.TryGetValue(unsupported.SheetId, out XElement? sheet))
                {
                    ordered.Add(sheet);
                }

                continue;
            }

            if (nextSupported >= supported.Count)
            {
                continue;
            }

            XElement supportedSheet = supported[nextSupported++];
            ordered.Add(supportedSheet);

            if (foundVisible)
            {
                continue;
            }

            if (supportedSheet.Attribute("state")?.Value is null or "visible")
            {
                foundVisible = true;
            }
            else
            {
                firstSheetVisible++;
            }
        }

        sheets.RemoveNodes();
        sheets.Add(ordered);

        WriteBookViews(workbook, xlWorkbook, firstSheetVisible);
    }

    private static void WriteBookViews(
        XElement workbook,
        XLWorkbook xlWorkbook,
        uint firstSheetVisible
    )
    {
        uint activeTab = (
            from us in xlWorkbook.UnsupportedSheets
            where us.IsActive
            select (uint)us.Position - 1
        ).FirstOrDefault();

        if (activeTab == 0)
        {
            uint? firstActiveTab = null;
            uint? firstSelectedTab = null;
            foreach (XLWorksheet ws in xlWorkbook.WorksheetsInternal)
            {
                if (ws.TabActive)
                {
                    firstActiveTab = (uint)(ws.Position - 1);
                    break;
                }

                if (ws.TabSelected)
                {
                    firstSelectedTab = (uint)(ws.Position - 1);
                }
            }

            activeTab = firstActiveTab ?? firstSelectedTab ?? firstSheetVisible;
        }

        XElement bookViews = Ensure(workbook, "bookViews");
        XElement? workbookView = bookViews.Element(Main + "workbookView");
        if (workbookView is null)
        {
            workbookView = new XElement(Main + "workbookView");
            bookViews.Add(workbookView);
        }

        workbookView.SetAttributeValue("firstSheet", firstSheetVisible.ToInvariantString());
        workbookView.SetAttributeValue("activeTab", activeTab.ToInvariantString());
    }

    private static XElement BuildDefinedNames(XElement workbook, XLWorkbook xlWorkbook)
    {
        XElement definedNames = new(Main + "definedNames");

        // The local sheet id of a defined name is the sheet's index in the sheet list, not its
        // sheetId, so it has to be read off the list that was just built.
        List<uint> sheetIds =
        [
            .. workbook
                .Element(Main + "sheets")
                ?.Elements(Main + "sheet")
                .Select(s => uint.TryParse(s.Attribute("sheetId")?.Value, out uint id) ? id : 0u)
                ?? [],
        ];

        foreach (XLWorksheet worksheet in xlWorkbook.WorksheetsInternal)
        {
            uint localSheetId = (uint)Math.Max(0, sheetIds.IndexOf(worksheet.SheetId));

            if (worksheet.PageSetup.PrintAreas.Any())
            {
                string printAreas = string.Join(
                    ",",
                    worksheet.PageSetup.PrintAreas.Select(printArea =>
                        worksheet.Name.EscapeSheetName()
                        + "!"
                        + printArea.RangeAddress.FirstAddress.ToStringFixed(XLReferenceStyle.A1)
                        + ":"
                        + printArea.RangeAddress.LastAddress.ToStringFixed(XLReferenceStyle.A1)
                    )
                );

                definedNames.Add(
                    DefinedName("_xlnm.Print_Area", printAreas, localSheetId, hidden: false, null)
                );
            }

            if (worksheet.AutoFilter.IsEnabled)
            {
                string range =
                    worksheet.Name.EscapeSheetName()
                    + "!"
                    + worksheet.AutoFilter.Range.RangeAddress.FirstAddress.ToStringFixed(
                        XLReferenceStyle.A1
                    )
                    + ":"
                    + worksheet.AutoFilter.Range.RangeAddress.LastAddress.ToStringFixed(
                        XLReferenceStyle.A1
                    );

                definedNames.Add(
                    DefinedName("_xlnm._FilterDatabase", range, localSheetId, hidden: true, null)
                );
            }

            foreach (
                XLDefinedName xlDefinedName in worksheet.DefinedNames.Where<XLDefinedName>(n =>
                    n.Name != "_xlnm._FilterDatabase"
                )
            )
            {
                definedNames.Add(
                    DefinedName(
                        xlDefinedName.Name,
                        xlDefinedName.ToString(),
                        localSheetId,
                        !xlDefinedName.Visible,
                        xlDefinedName.Comment
                    )
                );
            }

            if (PrintTitles(worksheet) is { } titles)
            {
                definedNames.Add(
                    DefinedName("_xlnm.Print_Titles", titles, localSheetId, hidden: false, null)
                );
            }
        }

        foreach (XLDefinedName xlDefinedName in xlWorkbook.DefinedNamesInternal)
        {
            definedNames.Add(
                DefinedName(
                    xlDefinedName.Name,
                    xlDefinedName.RefersTo,
                    localSheetId: null,
                    !xlDefinedName.Visible,
                    xlDefinedName.Comment
                )
            );
        }

        return definedNames;
    }

    /// <summary>The rows and columns repeated on every printed page, or nothing when there are none.</summary>
    private static string? PrintTitles(XLWorksheet worksheet)
    {
        string rows =
            worksheet.PageSetup.FirstRowToRepeatAtTop > 0
                ? worksheet.Name.EscapeSheetName()
                    + "!"
                    + worksheet.PageSetup.FirstRowToRepeatAtTop
                    + ":"
                    + worksheet.PageSetup.LastRowToRepeatAtTop
                : string.Empty;

        string columns =
            worksheet.PageSetup.FirstColumnToRepeatAtLeft > 0
                ? worksheet.Name.EscapeSheetName()
                    + "!"
                    + XlsxSharp.XLHelper.GetColumnLetterFromNumber(
                        worksheet.PageSetup.FirstColumnToRepeatAtLeft
                    )
                    + ":"
                    + XlsxSharp.XLHelper.GetColumnLetterFromNumber(
                        worksheet.PageSetup.LastColumnToRepeatAtLeft
                    )
                : string.Empty;

        string titles =
            columns.Length > 0
                ? rows.Length > 0
                    ? columns + "," + rows
                    : columns
                : rows;

        return titles.Length > 0 ? titles : null;
    }

    private static XElement DefinedName(
        string name,
        string text,
        uint? localSheetId,
        bool hidden,
        string? comment
    )
    {
        XElement definedName = new(Main + "definedName", text);
        definedName.SetAttributeValue("name", name);
        if (comment is not null && !string.IsNullOrWhiteSpace(comment))
        {
            definedName.SetAttributeValue("comment", comment);
        }

        if (localSheetId is not null)
        {
            definedName.SetAttributeValue(
                "localSheetId",
                localSheetId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
            );
        }

        if (hidden)
        {
            definedName.SetAttributeValue("hidden", "1");
        }

        return definedName;
    }

    private static void WriteCalculationProperties(XElement workbook, XLWorkbook xlWorkbook)
    {
        XElement calcPr = Ensure(workbook, "calcPr");

        if (calcPr.Attribute("calcId") is null)
        {
            calcPr.SetAttributeValue("calcId", "125725");
        }

        calcPr.SetAttributeValue(
            "calcMode",
            xlWorkbook.CalculateMode == XLCalculateMode.Default
                ? null
                : CalculateMode(xlWorkbook.CalculateMode)
        );

        calcPr.SetAttributeValue(
            "refMode",
            xlWorkbook.ReferenceStyle == XLReferenceStyle.Default
                ? null
                : ReferenceMode(xlWorkbook.ReferenceStyle)
        );

        // These four are only written when set; the previous writer left whatever was loaded
        // otherwise, and so does this one.
        SetIfTrue(calcPr, "calcOnSave", xlWorkbook.CalculationOnSave);
        SetIfTrue(calcPr, "forceFullCalc", xlWorkbook.ForceFullCalculation);
        SetIfTrue(calcPr, "fullCalcOnLoad", xlWorkbook.FullCalculationOnLoad);
        SetIfTrue(calcPr, "fullPrecision", xlWorkbook.FullPrecision);
    }

    private static void WritePivotCaches(XElement workbook, XLWorkbook xlWorkbook)
    {
        List<XLPivotCache> used =
        [
            .. xlWorkbook
                .WorksheetsInternal.SelectMany<XLWorksheet, XLPivotTable>(ws => ws.PivotTables)
                .Select(pt => pt.PivotCache)
                .Distinct()
                .Cast<XLPivotCache>(),
        ];

        if (used.Count == 0)
        {
            workbook.Element(Main + "pivotCaches")?.Remove();
            return;
        }

        // Rebuilt rather than patched, to drop references a previous save left behind.
        XElement pivotCaches = new(Main + "pivotCaches");
        foreach (XLPivotCache source in used)
        {
            XElement pivotCache = new(Main + "pivotCache");
            pivotCache.SetAttributeValue("cacheId", (source.CacheId ?? 0u).ToInvariantString());
            pivotCache.SetAttributeValue(Rel + "id", source.WorkbookCacheRelId);
            pivotCaches.Add(pivotCache);
        }

        SetElement(workbook, "pivotCaches", pivotCaches);
    }

    private static string SheetState(XLWorksheetVisibility visibility) =>
        visibility switch
        {
            XLWorksheetVisibility.Visible => "visible",
            XLWorksheetVisibility.Hidden => "hidden",
            XLWorksheetVisibility.VeryHidden => "veryHidden",
            _ => throw new ArgumentOutOfRangeException(nameof(visibility)),
        };

    private static string CalculateMode(XLCalculateMode mode) =>
        mode switch
        {
            XLCalculateMode.Auto => "auto",
            XLCalculateMode.AutoNoTable => "autoNoTable",
            XLCalculateMode.Manual => "manual",
            _ => throw new ArgumentOutOfRangeException(nameof(mode)),
        };

    private static string ReferenceMode(XLReferenceStyle style) =>
        style switch
        {
            XLReferenceStyle.A1 => "A1",
            XLReferenceStyle.R1C1 => "R1C1",
            _ => throw new ArgumentOutOfRangeException(nameof(style)),
        };

    /// <summary>The child of that name, added at its place in the schema order when missing.</summary>
    private static XElement Ensure(XElement workbook, string localName)
    {
        XElement? existing = workbook.Element(Main + localName);
        if (existing is not null)
        {
            return existing;
        }

        XElement created = new(Main + localName);
        SetElement(workbook, localName, created);
        return created;
    }

    private static void SetElement(XElement workbook, string localName, XElement replacement)
    {
        XElement? existing = workbook.Element(Main + localName);
        if (existing is not null)
        {
            existing.ReplaceWith(replacement);
            return;
        }

        XElement? predecessor = null;
        int position = Array.IndexOf(ElementOrder, localName);
        for (int i = 0; i < position; i++)
        {
            if (workbook.Element(Main + ElementOrder[i]) is { } candidate)
            {
                predecessor = candidate;
            }
        }

        if (predecessor is null)
        {
            workbook.AddFirst(replacement);
        }
        else
        {
            predecessor.AddAfterSelf(replacement);
        }
    }

    /// <summary>Writes the attribute only when it differs from what a reader would assume.</summary>
    private static void SetOptionalBool(
        XElement element,
        string name,
        bool value,
        bool defaultValue
    ) =>
        element.SetAttributeValue(
            name,
            value == defaultValue ? null
                : value ? "1"
                : "0"
        );

    private static void SetIfTrue(XElement element, string name, bool value)
    {
        if (value)
        {
            element.SetAttributeValue(name, "1");
        }
    }
}

#nullable disable

using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using XlsxSharp.Excel.IO;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using Path = System.IO.Path;

namespace XlsxSharp.Excel;

public partial class XLWorkbook
{
    private static void Validate(SpreadsheetDocument package)
    {
        CultureInfo backupCulture = Thread.CurrentThread.CurrentCulture;

        IList<ValidationErrorInfo> errors;
        try
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            OpenXmlValidator validator = new();
            errors = validator.Validate(package).ToArray();

            // The styles part is written directly into the part stream, without intermediate OOXML SDK
            // representation. The validation loads the written XML into a memory so it can validate it.
            // But the loaded in-memory represenation from validation is then picked up by
            // the SpreadsheetDocument saving code and the XML would be re-serialized and the original XML
            // from part writer would be discarded. That is a problem for the following reasons:
            // * XML of a part would be different when saved with a validation and without a validation
            // * There is no control over XML serialization. The writer is attempting to mimic Excel for
            //   easier comparison against Excel (e.g. default namespace doesn't have a prefix), but part
            //   writer just uses default serialization setting.
            // To solve this, we discard the in-memory represenation created by the validator. Thus the save
            // code will not re-serialize the part that has already been written.
            if (package.WorkbookPart.WorkbookStylesPart is { } stylesPart)
            {
                stylesPart.UnloadRootElement();
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = backupCulture;
        }

        if (errors.Any())
        {
            string message = string.Join(
                "\r\n",
                errors
                    .Select(e =>
                        string.Format(
                            "Part {0}, Path {1}: {2}",
                            e.Part.Uri,
                            e.Path.XPath,
                            e.Description
                        )
                    )
                    .ToArray()
            );
            throw new ApplicationException(message);
        }
    }

    /// <summary>
    /// The one place the workbook's own document type meets the SDK's, so that nothing else in
    /// the model has to know the SDK has an enum for this.
    /// </summary>
    private static SpreadsheetDocumentType ToOpenXml(XLSpreadsheetDocumentType documentType) =>
        documentType switch
        {
            XLSpreadsheetDocumentType.Workbook => SpreadsheetDocumentType.Workbook,
            XLSpreadsheetDocumentType.Template => SpreadsheetDocumentType.Template,
            XLSpreadsheetDocumentType.MacroEnabledWorkbook =>
                SpreadsheetDocumentType.MacroEnabledWorkbook,
            XLSpreadsheetDocumentType.MacroEnabledTemplate =>
                SpreadsheetDocumentType.MacroEnabledTemplate,
            _ => throw new ArgumentOutOfRangeException(nameof(documentType)),
        };

    private void CreatePackage(
        string filePath,
        XLSpreadsheetDocumentType spreadsheetDocumentType,
        SaveOptions options
    )
    {
        string directoryName = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(directoryName))
        {
            Directory.CreateDirectory(directoryName);
        }

        SpreadsheetDocumentType documentType = ToOpenXml(spreadsheetDocumentType);
        SpreadsheetDocument package = File.Exists(filePath)
            ? SpreadsheetDocument.Open(filePath, true)
            : SpreadsheetDocument.Create(filePath, documentType);

        using (package)
        {
            if (package.DocumentType != documentType)
            {
                package.ChangeDocumentType(documentType);
            }

            this.CreateParts(package, options);
            if (options.ValidatePackage)
            {
                Validate(package);
            }
        }
    }

    private void CreatePackage(
        Stream stream,
        bool newStream,
        XLSpreadsheetDocumentType spreadsheetDocumentType,
        SaveOptions options
    )
    {
        SpreadsheetDocument package = newStream
            ? SpreadsheetDocument.Create(stream, ToOpenXml(spreadsheetDocumentType))
            : SpreadsheetDocument.Open(stream, true);

        using (package)
        {
            this.CreateParts(package, options);
            if (options.ValidatePackage)
            {
                Validate(package);
            }
        }
    }

    // http://blogs.msdn.com/b/vsod/archive/2010/02/05/how-to-delete-a-worksheet-from-excel-using-open-xml-sdk-2-0.aspx
    private static void DeleteSheetAndDependencies(WorkbookPart wbPart, string sheetId)
    {
        //Get the SheetToDelete from workbook.xml
        Sheet worksheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id == sheetId);
        if (worksheet == null)
        {
            return;
        }

        string sheetName = worksheet.Name;
        // Get the pivot Table Parts
        IEnumerable<PivotTableCacheDefinitionPart> pvtTableCacheParts =
            wbPart.PivotTableCacheDefinitionParts;
        Dictionary<PivotTableCacheDefinitionPart, string> pvtTableCacheDefinitionPart = new();
        foreach (PivotTableCacheDefinitionPart Item in pvtTableCacheParts)
        {
            PivotCacheDefinition pvtCacheDef = Item.PivotCacheDefinition;
            //Check if this CacheSource is linked to SheetToDelete
            if (
                pvtCacheDef
                    .Descendants<CacheSource>()
                    .Any(cacheSource => cacheSource.WorksheetSource?.Sheet == sheetName)
            )
            {
                pvtTableCacheDefinitionPart.Add(Item, Item.ToString());
            }
        }
        foreach (
            KeyValuePair<PivotTableCacheDefinitionPart, string> Item in pvtTableCacheDefinitionPart
        )
        {
            wbPart.DeletePart(Item.Key);
        }

        // Remove the sheet reference from the workbook.
        WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(sheetId));
        worksheet.Remove();

        // Delete the worksheet part.
        wbPart.DeletePart(worksheetPart);

        //Get the DefinedNames
        DefinedNames definedNames = wbPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
        if (definedNames != null)
        {
            List<DefinedName> defNamesToDelete = [];

            foreach (DefinedName Item in definedNames.OfType<DefinedName>())
            {
                // This condition checks to delete only those names which are part of Sheet in question
                if (Item.Text.Contains(worksheet.Name + "!"))
                {
                    defNamesToDelete.Add(Item);
                }
            }

            foreach (DefinedName Item in defNamesToDelete)
            {
                Item.Remove();
            }
        }
        // Get the CalculationChainPart
        //Note: An instance of this part type contains an ordered set of references to all cells in all worksheets in the
        //workbook whose value is calculated from any formula

        CalculationChainPart calChainPart;
        calChainPart = wbPart.CalculationChainPart;
        if (calChainPart != null)
        {
            IEnumerable<CalculationCell> calChainEntries = calChainPart
                .CalculationChain.Descendants<CalculationCell>()
                .Where(c => c.SheetId == sheetId);
            List<CalculationCell> calcsToDelete = [];
            foreach (CalculationCell Item in calChainEntries)
            {
                calcsToDelete.Add(Item);
            }

            foreach (CalculationCell Item in calcsToDelete)
            {
                Item.Remove();
            }

            if (!calChainPart.CalculationChain.Any())
            {
                wbPart.DeletePart(calChainPart);
            }
        }
    }

    // Adds child parts and generates content of the specified part.
    private void CreateParts(SpreadsheetDocument document, SaveOptions options)
    {
        SaveContext context = new();

        WorkbookPart workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();

        XLWorksheets worksheets = this.WorksheetsInternal;

        List<IdPartPair> partsToRemove =
        [
            .. workbookPart.Parts.Where(s => worksheets.Deleted.Contains(s.RelationshipId)),
        ];

        List<PivotTableCacheDefinitionPart> pivotCacheDefinitionsToRemove =
        [
            .. partsToRemove
                .SelectMany(s =>
                    ((WorksheetPart)s.OpenXmlPart).PivotTableParts.Select(pt =>
                        pt.PivotTableCacheDefinitionPart
                    )
                )
                .Distinct(),
        ];
        pivotCacheDefinitionsToRemove.ForEach(c => workbookPart.DeletePart(c));

        if (workbookPart.Workbook != null && workbookPart.Workbook.PivotCaches != null)
        {
            List<OpenXmlElement> pivotCachesToRemove =
            [
                .. workbookPart
                    .Workbook.PivotCaches.Where(pc =>
                        pivotCacheDefinitionsToRemove
                            .Select(pcd => workbookPart.GetIdOfPart(pcd))
                            .ToList()
                            .Contains(((PivotCache)pc).Id)
                    )
                    .Distinct(),
            ];
            pivotCachesToRemove.ForEach(c => workbookPart.Workbook.PivotCaches.RemoveChild(c));
        }

        worksheets.Deleted.ToList().ForEach(ws => DeleteSheetAndDependencies(workbookPart, ws));

        // Ensure all RelId's have been added to the context
        context.RelIdGenerator.AddExistingValues(workbookPart, this);

        ExtendedFilePropertiesPart extendedFilePropertiesPart =
            document.ExtendedFilePropertiesPart
            ?? document.AddNewPart<ExtendedFilePropertiesPart>(
                context.RelIdGenerator.GetNext(RelType.Workbook)
            );

        ExtendedFilePropertiesPartWriter.GenerateContent(extendedFilePropertiesPart, this);

        WorkbookPartWriter.GenerateContent(workbookPart, this, options, context);

        WorkbookStylesPart workbookStylesPart =
            workbookPart.WorkbookStylesPart
            ?? workbookPart.AddNewPart<WorkbookStylesPart>(
                context.RelIdGenerator.GetNext(RelType.Workbook)
            );

        new StylesWriter().WriteContent(
            workbookStylesPart,
            XmlToEnumMapper.Instance,
            this.Styles,
            this,
            context
        );

        SharedStringTablePart sharedStringTablePart =
            workbookPart.SharedStringTablePart
            ?? workbookPart.AddNewPart<SharedStringTablePart>(
                context.RelIdGenerator.GetNext(RelType.Workbook)
            );

        SharedStringTableWriter.GenerateSharedStringTablePartContent(
            this,
            sharedStringTablePart,
            context
        );

        List<IXLPivotTable> allPivotTables =
        [
            .. this.WorksheetsInternal.SelectMany<XLWorksheet, IXLPivotTable>(ws => ws.PivotTables),
        ];

        // Phase 1 - Synchronize all pivot cache parts in the document, so each
        // source that will be saved has all required parts created and relationship
        // ids are set (in this case `Workbook.PivotCaches` relationship table).
        // Only sources that are used by a table are saved.
        SynchronizePivotTableParts(workbookPart, allPivotTables, context);

        // Phase 2 - All parts and relationships are set, fill in the parts.
        if (allPivotTables.Any())
        {
            this.GeneratePivotCaches(workbookPart, context);
        }

        foreach (
            XLWorksheet worksheet in this
                .WorksheetsInternal.Cast<XLWorksheet>()
                .OrderBy(w => w.Position)
        )
        {
            WorksheetPart worksheetPart;
            string wsRelId = worksheet.RelId;
            bool partIsEmpty;
            if (workbookPart.Parts.Any(p => p.RelationshipId == wsRelId))
            {
                worksheetPart = (WorksheetPart)workbookPart.GetPartById(wsRelId);
                partIsEmpty = false;
            }
            else
            {
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>(wsRelId);
                partIsEmpty = true;
            }

            bool worksheetHasComments = worksheet
                .Internals.CellsCollection.GetCells(c => c.HasComment)
                .Any();

            WorksheetCommentsPart commentsPart = worksheetPart.WorksheetCommentsPart;

            // VML part is the source of truth for shapes of comments, form controls and likely others.
            // Excel won't display any shape without VML. The drawing part is always present, but is likely
            // only different rendering of VML (more precisely the shapes behind VML).
            VmlDrawingPart vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
            bool hasAnyVmlElements = DeleteExistingCommentsShapes(vmlDrawingPart);

            if (worksheetHasComments)
            {
                // If sheet has comments, we must keep VML in legacy drawing part to display them
                // as well as comments part for semantic reasons.
                if (commentsPart == null)
                {
                    commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>(
                        context.RelIdGenerator.GetNext(RelType.Workbook)
                    );
                }

                if (vmlDrawingPart == null)
                {
                    if (string.IsNullOrWhiteSpace(worksheet.LegacyDrawingId))
                    {
                        worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(
                            RelType.Workbook
                        );
                    }

                    vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(
                        worksheet.LegacyDrawingId
                    );
                }

                CommentPartWriter.GenerateWorksheetCommentsPartContent(commentsPart, worksheet);
                hasAnyVmlElements = VmlDrawingPartWriter.GenerateContent(vmlDrawingPart, worksheet);
            }
            else
            {
                // There are no comments in the worksheet = the comment part is no longer needed,
                // but VML part might contain other shapes, like form controls.
                if (commentsPart is not null)
                {
                    worksheetPart.DeletePart(commentsPart);
                }
            }

            if (!hasAnyVmlElements && vmlDrawingPart is not null)
            {
                worksheet.LegacyDrawingId = null;
                worksheetPart.DeletePart(vmlDrawingPart);
            }

            XLTables xlTables = worksheet.Tables;

            // The way forward is to have 2-phase save, this is a start of that
            // concept for tables:
            //
            // Phase 1 - synchronize part existence with tables xlWorksheet, so each
            // table has a corresponding part and part that don't are deleted.
            // This phase doesn't modify the content, it only ensures that RelIds are set
            // corresponding parts exist and the parts that don't exist are removed
            TablePartWriter.SynchronizeTableParts(xlTables, worksheetPart, context);

            // Phase 2 - At this point, all pieces must have corresponding parts
            // The only way to link between parts is through RelIds that were already
            // set in phase 1. The phase 2 is all about content of individual parts.
            // Each part should have individual writer.
            TablePartWriter.GenerateTableParts(xlTables, worksheetPart, context);

            WorksheetPartWriter.GenerateWorksheetPartContent(
                partIsEmpty,
                worksheetPart,
                worksheet,
                options,
                context
            );

            if (worksheet.PivotTables.Any<XLPivotTable>())
            {
                GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
            }
        }

        if (options.GenerateCalculationChain)
        {
            CalculationChainPartWriter.GenerateContent(workbookPart, this, context);
        }
        else
        {
            if (workbookPart.CalculationChainPart is not null)
            {
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
            }
        }

        if (workbookPart.ThemePart == null)
        {
            ThemePart themePart = workbookPart.AddNewPart<ThemePart>(
                context.RelIdGenerator.GetNext(RelType.Workbook)
            );
            ThemePartWriter.GenerateContent(themePart, (XLTheme)this.Theme);
        }

        // Custom properties
        if (this.CustomProperties.Any())
        {
            CustomFilePropertiesPart customFilePropertiesPart =
                document.CustomFilePropertiesPart
                ?? document.AddNewPart<CustomFilePropertiesPart>(
                    context.RelIdGenerator.GetNext(RelType.Workbook)
                );

            CustomFilePropertiesPartWriter.GenerateContent(customFilePropertiesPart, this);
        }
        else
        {
            if (document.CustomFilePropertiesPart != null)
            {
                document.DeletePart(document.CustomFilePropertiesPart);
            }
        }
        this.SetPackageProperties(document);

        // Clear list of deleted worksheets to prevent errors on multiple saves
        worksheets.Deleted.Clear();
    }

    private static bool DeleteExistingCommentsShapes(VmlDrawingPart vmlDrawingPart)
    {
        if (vmlDrawingPart == null)
        {
            return false;
        }

        // Nuke the VmlDrawingPart elements for comments.
        using (Stream vmlStream = vmlDrawingPart.GetStream(FileMode.Open))
        {
            XDocument xdoc = XDocumentExtensions.Load(vmlStream);
            if (xdoc == null)
            {
                return false;
            }

            // Remove existing shapes for comments
            xdoc.Root.Elements()
                .Where(e =>
                    e.Name.LocalName == "shapetype"
                    && e.Attribute("id").Value == XLConstants.Comment.ShapeTypeId
                )
                .Remove();

            xdoc.Root.Elements()
                .Where(e =>
                    e.Name.LocalName == "shape"
                    && e.Attribute("type").Value == "#" + XLConstants.Comment.ShapeTypeId
                )
                .Remove();

            vmlStream.Position = 0;

            using (XmlTextWriter writer = new(vmlStream, Encoding.UTF8))
            {
                string contents = xdoc.ToXmlString();
                writer.WriteRaw(contents);
                vmlStream.SetLength(contents.Length);
            }

            return xdoc.Root.HasElements;
        }
    }

    private void SetPackageProperties(OpenXmlPackage document)
    {
        DateTime created =
            this.Properties.Created == DateTime.MinValue ? DateTime.Now : this.Properties.Created;
        DateTime modified =
            this.Properties.Modified == DateTime.MinValue ? DateTime.Now : this.Properties.Modified;
        document.PackageProperties.Created = created;
        document.PackageProperties.Modified = modified;

#if true // Workaround: https://github.com/OfficeDev/Open-XML-SDK/issues/235

        if (this.Properties.LastModifiedBy == null)
        {
            document.PackageProperties.LastModifiedBy = "";
        }

        if (this.Properties.Author == null)
        {
            document.PackageProperties.Creator = "";
        }

        if (this.Properties.Title == null)
        {
            document.PackageProperties.Title = "";
        }

        if (this.Properties.Subject == null)
        {
            document.PackageProperties.Subject = "";
        }

        if (this.Properties.Category == null)
        {
            document.PackageProperties.Category = "";
        }

        if (this.Properties.Keywords == null)
        {
            document.PackageProperties.Keywords = "";
        }

        if (this.Properties.Comments == null)
        {
            document.PackageProperties.Description = "";
        }

        if (this.Properties.Status == null)
        {
            document.PackageProperties.ContentStatus = "";
        }

#endif

        document.PackageProperties.LastModifiedBy = this.Properties.LastModifiedBy;

        document.PackageProperties.Creator = this.Properties.Author;
        document.PackageProperties.Title = this.Properties.Title;
        document.PackageProperties.Subject = this.Properties.Subject;
        document.PackageProperties.Category = this.Properties.Category;
        document.PackageProperties.Keywords = this.Properties.Keywords;
        document.PackageProperties.Description = this.Properties.Comments;
        document.PackageProperties.ContentStatus = this.Properties.Status;
    }

    private static void SynchronizePivotTableParts(
        WorkbookPart workbookPart,
        IReadOnlyList<IXLPivotTable> allPivotTables,
        SaveContext context
    )
    {
        RemoveUnusedPivotCacheDefinitionParts(workbookPart, allPivotTables);
        AddUsedPivotCacheDefinitionParts(workbookPart, allPivotTables, context);

        // Ensure this in workbook.xml:
        //  <pivotCaches>
        //    <pivotCache cacheId="13" r:id="rId3"/>
        //  </pivotCaches>

        context.PivotSourceCacheId = 0;
        List<XLPivotCache> xlUsedCaches =
        [
            .. allPivotTables.Select(pt => pt.PivotCache).Distinct().Cast<XLPivotCache>(),
        ];
        if (xlUsedCaches.Any())
        {
            // Recreate the workbook pivot cache references to remove previous gunk
            PivotCaches pivotCaches = new();
            workbookPart.Workbook.PivotCaches = pivotCaches;

            foreach (XLPivotCache source in xlUsedCaches)
            {
                uint cacheId = context.PivotSourceCacheId++;
                source.CacheId = cacheId;
                PivotCache pivotCache = new() { CacheId = cacheId, Id = source.WorkbookCacheRelId };
                pivotCaches.AppendChild(pivotCache);
            }
        }
        else
        {
            // Remove empty pivot cache part
            if (workbookPart.Workbook.PivotCaches is not null)
            {
                workbookPart.Workbook.RemoveChild(workbookPart.Workbook.PivotCaches);
            }
        }

        // Remove pivot cache parts that are a part of the loaded document, but aren't used by a pivot table of the xlWorkbook
        // part of the first phase of saving
        static void RemoveUnusedPivotCacheDefinitionParts(
            WorkbookPart workbookPart,
            IReadOnlyList<IXLPivotTable> allPivotTables
        )
        {
            List<string> workbookCacheRelIds =
            [
                .. allPivotTables
                    .Select(pt => pt.PivotCache.CastTo<XLPivotCache>().WorkbookCacheRelId)
                    .Distinct(),
            ];

            List<PivotTableCacheDefinitionPart> orphanedParts =
            [
                .. workbookPart
                    .GetPartsOfType<PivotTableCacheDefinitionPart>()
                    .Where(pcdp => !workbookCacheRelIds.Contains(workbookPart.GetIdOfPart(pcdp))),
            ];

            foreach (PivotTableCacheDefinitionPart orphanPart in orphanedParts)
            {
                orphanPart.DeletePart(orphanPart.PivotTableCacheRecordsPart);
                workbookPart.DeletePart(orphanPart);
            }

            // Remove deleted pivot cache parts
            if (workbookPart.Workbook.PivotCaches is not null)
            {
                workbookPart
                    .Workbook.PivotCaches.Elements<PivotCache>()
                    .Where(pc => !workbookPart.HasPartWithId(pc.Id))
                    .ToList()
                    .ForEach(pc => pc.Remove());
            }
        }

        static void AddUsedPivotCacheDefinitionParts(
            WorkbookPart workbookPart,
            IReadOnlyList<IXLPivotTable> allPivotTables,
            SaveContext context
        )
        {
            // Add ids and part for the caches to workbooks
            // We might get a XLPivotSource with an id of apart that isn't in the file (e.g. loaded from a file and saved to a different one).
            List<XLPivotCache> newPivotSources =
            [
                .. allPivotTables
                    .Select(pt => pt.PivotCache.CastTo<XLPivotCache>())
                    .Where(ps =>
                        string.IsNullOrEmpty(ps.WorkbookCacheRelId)
                        || !workbookPart.HasPartWithId(ps.WorkbookCacheRelId)
                    )
                    .Distinct(),
            ];

            foreach (XLPivotCache pivotSource in newPivotSources)
            {
                string cacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                pivotSource.WorkbookCacheRelId = cacheRelId;

                workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(
                    pivotSource.WorkbookCacheRelId
                );
            }
        }
    }

    private void GeneratePivotCaches(WorkbookPart workbookPart, SaveContext context)
    {
        IEnumerable<XLPivotTable> pivotTables = this.WorksheetsInternal.SelectMany<
            XLWorksheet,
            XLPivotTable
        >(ws => ws.PivotTables);

        IEnumerable<XLPivotCache> xlPivotCaches = pivotTables
            .Select(pt => pt.PivotCache)
            .Distinct();
        foreach (XLPivotCache xlPivotCache in xlPivotCaches)
        {
            Debug.Assert(workbookPart.Workbook.PivotCaches is not null);
            Debug.Assert(!string.IsNullOrEmpty(xlPivotCache.WorkbookCacheRelId));

            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart =
                (PivotTableCacheDefinitionPart)
                    workbookPart.GetPartById(xlPivotCache.WorkbookCacheRelId);

            PivotTableCacheDefinitionPartWriter.GenerateContent(
                pivotTableCacheDefinitionPart,
                xlPivotCache,
                context
            );

            PivotTableCacheRecordsPart pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart
                .GetPartsOfType<PivotTableCacheRecordsPart>()
                .Any()
                ? pivotTableCacheDefinitionPart
                    .GetPartsOfType<PivotTableCacheRecordsPart>()
                    .Single()
                : pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

            PivotCacheRecordsWriter.WriteContent(pivotTableCacheRecordsPart, xlPivotCache);
        }
    }

    private static void GeneratePivotTables(
        WorkbookPart workbookPart,
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context
    )
    {
        foreach (XLPivotTable pt in xlWorksheet.PivotTables)
        {
            PivotTablePart pivotTablePart;
            bool createNewPivotTablePart = string.IsNullOrWhiteSpace(pt.RelId);
            if (createNewPivotTablePart)
            {
                string relId = context.RelIdGenerator.GetNext(RelType.Workbook);
                pt.RelId = relId;
                pivotTablePart = worksheetPart.AddNewPart<PivotTablePart>(relId);
            }
            else
            {
                pivotTablePart = (PivotTablePart)worksheetPart.GetPartById(pt.RelId);
            }

            XLPivotCache pivotSource = pt.PivotCache;
            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart =
                pivotTablePart.PivotTableCacheDefinitionPart;
            if (
                !workbookPart
                    .GetPartById(pivotSource.WorkbookCacheRelId)
                    .Equals(pivotTableCacheDefinitionPart)
            )
            {
                pivotTablePart.DeletePart(pivotTableCacheDefinitionPart);
                pivotTablePart.CreateRelationshipToPart(
                    workbookPart.GetPartById(pivotSource.WorkbookCacheRelId),
                    context.RelIdGenerator.GetNext(XLWorkbook.RelType.Workbook)
                );
            }

            PivotTableDefinitionPartWriter2.WriteContent(pivotTablePart, pt, context);
        }
    }
}

#nullable disable

using System.Diagnostics;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Excel.IO;
using XlsxSharp.Excel.IO.Schemas;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.IO.Packaging;
using Path = System.IO.Path;

namespace XlsxSharp.Excel;

public partial class XLWorkbook
{
    /// <summary>
    /// Checks every schema-mapped part of the saved package against the OOXML schemas, see
    /// <see cref="SchemaValidator"/>.
    /// </summary>
    private static void Validate(OpcPackage package)
    {
        IReadOnlyList<string> errors = SchemaValidator.Validate(package);
        if (errors.Count > 0)
        {
            throw new ApplicationException(string.Join("\r\n", errors));
        }
    }

    /// <summary>
    /// The one place the workbook's own document type meets the packaging layer's, so that
    /// nothing else in the model has to know the four kinds only differ in content type.
    /// </summary>
    private static OoxmlPartType WorkbookPartType(XLSpreadsheetDocumentType documentType) =>
        documentType switch
        {
            XLSpreadsheetDocumentType.Workbook => OoxmlPartTypes.Workbook,
            XLSpreadsheetDocumentType.Template => OoxmlPartTypes.WorkbookTemplate,
            XLSpreadsheetDocumentType.MacroEnabledWorkbook => OoxmlPartTypes.MacroEnabledWorkbook,
            XLSpreadsheetDocumentType.MacroEnabledTemplate =>
                OoxmlPartTypes.MacroEnabledWorkbookTemplate,
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

        OpcPackage package = File.Exists(filePath)
            ? OpcPackage.Open(filePath, writable: true)
            : OpcPackage.Create(filePath);

        using (package)
        {
            this.CreateParts(package, WorkbookPartType(spreadsheetDocumentType), options);
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
        if (newStream)
        {
            OoxmlPartType workbookPartType = WorkbookPartType(spreadsheetDocumentType);
            using OpcPackage package = OpcPackage.Create();
            this.CreateParts(package, workbookPartType, options);
            if (options.ValidatePackage)
            {
                Validate(package);
            }

            package.SaveTo(stream);
            return;
        }

        this.CreatePackage(stream, stream, spreadsheetDocumentType, options);
    }

    /// <summary>
    /// Reads the existing package from <paramref name="source"/> and saves the result to the
    /// separate <paramref name="destination"/>, without a caller-side copy from one to the other
    /// first: <see cref="OpcPackage.Open(Stream, Stream)"/> already reads all of
    /// <paramref name="source"/> into its own buffer, so staging the same bytes into
    /// <paramref name="destination"/> beforehand would only copy the whole package a second time
    /// for nothing.
    /// </summary>
    private void CreatePackage(
        Stream source,
        Stream destination,
        XLSpreadsheetDocumentType spreadsheetDocumentType,
        SaveOptions options
    )
    {
        using OpcPackage package = OpcPackage.Open(source, destination);
        this.CreateParts(package, WorkbookPartType(spreadsheetDocumentType), options);
        if (options.ValidatePackage)
        {
            Validate(package);
        }
    }

    /// <summary>
    /// Drops the parts a deleted sheet leaves orphaned: its own worksheet part, and any pivot
    /// cache definition part whose source was that sheet. The sheet element, defined names and
    /// calculation chain entries that referenced it need no equivalent cleanup here - the parts
    /// that carry them (<see cref="WorkbookPartWriter"/>, <see cref="CalculationChainPartWriter"/>)
    /// are rebuilt wholesale from the model on every save rather than patched, so a deleted sheet
    /// is simply absent from what they write.
    /// </summary>
    private static void DeleteSheetAndDependencies(
        OpcPackage package,
        OpcPart wbPart,
        string sheetId
    )
    {
        string? sheetName = WorkbookXml
            .Read(wbPart)
            .Element(SpreadsheetXml.Main + "sheets")
            ?.Elements(SpreadsheetXml.Main + "sheet")
            .FirstOrDefault(s => s.Attribute(SpreadsheetXml.Rel + "id")?.Value == sheetId)
            ?.Attribute("name")
            ?.Value;

        if (sheetName is null)
        {
            return;
        }

        foreach (
            OpcPart cacheDefinitionPart in wbPart
                .PartsOfType(OoxmlPartTypes.PivotCacheDefinition)
                .Where(part => ReadsFromSheet(part, sheetName))
                .ToList()
        )
        {
            package.DeletePart(cacheDefinitionPart.Name);
        }

        OpcPart worksheetPart = wbPart.GetRelatedPart(sheetId);
        package.DeletePart(worksheetPart.Name);
    }

    private static bool ReadsFromSheet(OpcPart part, string sheetName)
    {
        using Stream stream = part.GetReadStream();
        return XDocument
                .Load(stream)
                .Root?.Element(SpreadsheetXml.Main + "cacheSource")
                ?.Element(SpreadsheetXml.Main + "worksheetSource")
                ?.Attribute("sheet")
                ?.Value == sheetName;
    }

    // Adds child parts and generates content of the specified part.
    private void CreateParts(
        OpcPackage package,
        OoxmlPartType workbookPartType,
        SaveOptions options
    )
    {
        SaveContext context = new();

        // Package-level relationships live in /_rels/.rels, an id space of their own, so they take
        // the next free id there rather than one from RelIdGenerator's pool for the workbook.
        OpcPart workbookPart =
            package.PartOfType(OoxmlPartTypes.Workbook)
            ?? package.AddPartOfType(workbookPartType).Part;

        // The workbook, template and macro-enabled variants all point at the same part through
        // the same "officeDocument" relationship, differing only in declared content type - a
        // template loaded and saved as an ordinary workbook (or vice versa) keeps the same part
        // and relationships, but needs that declaration brought in line with what was asked for.
        if (workbookPart.ContentType != workbookPartType.ContentType)
        {
            package.ChangeContentType(workbookPart, workbookPartType.ContentType);
        }

        XLWorksheets worksheets = this.WorksheetsInternal;

        List<OpcPart> partsToRemove =
        [
            .. workbookPart
                .Relationships.Where(r =>
                    r.TargetMode == OpcTargetMode.Internal && worksheets.Deleted.Contains(r.Id)
                )
                .Select(r => package.GetPart(r.TargetPartName!)),
        ];

        // Deleting a worksheet part orphans the pivot cache definition parts owned by pivot
        // tables that lived on it - WorkbookPartWriter rebuilds <pivotCaches> from the surviving
        // pivot tables afterwards, so no reference to these parts needs cleaning up here.
        List<OpcPart> pivotCacheDefinitionsToRemove =
        [
            .. partsToRemove
                .SelectMany(s =>
                    s.PartsOfType(OoxmlPartTypes.PivotTable)
                        .Select(pt => pt.PartOfType(OoxmlPartTypes.PivotCacheDefinition))
                )
                .Where(c => c is not null)
                .Distinct(),
        ];
        pivotCacheDefinitionsToRemove.ForEach(c => package.DeletePart(c!.Name));

        worksheets
            .Deleted.ToList()
            .ForEach(ws => DeleteSheetAndDependencies(package, workbookPart, ws));

        // Ensure all RelId's have been added to the context
        context.RelIdGenerator.AddExistingValues(workbookPart, this);

        OpcPart extendedFilePropertiesPart =
            package.PartOfType(OoxmlPartTypes.ExtendedFileProperties)
            ?? package.AddPartOfType(OoxmlPartTypes.ExtendedFileProperties).Part;

        ExtendedFilePropertiesPartWriter.GenerateContent(extendedFilePropertiesPart, this);

        // Only the relationship ids here. The XML goes out at the end of this method, once the
        // pivot caches exist, but the ids have to be handed out at this point so that the parts
        // created below get the same ones they did before.
        WorkbookPartWriter.AssignSheetRelIds(this, context);

        OpcPart workbookStylesPart =
            workbookPart.PartOfType(OoxmlPartTypes.Styles)
            ?? workbookPart
                .AddPartOfType(
                    package,
                    OoxmlPartTypes.Styles,
                    relationshipId: context.RelIdGenerator.GetNext(RelType.Workbook)
                )
                .Part;

        new StylesWriter().WriteContent(
            workbookStylesPart,
            XmlToEnumMapper.Instance,
            this.Styles,
            this,
            context
        );

        OpcPart sharedStringTablePart =
            workbookPart.PartOfType(OoxmlPartTypes.SharedStringTable)
            ?? workbookPart
                .AddPartOfType(
                    package,
                    OoxmlPartTypes.SharedStringTable,
                    relationshipId: context.RelIdGenerator.GetNext(RelType.Workbook)
                )
                .Part;

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
        SynchronizePivotTableParts(package, workbookPart, allPivotTables, context);

        // Phase 2 - All parts and relationships are set, fill in the parts.
        if (allPivotTables.Any())
        {
            this.GeneratePivotCaches(package, workbookPart, context);
        }

        foreach (
            XLWorksheet worksheet in this
                .WorksheetsInternal.Cast<XLWorksheet>()
                .OrderBy(w => w.Position)
        )
        {
            OpcPart worksheetPart;
            string wsRelId = worksheet.RelId;
            bool partIsEmpty;
            if (workbookPart.Relationships.TryGetById(wsRelId, out _))
            {
                worksheetPart = workbookPart.GetRelatedPart(wsRelId);
                partIsEmpty = false;
            }
            else
            {
                (worksheetPart, _) = workbookPart.AddPartOfType(
                    package,
                    OoxmlPartTypes.Worksheet,
                    relationshipId: wsRelId
                );
                partIsEmpty = true;
            }

            bool worksheetHasComments = worksheet
                .Internals.CellsCollection.GetCells(c => c.HasComment)
                .Any();

            OpcPart commentsPart = worksheetPart.PartOfType(OoxmlPartTypes.Comments);

            // VML part is the source of truth for shapes of comments, form controls and likely others.
            // Excel won't display any shape without VML. The drawing part is always present, but is likely
            // only different rendering of VML (more precisely the shapes behind VML).
            OpcPart vmlDrawingPart = worksheetPart.PartOfType(OoxmlPartTypes.VmlDrawing);
            bool hasAnyVmlElements = DeleteExistingCommentsShapes(vmlDrawingPart);

            if (worksheetHasComments)
            {
                // If sheet has comments, we must keep VML in legacy drawing part to display them
                // as well as comments part for semantic reasons.
                if (commentsPart == null)
                {
                    (commentsPart, _) = worksheetPart.AddPartOfType(
                        package,
                        OoxmlPartTypes.Comments,
                        relationshipId: context.RelIdGenerator.GetNext(RelType.Workbook)
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

                    (vmlDrawingPart, _) = worksheetPart.AddPartOfType(
                        package,
                        OoxmlPartTypes.VmlDrawing,
                        relationshipId: worksheet.LegacyDrawingId
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
                    package.DeletePart(commentsPart.Name);
                }
            }

            if (!hasAnyVmlElements && vmlDrawingPart is not null)
            {
                worksheet.LegacyDrawingId = null;
                package.DeletePart(vmlDrawingPart.Name);
            }

            XLTables xlTables = worksheet.Tables;

            // The way forward is to have 2-phase save, this is a start of that
            // concept for tables:
            //
            // Phase 1 - synchronize part existence with tables xlWorksheet, so each
            // table has a corresponding part and part that don't are deleted.
            // This phase doesn't modify the content, it only ensures that RelIds are set
            // corresponding parts exist and the parts that don't exist are removed
            TablePartWriter.SynchronizeTableParts(package, xlTables, worksheetPart, context);

            // Phase 2 - At this point, all pieces must have corresponding parts
            // The only way to link between parts is through RelIds that were already
            // set in phase 1. The phase 2 is all about content of individual parts.
            // Each part should have individual writer.
            TablePartWriter.GenerateTableParts(xlTables, worksheetPart, context);

            WorksheetPartWriter.GenerateWorksheetPartContent(
                partIsEmpty,
                package,
                worksheetPart,
                worksheet,
                options,
                context
            );

            if (worksheet.PivotTables.Any<XLPivotTable>())
            {
                GeneratePivotTables(package, workbookPart, worksheetPart, worksheet, context);
            }
        }

        if (options.GenerateCalculationChain)
        {
            CalculationChainPartWriter.GenerateContent(package, workbookPart, this, context);
        }
        else
        {
            if (workbookPart.PartOfType(OoxmlPartTypes.CalculationChain) is { } calcChainPart)
            {
                package.DeletePart(calcChainPart.Name);
            }
        }

        if (workbookPart.PartOfType(OoxmlPartTypes.Theme) is null)
        {
            (OpcPart themePart, _) = workbookPart.AddPartOfType(
                package,
                OoxmlPartTypes.Theme,
                relationshipId: context.RelIdGenerator.GetNext(RelType.Workbook)
            );
            ThemePartWriter.GenerateContent(themePart, (XLTheme)this.Theme);
        }

        // Custom properties
        if (this.CustomProperties.Any())
        {
            OpcPart customFilePropertiesPart =
                package.PartOfType(OoxmlPartTypes.CustomFileProperties)
                ?? package
                    .AddPartOfType(
                        OoxmlPartTypes.CustomFileProperties,
                        relationshipId: context.RelIdGenerator.GetNext(RelType.Workbook)
                    )
                    .Part;

            CustomFilePropertiesPartWriter.GenerateContent(customFilePropertiesPart, this);
        }
        else
        {
            if (
                package.PartOfType(OoxmlPartTypes.CustomFileProperties) is
                { } customFilePropertiesPart
            )
            {
                package.DeletePart(customFilePropertiesPart.Name);
            }
        }
        this.SetPackageProperties(package);

        // Last, because the pivot cache references it writes are only known once the cache parts
        // above have been created.
        WorkbookPartWriter.GenerateContent(workbookPart, this, options, context);

        // Clear list of deleted worksheets to prevent errors on multiple saves
        worksheets.Deleted.Clear();
    }

    private static bool DeleteExistingCommentsShapes(OpcPart vmlDrawingPart)
    {
        if (vmlDrawingPart == null)
        {
            return false;
        }

        // Nuke the VmlDrawingPart elements for comments.
        XDocument xdoc;
        using (Stream vmlStream = vmlDrawingPart.GetReadStream())
        {
            xdoc = XDocumentExtensions.Load(vmlStream);
        }

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

        using (Stream vmlStream = vmlDrawingPart.GetWriteStream())
        using (XmlTextWriter writer = new(vmlStream, Encoding.UTF8))
        {
            writer.WriteRaw(xdoc.ToXmlString());
        }

        return xdoc.Root.HasElements;
    }

    private void SetPackageProperties(OpcPackage package)
    {
        DateTime created =
            this.Properties.Created == DateTime.MinValue ? DateTime.Now : this.Properties.Created;
        DateTime modified =
            this.Properties.Modified == DateTime.MinValue ? DateTime.Now : this.Properties.Modified;
        package.Properties.Created = created;
        package.Properties.Modified = modified;

        package.Properties.LastModifiedBy = this.Properties.LastModifiedBy;

        package.Properties.Creator = this.Properties.Author;
        package.Properties.Title = this.Properties.Title;
        package.Properties.Subject = this.Properties.Subject;
        package.Properties.Category = this.Properties.Category;
        package.Properties.Keywords = this.Properties.Keywords;
        package.Properties.Description = this.Properties.Comments;
        package.Properties.ContentStatus = this.Properties.Status;
    }

    private static void SynchronizePivotTableParts(
        OpcPackage package,
        OpcPart workbookPart,
        IReadOnlyList<IXLPivotTable> allPivotTables,
        SaveContext context
    )
    {
        RemoveUnusedPivotCacheDefinitionParts(package, workbookPart, allPivotTables);
        AddUsedPivotCacheDefinitionParts(package, workbookPart, allPivotTables, context);

        // Ensure this in workbook.xml:
        //  <pivotCaches>
        //    <pivotCache cacheId="13" r:id="rId3"/>
        //  </pivotCaches>

        context.PivotSourceCacheId = 0;
        List<XLPivotCache> xlUsedCaches =
        [
            .. allPivotTables.Select(pt => pt.PivotCache).Distinct().Cast<XLPivotCache>(),
        ];
        // Only the cache ids are handed out here. The <pivotCaches> element they end up in is
        // written by WorkbookPartWriter, which runs after this and reads them back off the model.
        foreach (XLPivotCache source in xlUsedCaches)
        {
            source.CacheId = context.PivotSourceCacheId++;
        }

        // Remove pivot cache parts that are a part of the loaded document, but aren't used by a pivot table of the xlWorkbook
        // part of the first phase of saving
        static void RemoveUnusedPivotCacheDefinitionParts(
            OpcPackage package,
            OpcPart workbookPart,
            IReadOnlyList<IXLPivotTable> allPivotTables
        )
        {
            List<string> workbookCacheRelIds =
            [
                .. allPivotTables
                    .Select(pt => pt.PivotCache.CastTo<XLPivotCache>().WorkbookCacheRelId)
                    .Distinct(),
            ];

            List<OpcPart> orphanedParts =
            [
                .. workbookPart
                    .PartsOfType(OoxmlPartTypes.PivotCacheDefinition)
                    .Where(pcdp =>
                        !workbookCacheRelIds.Contains(
                            workbookPart.Relationships.GetIdOfTarget(pcdp.Name)
                        )
                    ),
            ];

            foreach (OpcPart orphanPart in orphanedParts)
            {
                if (orphanPart.PartOfType(OoxmlPartTypes.PivotCacheRecords) is { } recordsPart)
                {
                    package.DeletePart(recordsPart.Name);
                }

                package.DeletePart(orphanPart.Name);
            }
        }

        static void AddUsedPivotCacheDefinitionParts(
            OpcPackage package,
            OpcPart workbookPart,
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
                        || !workbookPart.Relationships.TryGetById(ps.WorkbookCacheRelId, out _)
                    )
                    .Distinct(),
            ];

            foreach (XLPivotCache pivotSource in newPivotSources)
            {
                string cacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                pivotSource.WorkbookCacheRelId = cacheRelId;

                workbookPart.AddPartOfType(
                    package,
                    OoxmlPartTypes.PivotCacheDefinition,
                    relationshipId: cacheRelId
                );
            }
        }
    }

    private void GeneratePivotCaches(OpcPackage package, OpcPart workbookPart, SaveContext context)
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
            // The <pivotCaches> element is written later, from the model; what has to hold here
            // is that the cache knows which part it lives in.
            Debug.Assert(xlPivotCache.CacheId is not null);
            Debug.Assert(!string.IsNullOrEmpty(xlPivotCache.WorkbookCacheRelId));

            OpcPart pivotTableCacheDefinitionPart = workbookPart.GetRelatedPart(
                xlPivotCache.WorkbookCacheRelId
            );

            PivotTableCacheDefinitionPartWriter.GenerateContent(
                pivotTableCacheDefinitionPart,
                xlPivotCache,
                context
            );

            OpcPart pivotTableCacheRecordsPart =
                pivotTableCacheDefinitionPart.PartOfType(OoxmlPartTypes.PivotCacheRecords)
                ?? pivotTableCacheDefinitionPart
                    .AddPartOfType(
                        package,
                        OoxmlPartTypes.PivotCacheRecords,
                        relationshipId: "rId1"
                    )
                    .Part;

            PivotCacheRecordsWriter.WriteContent(pivotTableCacheRecordsPart, xlPivotCache);
        }
    }

    private static void GeneratePivotTables(
        OpcPackage package,
        OpcPart workbookPart,
        OpcPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context
    )
    {
        foreach (XLPivotTable pt in xlWorksheet.PivotTables)
        {
            OpcPart pivotTablePart;
            bool createNewPivotTablePart = string.IsNullOrWhiteSpace(pt.RelId);
            if (createNewPivotTablePart)
            {
                string relId = context.RelIdGenerator.GetNext(RelType.Workbook);
                pt.RelId = relId;
                (pivotTablePart, _) = worksheetPart.AddPartOfType(
                    package,
                    OoxmlPartTypes.PivotTable,
                    relationshipId: relId
                );
            }
            else
            {
                pivotTablePart = worksheetPart.GetRelatedPart(pt.RelId);
            }

            XLPivotCache pivotSource = pt.PivotCache;
            OpcPart pivotTableCacheDefinitionPart = pivotTablePart.PartOfType(
                OoxmlPartTypes.PivotCacheDefinition
            );
            OpcPart expectedCacheDefinitionPart = workbookPart.GetRelatedPart(
                pivotSource.WorkbookCacheRelId
            );

            if (!ReferenceEquals(expectedCacheDefinitionPart, pivotTableCacheDefinitionPart))
            {
                // The cache definition part is shared with the workbook part, so only the
                // relationship pointing at it from here is dropped, never the part itself.
                if (
                    pivotTableCacheDefinitionPart is not null
                    && pivotTablePart.Relationships.GetIdOfTarget(
                        pivotTableCacheDefinitionPart.Name
                    )
                        is { } existingRelId
                )
                {
                    pivotTablePart.Relationships.Remove(existingRelId);
                }

                pivotTablePart.Relationships.Add(
                    expectedCacheDefinitionPart.Name,
                    OoxmlPartTypes.PivotCacheDefinition.RelationshipType,
                    context.RelIdGenerator.GetNext(RelType.Workbook)
                );
            }

            PivotTableDefinitionPartWriter.WriteContent(pivotTablePart, pt, context);
        }
    }
}

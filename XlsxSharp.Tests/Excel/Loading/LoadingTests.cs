using System.Globalization;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.PivotValues;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using Point = System.Drawing.Point;

namespace XlsxSharp.Tests.Excel.Loading;

// Tests in this fixture test only the successful loading of existing Excel files,
// i.e. we test that XlsxSharp doesn't choke on a given input file
// These tests DO NOT test that XlsxSharp successfully recognises all the Excel parts or that it can successfully save those parts again.
public class LoadingTests
{
    internal static IEnumerable<string> TryToLoad =>
        TestHelper.ListResourceFiles(s =>
            s.Contains(".TryToLoad.") && !s.Contains(".LO.") && !s.Contains(".Malformed.")
        );

    [Test]
    [MethodDataSource(nameof(TryToLoad))]
    public void CanSuccessfullyLoadFiles(string file) => TestHelper.LoadFile(file);

    [Test]
    [MethodDataSource(nameof(LoFiles))]
    public void CanSuccessfullyLoadLoFiles(string file) => TestHelper.LoadFile(file);

    internal static IEnumerable<string> LoFiles
    {
        get
        {
            // TODO: unpark all files
            string[] parkedForLater =
            [
                "TryToLoad.LO.xlsx.column-style-autofilter.xlsx",
                "TryToLoad.LO.xlsx.formats.xlsx",
                "TryToLoad.LO.xlsx.pivot_table.shared-group-field.xlsx",
                "TryToLoad.LO.xlsx.pivot_table.shared-nested-dategroup.xlsx",
                "TryToLoad.LO.xlsx.pivottable_bool_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_date_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_double_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_duplicated_member_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_rowcolpage_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_string_field_filter.xlsx",
                "TryToLoad.LO.xlsx.pivottable_tabular_mode.xlsx",
                "TryToLoad.LO.xlsx.pivot_table_first_header_row.xlsx",
                "TryToLoad.LO.xlsx.tdf100709.xlsx",
                "TryToLoad.LO.xlsx.tdf89139_pivot_table.xlsx",
                "TryToLoad.LO.xlsx.universal-content-strict.xlsx",
                "TryToLoad.LO.xlsx.universal-content.xlsx",
                "TryToLoad.LO.xlsx.xf_default_values.xlsx",
                "TryToLoad.LO.xlsm.pass.CVE-2016-0122-1.xlsm",
                "TryToLoad.LO.xlsm.tdf111974.xlsm",
                "TryToLoad.LO.xlsm.vba-user-function.xlsm",
            ];

            return TestHelper.ListResourceFiles(s =>
                s.Contains(".LO.") && !parkedForLater.Any(i => s.Contains(i))
            );
        }
    }

    [Test]
    public void CorrectlyLoadValidationWithSheetReference()
    {
        // Arrange
        string path = TestHelper.GetResourcePath(@"TryToLoad\ValidationWithSheetReference.xlsx");
        using Stream stream = TestHelper.GetStreamFromResource(path);

        // Act
        using XLWorkbook wb = new(stream);

        // Assert
        IXLWorksheet ws = wb.Worksheet("UI Sheet");
        IXLCell B2 = ws.Cell("B2");
        ClassicAssert.AreEqual(XLAllowedValues.List, B2.GetDataValidation().AllowedValues);
        ClassicAssert.AreEqual("$E$1:$E$4", B2.GetDataValidation().Value);
        IXLCell A2 = ws.Cell("A2");
        ClassicAssert.AreEqual(XLAllowedValues.List, A2.GetDataValidation().AllowedValues);
        ClassicAssert.AreEqual("ValuesSheet!$A$1:$A$4", A2.GetDataValidation().Value);
    }

    [Test]
    public void CanLoadAndManipulateFileWithEmptyTable()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\EmptyTable.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            IXLTable table = ws.Tables.First();
            table.DataRange.InsertRowsBelow(5);
        }
    }

    [Test]
    public void CanLoadDate1904SystemCorrectly()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\Date1904System.xlsx")
            )
        )
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new(stream))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLCell c = ws.Cell("A2");
                ClassicAssert.AreEqual(XLDataType.DateTime, c.DataType);
                ClassicAssert.AreEqual(new DateTime(2017, 10, 27, 21, 0, 0), c.GetDateTime());
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                IXLCell c = ws.Cell("A2");
                ClassicAssert.AreEqual(XLDataType.DateTime, c.DataType);
                ClassicAssert.AreEqual(new DateTime(2017, 10, 27, 21, 0, 0), c.GetDateTime());
                wb.SaveAs(ms);
            }
        }
    }

    [Test]
    public void CanLoadAndSaveFileWithMismatchingSheetIdAndRelId()
    {
        // This file's workbook.xml contains:
        // <x:sheet name="Data" sheetId="13" r:id="rId1" />
        // and the mismatch between the sheetId and r:id can create problems.
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\FileWithMismatchSheetIdAndRelId.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            using (MemoryStream ms = new())
            {
                wb.SaveAs(ms, true);
            }
        }
    }

    [Test]
    public void CanLoadBasicPivotTable()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet("PivotTable1");
            IXLPivotTable pt = ws.PivotTable("PivotTable1");
            ClassicAssert.AreEqual("PivotTable1", pt.Name);

            ClassicAssert.AreEqual(1, pt.RowLabels.Count());
            ClassicAssert.AreEqual("Name", pt.RowLabels.Single().SourceName);

            ClassicAssert.AreEqual(1, pt.ColumnLabels.Count());
            ClassicAssert.AreEqual("Month", pt.ColumnLabels.Single().SourceName);

            IXLPivotValue pv = pt.Values.Single();
            ClassicAssert.AreEqual("Sum of NumberOfOrders", pv.CustomName);
            ClassicAssert.AreEqual("NumberOfOrders", pv.SourceName);
        }
    }

    [Test]
    public void CanLoadOrderedPivotTable()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet("OrderedPivotTable");
            IXLPivotTable pt = ws.PivotTable("OrderedPivotTable");

            ClassicAssert.AreEqual(XLPivotSortType.Ascending, pt.RowLabels.Single().SortType);
            ClassicAssert.AreEqual(XLPivotSortType.Descending, pt.ColumnLabels.Single().SortType);
        }
    }

    [Test]
    public void CanLoadPivotTableSubtotals()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\LoadPivotTables.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet("PivotTableSubtotals");
            IXLPivotTable pt = ws.PivotTable("PivotTableSubtotals");

            XLSubtotalFunction[] subtotals = [.. pt.RowLabels.Get("Group").Subtotals];

            CollectionAssert.AreEquivalent(
                new[]
                {
                    XLSubtotalFunction.Average,
                    XLSubtotalFunction.Count,
                    XLSubtotalFunction.Sum,
                },
                subtotals
            );
        }
    }

    [Test]
    [Skip("PT styles will be fixed in a different PR")]
    public void CanLoadPivotTableWithBorder()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\PivotTableWithBorder.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLPivotTable pt = wb.Worksheet(1).PivotTables.PivotTable("PivotTable1");
            IXLBorder border = pt.RowLabels.Single().StyleFormats.DataValuesFormat.Style.Border;

            ClassicAssert.AreEqual(XLBorderStyleValues.Thin, border.LeftBorder);
            ClassicAssert.AreEqual(XLBorderStyleValues.Thin, border.TopBorder);
            ClassicAssert.AreEqual(XLBorderStyleValues.Thin, border.RightBorder);
            ClassicAssert.AreEqual(XLBorderStyleValues.Thin, border.BottomBorder);
        }
    }

    /// <summary>
    /// For non-English locales, the default style ("Normal" in English) can be
    /// another piece of text (e.g. ??????? in Russian).
    /// This test ensures that the default style is correctly detected and
    /// no style conflicts occur on save.
    /// </summary>
    [Test]
    public void CanSaveFileWithDefaultStyleNameNotInEnglish()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\FileWithDefaultStyleNameNotInEnglish.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            using (MemoryStream ms = new())
            {
                wb.SaveAs(ms, true);
            }
        }
    }

    /// <summary>
    /// As per https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.cellvalues(v=office.15).aspx
    /// the 'Date' DataType is available only in files saved with Microsoft Office
    /// In other files, the data type will be saved as numeric
    /// XlsxSharp then deduces the data type by inspecting the number format string
    /// </summary>
    [Test]
    public void CanLoadLibreOfficeFileWithDates()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\LibreOfficeFileWithDates.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            foreach (IXLCell cell in ws.CellsUsed())
            {
                ClassicAssert.AreEqual(XLDataType.DateTime, cell.DataType);
            }
        }
    }

    [Test]
    public void CanLoadFileWithImagesWithCorrectAnchorTypes()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            ClassicAssert.AreEqual(2, ws.Pictures.Count);
            ClassicAssert.AreEqual(XLPicturePlacement.FreeFloating, ws.Pictures.First().Placement);
            ClassicAssert.AreEqual(XLPicturePlacement.Move, ws.Pictures.Skip(1).First().Placement);

            IXLWorksheet ws2 = wb.Worksheets.Skip(1).First();
            ClassicAssert.AreEqual(1, ws2.Pictures.Count);
            ClassicAssert.AreEqual(XLPicturePlacement.MoveAndSize, ws2.Pictures.First().Placement);
        }
    }

    [Test]
    public void CanLoadFileWithImagesWithCorrectImageType()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageFormats.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            ClassicAssert.AreEqual(1, ws.Pictures.Count);
            ClassicAssert.AreEqual(XLPictureFormat.Jpeg, ws.Pictures.First().Format);

            IXLWorksheet ws2 = wb.Worksheets.Skip(1).First();
            ClassicAssert.AreEqual(1, ws2.Pictures.Count);
            ClassicAssert.AreEqual(XLPictureFormat.Png, ws2.Pictures.First().Format);
        }
    }

    [Test]
    public void CanLoadAndDeduceAnchorsFromExcelGeneratedFile()
    {
        // This file was produced by Excel. It contains 3 images, but the latter 2 were copied from the first.
        // There is actually only 1 embedded image if you inspect the file's internals.
        // Additionally, Excel saves all image anchors as TwoCellAnchor, but uses the EditAs attribute to distinguish the types
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\ExcelProducedWorkbookWithImages.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            ClassicAssert.AreEqual(3, ws.Pictures.Count);

            ClassicAssert.AreEqual(
                XLPicturePlacement.MoveAndSize,
                ws.Picture("Picture 1").Placement
            );
            ClassicAssert.AreEqual(XLPicturePlacement.Move, ws.Picture("Picture 2").Placement);
            ClassicAssert.AreEqual(
                XLPicturePlacement.FreeFloating,
                ws.Picture("Picture 3").Placement
            );

            using (MemoryStream ms = new())
            {
                wb.SaveAs(ms, true);
            }
        }
    }

    [Test]
    public void CanLoadFromTemplate()
    {
        using (TemporaryFile tf1 = new())
        using (TemporaryFile tf2 = new())
        {
            using (
                Stream stream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(@"TryToLoad\AllShapes.xlsx")
                )
            )
            using (XLWorkbook wb = new(stream))
            {
                // Save as temporary file
                wb.SaveAs(tf1.Path);
            }

            XLWorkbook workbook = XLWorkbook.OpenFromTemplate(tf1.Path);
            ClassicAssert.True(workbook.Worksheets.Any());
            ClassicAssert.Throws<InvalidOperationException>(() => workbook.Save());

            workbook.SaveAs(tf2.Path);
        }
    }

    /// <summary>
    /// Excel escapes symbol ' in worksheet title so we have to process this correctly.
    /// </summary>
    [Test]
    public void CanOpenWorksheetWithEscapedApostrophe()
    {
        string title = "";
        TestDelegate openWorkbook = () =>
        {
            using (
                Stream stream = TestHelper.GetStreamFromResource(
                    TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx")
                )
            )
            using (XLWorkbook wb = new(stream))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                title = ws.Name;
            }
        };

        ClassicAssert.DoesNotThrow(openWorkbook);
        ClassicAssert.AreEqual("L'E", title);
    }

    [Test]
    public void CanRoundTripSheetProtectionForObjects()
    {
        using (XLWorkbook book = new())
        {
            IXLWorksheet sheet = book.AddWorksheet("TestSheet");
            sheet
                .Protect()
                .AllowElement(
                    XLSheetProtectionElements.EditObjects | XLSheetProtectionElements.EditScenarios
                );

            ClassicAssert.AreEqual(
                XLSheetProtectionElements.SelectEverything
                    | XLSheetProtectionElements.EditObjects
                    | XLSheetProtectionElements.EditScenarios,
                sheet.Protection.AllowedElements
            );

            using (MemoryStream xlStream = new())
            {
                book.SaveAs(xlStream);

                using (XLWorkbook persistedBook = new(xlStream))
                {
                    IXLWorksheet persistedSheet = persistedBook.Worksheets.Worksheet(1);

                    ClassicAssert.AreEqual(
                        sheet.Protection.AllowedElements,
                        persistedSheet.Protection.AllowedElements
                    );
                }
            }
        }
    }

    [Test]
    [Arguments("A1*10", 1230)]
    [Arguments("A1/10", 12.3)]
    [Arguments("A1&\" cells\"", "123 cells")]
    [Arguments("A1&\"000\"", "123000")]
    [Arguments("ISNUMBER(A1)", true)]
    [Arguments("ISBLANK(A1)", false)]
    [Arguments("DATE(2018,1,28)", 43128)]
    public void LoadFormulaCachedValue(string formula, object expectedCachedValue)
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook book1 = new())
            {
                IXLWorksheet sheet = book1.AddWorksheet("sheet1");
                sheet.Cell("A1").Value = 123;
                sheet.Cell("A2").FormulaA1 = formula;
                SaveOptions options = new() { EvaluateFormulasBeforeSaving = true };

                book1.SaveAs(ms, options);
            }
            ms.Position = 0;

            using (XLWorkbook book2 = new(ms))
            {
                IXLWorksheet ws = book2.Worksheet(1);
                ClassicAssert.IsFalse(ws.Cell("A2").NeedsRecalculation);
                ClassicAssert.AreEqual(expectedCachedValue, ws.Cell("A2").CachedValue);
            }
        }
    }

    [Test]
    public void LoadingOptions()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\Misc\Formulas.xlsx")
            )
        )
        {
            ClassicAssert.DoesNotThrow(() =>
            {
                // The value in the file is blank and kept.
                using XLWorkbook wb = new(
                    stream,
                    new LoadOptions { RecalculateAllFormulas = false }
                );
                ClassicAssert.AreEqual(Blank.Value, wb.Worksheets.Single().Cell("C2").CachedValue);
            });

            ClassicAssert.DoesNotThrow(() =>
            {
                // The value in the file is blank, but recalculation sets it to correct 3.
                using XLWorkbook wb = new(
                    stream,
                    new LoadOptions { RecalculateAllFormulas = true }
                );
                ClassicAssert.AreEqual(3, wb.Worksheets.Single().Cell("C2").CachedValue);
            });

            ClassicAssert.AreEqual(
                30,
                new XLWorkbook(stream, new LoadOptions { Dpi = new Point(30, 14) }).DpiX
            );
            ClassicAssert.AreEqual(
                14,
                new XLWorkbook(stream, new LoadOptions { Dpi = new Point(30, 14) }).DpiY
            );
        }
    }

    [Test]
    public void CanLoadWorksheetStyle()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\BaseColumnWidth.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);

            ClassicAssert.AreEqual(8, ws.Style.Font.FontSize);
            ClassicAssert.AreEqual("Arial", ws.Style.Font.FontName);
            ClassicAssert.AreEqual(8, ws.Cell("A1").Style.Font.FontSize);
            ClassicAssert.AreEqual("Arial", ws.Cell("A1").Style.Font.FontName);
        }
    }

    [Test]
    public void CanCorrectLoadWorkbookCellWithStringDataType()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\CellWithStringDataType.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLCell cellToCheck = wb.Worksheet(1).Cell("B2");
            ClassicAssert.AreEqual(XLDataType.Text, cellToCheck.DataType);
            ClassicAssert.AreEqual("String with String Data type", cellToCheck.Value);
        }
    }

    [Test]
    public void CanCorrectLoadWorkbookCellsWithDateTimeDataTypeOrFormatting()
    {
        const string expected = "03/14/2012 13:30:55";
        TestHelper.LoadAndAssert(
            wb =>
            {
                for (int row = 2; row < 18; row++)
                {
                    IXLCell cellToCheck = wb.Worksheet(1).Cell(row, 2);
                    ClassicAssert.AreEqual(
                        XLDataType.DateTime,
                        cellToCheck.DataType,
                        $"Cell B{row} has incorrect DataType"
                    );
                    ClassicAssert.AreEqual(
                        expected,
                        cellToCheck.Value.ToString(CultureInfo.InvariantCulture),
                        $"Cell B{row} value differs"
                    );
                }
            },
            @"TryToLoad\CellsWithDateTimeDataTypeOrFormatting.xlsx"
        );
    }

    [Test]
    public void CanCorrectLoadWorkbookCellsWithTimeSpanDataTypeOrFormatting()
    {
        string[] expected = [.. Enumerable.Range(0, 10).Select(_ => "13:30:55.2"), "0:30:55.2"];
        TestHelper.LoadAndAssert(
            wb =>
            {
                for (int i = 0, row = 2; i < expected.Length; i++, row++)
                {
                    IXLCell cellToCheck = wb.Worksheet(1).Cell(row, 2);
                    ClassicAssert.AreEqual(
                        XLDataType.TimeSpan,
                        cellToCheck.DataType,
                        $"Cell B{row} has incorrect DataType"
                    );
                    ClassicAssert.AreEqual(
                        expected[i],
                        cellToCheck.Value.ToString(CultureInfo.InvariantCulture),
                        $"Cell B{row} value differs"
                    );
                }
            },
            @"TryToLoad\CellsWithTimeSpanDataTypeOrFormatting.xlsx"
        );
    }

    [Test]
    public void CanCorrectLoadWorkbookCellsWithDateTimesWithLocalePrefix() =>
        TestHelper.LoadAndAssert(
            wb =>
            {
                IXLWorksheet ws = wb.Worksheet(1);

                ClassicAssert.AreEqual("21 January 2019", ws.Cell(1, 1).GetFormattedString());
                ClassicAssert.AreEqual("21-Jan-19", ws.Cell(2, 1).GetFormattedString());
                ClassicAssert.AreEqual(
                    "Monday, 21 January 2019",
                    ws.Cell(3, 1).GetFormattedString()
                );
                ClassicAssert.AreEqual("21 Jan 2019", ws.Cell(4, 1).GetFormattedString());
            },
            @"TryToLoad\CellsWithDateTimeWithLocalePrefix.xlsx"
        );

    [Test]
    public void CanCorrectLoadWorkbookDefaultColumnWidth()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\Styles\DefaultStyles.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            double defaultColumnWidth = wb.ColumnWidth;
            int pixelWidth = XLHelper.NoCToPixels(defaultColumnWidth, wb.Format.Font, wb);
            ClassicAssert.AreEqual(8.43, defaultColumnWidth, XLHelper.Epsilon);
            ClassicAssert.AreEqual(64, pixelWidth);
        }
    }

    [Test]
    public void CanCorrectLoadWorkbookDefaultColumnWidthOfNonDefaultFont()
    {
        // The width is derived from the metric of the font of the workbook, so the numbers Excel
        // wrote can only be reproduced on a machine that has that font.
        TestHelper.IgnoreIfFontIsMissing("Arial");

        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\DefaultColumnWidth.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            double defaultColumnWidth = wb.ColumnWidth;
            int pixelWidth = XLHelper.NoCToPixels(defaultColumnWidth, wb.Format.Font, wb);
            ClassicAssert.AreEqual(8.5, defaultColumnWidth, XLHelper.Epsilon);
            ClassicAssert.AreEqual(56, pixelWidth);
        }
    }

    [Test]
    public void CanCorrectLoadWorksheetBaseColumnWidth()
    {
        // default calibi font case
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\Styles\DefaultStyles.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);
            ClassicAssert.AreEqual(8.43, ws.ColumnWidth, XLHelper.Epsilon);
            ClassicAssert.AreEqual(8.43, ws.Column(1).Width, XLHelper.Epsilon);
        }
    }

    [Test]
    public void CanCorrectLoadWorksheetBaseColumnWidthOfNonDefaultFont()
    {
        // worksheet has base column width, converted through the metric of the font of the worksheet.
        TestHelper.IgnoreIfFontIsMissing("Arial");

        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\BaseColumnWidth.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);
            ClassicAssert.AreEqual(11.17, ws.ColumnWidth, XLHelper.Epsilon);
            ClassicAssert.AreEqual(11.17, ws.Column(1).Width, XLHelper.Epsilon);
        }
    }

    [Test]
    public void CanCorrectLoadWorksheetDefaultColumnWidth()
    {
        // The worksheet has a default column width of 20.375 MDWs and the font of the workbook
        // (游ゴシック) is not installed on most machines, so the default engine would measure
        // whatever font it falls back to. Unlike the tests above, the presence of the font of the
        // workbook is not what makes the numbers reproducible here - the maximum digit width is,
        // so the test supplies it and checks the conversion done during the load.
        LoadOptions loadOptions = new() { GraphicEngine = new FixedMaxDigitWidthEngine(8) };

        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\SheetDefaultColumnWidth.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream, loadOptions))
        {
            IXLWorksheet ws = wb.Worksheet(1);
            double pixelWidth = XLHelper.NoCToPixels(ws.Column(1).Width, ws.Style.Font, wb);
            ClassicAssert.AreEqual(19.75, ws.ColumnWidth, XLHelper.Epsilon);
            ClassicAssert.AreEqual(163, pixelWidth, XLHelper.Epsilon);
        }
    }

    [Test]
    public void CanLoadFileWithInvalidSelectedRanges()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Other\SelectedRanges\InvalidSelectedRange.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);

            ClassicAssert.AreEqual(2, ws.SelectedRanges.Count);
            ClassicAssert.AreEqual("B2:B2", ws.SelectedRanges.First().RangeAddress.ToString());
            ClassicAssert.AreEqual("B2:C2", ws.SelectedRanges.Last().RangeAddress.ToString());
        }
    }

    [Test]
    public void CanLoadCellsWithoutReferencesCorrectly()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\LO\xlsx\row-index-1-based.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);

            ClassicAssert.AreEqual("Page 1", ws.Name);

            Dictionary<string, XLCellValue> expected = new()
            {
                ["A1"] = "Action Plan.Name",
                ["B1"] = "Action Plan.Description",
                ["A2"] = "Jerry",
                ["B2"] = "This is a longer Text.\nSecond line.\nThird line.",
                ["A3"] = Blank.Value,
                ["B3"] = Blank.Value,
            };

            foreach (KeyValuePair<string, XLCellValue> pair in expected)
            {
                ClassicAssert.AreEqual(pair.Value, ws.Cell(pair.Key).Value, pair.Key);
            }
        }
    }

    [Test]
    public void CorrectlyLoadThemeColors()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\ThemeColors\inputfile.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);

            IXLCell c = ws.Cell("A1");
            XLThemeColor themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
            ClassicAssert.AreEqual(XLThemeColor.Accent2, themeColor);
            ClassicAssert.AreEqual(
                "FFED7D31",
                wb.Theme.ResolveThemeColor(themeColor).Color.ToHex()
            );

            c = ws.Cell("A2");
            themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
            ClassicAssert.AreEqual(XLThemeColor.Accent4, themeColor);
            ClassicAssert.AreEqual(
                "FFFFC000",
                wb.Theme.ResolveThemeColor(themeColor).Color.ToHex()
            );

            c = ws.Cell("A3");
            themeColor = c.Style.Fill.BackgroundColor.ThemeColor;
            ClassicAssert.AreEqual(XLThemeColor.Accent6, themeColor);
            ClassicAssert.AreEqual(
                "FF70AD47",
                wb.Theme.ResolveThemeColor(themeColor).Color.ToHex()
            );
        }
    }

    [Test]
    public void CorrectlyLoadMergedCellsBorder()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(
                    @"Other\StyleReferenceFiles\MergedCellsBorder\inputfile.xlsx"
                )
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);

            IXLCell c = ws.Cell("B2");
            ClassicAssert.AreEqual(XLColorType.Theme, c.Style.Border.TopBorderColor.ColorType);
            ClassicAssert.AreEqual(XLThemeColor.Accent1, c.Style.Border.TopBorderColor.ThemeColor);
            ClassicAssert.AreEqual(
                0.39994506668294322d,
                c.Style.Border.TopBorderColor.ThemeTint,
                XLHelper.Epsilon
            );
        }
    }

    [Test]
    public void CorrectlyLoadDefaultRowAndColumnStyles()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(
                    @"Other\StyleReferenceFiles\RowAndColumnStyles\inputfile.xlsx"
                )
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);

            ClassicAssert.AreEqual(8, ws.Row(1).Style.Font.FontSize);
            ClassicAssert.AreEqual(8, ws.Row(2).Style.Font.FontSize);
            ClassicAssert.AreEqual(8, ws.Column("A").Style.Font.FontSize);
        }
    }

    [Test]
    public void EmptyNumberFormatIdTreatedAsGeneral()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"TryToLoad\EmptyNumberFormatId.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheet(1);

            ClassicAssert.AreEqual(
                XLPredefinedFormat.General,
                ws.Cell("A2").Style.NumberFormat.NumberFormatId
            );
        }
    }

    [Test]
    public void CanLoadProperties()
    {
        const string author = "TestAuthor";
        const string title = "TestTitle";
        const string subject = "TestSubject";
        const string category = "TestCategory";
        const string keywords = "TestKeywords";
        const string comments = "TestComments";
        const string status = "TestStatus";
        DateTime created = new(2019, 10, 19, 20, 42, 30);
        DateTime modified = new(2020, 11, 20, 09, 51, 20);
        const string lastModifiedBy = "TestLastModifiedBy";
        const string company = "TestCompany";
        const string manager = "TestManager";

        using (MemoryStream stream = new())
        {
            using (XLWorkbook wb = new())
            {
                wb.AddWorksheet("sheet1");

                wb.Properties.Author = author;
                wb.Properties.Title = title;
                wb.Properties.Subject = subject;
                wb.Properties.Category = category;
                wb.Properties.Keywords = keywords;
                wb.Properties.Comments = comments;
                wb.Properties.Status = status;
                wb.Properties.Created = created;
                wb.Properties.Modified = modified;
                wb.Properties.LastModifiedBy = lastModifiedBy;
                wb.Properties.Company = company;
                wb.Properties.Manager = manager;

                wb.SaveAs(stream, true);
            }

            stream.Position = 0;

            using (XLWorkbook wb = new(stream))
            {
                ClassicAssert.AreEqual(author, wb.Properties.Author);
                ClassicAssert.AreEqual(title, wb.Properties.Title);
                ClassicAssert.AreEqual(subject, wb.Properties.Subject);
                ClassicAssert.AreEqual(category, wb.Properties.Category);
                ClassicAssert.AreEqual(keywords, wb.Properties.Keywords);
                ClassicAssert.AreEqual(comments, wb.Properties.Comments);
                ClassicAssert.AreEqual(status, wb.Properties.Status);
                ClassicAssert.AreEqual(created, wb.Properties.Created);
                ClassicAssert.AreEqual(modified, wb.Properties.Modified);
                ClassicAssert.AreEqual(lastModifiedBy, wb.Properties.LastModifiedBy);
                ClassicAssert.AreEqual(company, wb.Properties.Company);
                ClassicAssert.AreEqual(manager, wb.Properties.Manager);
            }
        }
    }

    [Test]
    public void CanLoadEmptyStyles() =>
        // Stylesheet part exists, but no style collection elements are present
        TestHelper.LoadAndAssert(
            wb =>
            {
                using MemoryStream ms = new();
                wb.SaveAs(ms, true);
            },
            @"TryToLoad\EmptyStyles.xlsx"
        );

    [Test]
    public void CanLoadInvalidColors() =>
        // The styles.xml contains two invalid colors: '0' and 'FED+'. Both
        // should be loaded and no exception thrown. The colors are
        // converted using an Excel algorithm.
        TestHelper.LoadAndAssert(
            wb =>
            {
                IXLWorksheet ws = wb.Worksheets.Single();
                ClassicAssert.AreEqual(
                    XLColor.FromArgb(0xFF000000),
                    ws.Cell("A1").Style.Font.FontColor
                );
                ClassicAssert.AreEqual(
                    XLColor.FromArgb(0xFF000FED),
                    ws.Cell("A2").Style.Fill.BackgroundColor
                );
            },
            @"TryToLoad\InvalidColors.xlsx"
        );

    [Test]
    public void WontCrashOnSheetsWithoutRelId() =>
        // Some non-Excel producers create workbooks where workbookPart declares
        // sheet with empty r:id, but with name and sheetId. Content of such sheets
        // isn't loaded even if relationship part declares implicit relationship to
        // the worksheets, because workbook has explicit relationships with worksheet
        // part (ISO29500 12.3.23).
        //
        // If excel finds sheet in workbook without r:id, it adds empty sheet with
        // the specified name and so does XlsxSharp.
        TestHelper.LoadAndAssert(
            wb =>
            {
                ClassicAssert.AreEqual(3, wb.Worksheets.Count);

                // First sheet has r:id, so it keeps content
                ClassicAssert.AreEqual("Sheet1", wb.Worksheet("Sheet1").Cell("A1").Value);

                // Second sheet doesn't have r:id, so it is empty after load.
                ClassicAssert.AreEqual(
                    Blank.Value,
                    wb.Worksheet("Sheet without relId").Cell("A1").Value
                );

                // Third sheet doesn't have r:id and it contains pivot table that is not loaded.
                IXLWorksheet ptSheet = wb.Worksheet("Pivot Sheet without relId");
                ClassicAssert.AreEqual(Blank.Value, ptSheet.Cell("A1").Value);
                ClassicAssert.False(ptSheet.PivotTables.Any());
            },
            @"TryToLoad\SheetsWithoutRelId.xlsx"
        );

    [Test]
    public void CanLoadDialogSheet() =>
        // Workbook can reference multiple different types of sheet, most common is worksheet,
        // but there is also possibility of referencing dialogSheet (basically VBA dialog).
        // dialogSheet is basically obsolete (from Excel 5.0), but still supported. Do not
        // crash when such sheet is encountered. Test file also contains pivot table, because
        // it originally crashed just before pivot table loading.
        TestHelper.LoadAndAssert(
            wb =>
            {
                // Dialog sheet
                ClassicAssert.AreEqual(1, wb.UnsupportedSheets.Count);

                // Data and pivot sheets
                ClassicAssert.AreEqual(2, wb.Worksheets.Count);
                ClassicAssert.NotNull(wb.Worksheet("Pivot").PivotTables.Contains("PivotTable1"));
            },
            @"TryToLoad\DialogSheet.xlsx"
        );

    [Test]
    public void CanLoadWorkbookWithInvalidAttributesWhenStrictParsingIsDisabled() =>
        TestHelper.LoadAndAssert(
            (_, ws) =>
            {
                // "Center" - wrong case.
                ClassicAssert.AreEqual(
                    XLAlignmentVerticalValues.Bottom,
                    ws.Cell("A1").Style.Alignment.Vertical
                );

                // "richtig" - invalid value.
                ClassicAssert.AreEqual(
                    XLAlignmentHorizontalValues.General,
                    ws.Cell("A2").Style.Alignment.Horizontal
                );

                // "three" - not a number.
                ClassicAssert.AreEqual(0, ws.Cell("A3").Style.Alignment.Indent);

                // "PRAVDA" - not a bool, thus treated as a missing. Thanks to the default true it is still bold.
                ClassicAssert.IsTrue(ws.Cell("A4").Style.Font.Bold);
            },
            @"TryToLoad\Malformed\AttributesWithInvalidValues.xlsx",
            new LoadOptions { StrictAttributeParsing = false }
        );
}

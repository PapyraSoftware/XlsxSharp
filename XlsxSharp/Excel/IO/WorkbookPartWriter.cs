#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Extensions;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel.IO;

internal class WorkbookPartWriter
{
    internal static void GenerateContent(
        WorkbookPart workbookPart,
        XLWorkbook xlWorkbook,
        SaveOptions options,
        XLWorkbook.SaveContext context
    )
    {
        if (workbookPart.Workbook == null)
        {
            workbookPart.Workbook = new Workbook();
        }

        Workbook workbook = workbookPart.Workbook;
        if (
            !workbook.NamespaceDeclarations.Contains(
                new KeyValuePair<string, string>(
                    "r",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                )
            )
        )
        {
            workbook.AddNamespaceDeclaration(
                "r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            );
        }

        #region WorkbookProperties

        if (workbook.WorkbookProperties == null)
        {
            workbook.WorkbookProperties = new WorkbookProperties();
        }

        if (workbook.WorkbookProperties.CodeName == null)
        {
            workbook.WorkbookProperties.CodeName = "ThisWorkbook";
        }

        workbook.WorkbookProperties.Date1904 = OpenXmlHelper.GetBooleanValue(
            xlWorkbook.Use1904DateSystem,
            false
        );

        if (options.FilterPrivacy.HasValue)
        {
            workbook.WorkbookProperties.FilterPrivacy = OpenXmlHelper.GetBooleanValue(
                options.FilterPrivacy.Value,
                false
            );
        }

        #endregion WorkbookProperties

        #region FileSharing

        if (workbook.FileSharing == null)
        {
            workbook.FileSharing = new FileSharing();
        }

        workbook.FileSharing.ReadOnlyRecommended = OpenXmlHelper.GetBooleanValue(
            xlWorkbook.FileSharing.ReadOnlyRecommended,
            false
        );
        workbook.FileSharing.UserName = string.IsNullOrWhiteSpace(xlWorkbook.FileSharing.UserName)
            ? null
            : StringValue.FromString(xlWorkbook.FileSharing.UserName);

        if (!workbook.FileSharing.HasChildren && !workbook.FileSharing.HasAttributes)
        {
            workbook.FileSharing = null;
        }

        #endregion FileSharing

        #region WorkbookProtection

        if (xlWorkbook.Protection.IsProtected)
        {
            if (workbook.WorkbookProtection == null)
            {
                workbook.WorkbookProtection = new WorkbookProtection();
            }

            WorkbookProtection workbookProtection = workbook.WorkbookProtection;

            XLWorkbookProtection protection = xlWorkbook.Protection;

            workbookProtection.WorkbookPassword = null;
            workbookProtection.WorkbookAlgorithmName = null;
            workbookProtection.WorkbookHashValue = null;
            workbookProtection.WorkbookSpinCount = null;
            workbookProtection.WorkbookSaltValue = null;

            if (protection.Algorithm == XLProtectionAlgorithm.Algorithm.SimpleHash)
            {
                if (!string.IsNullOrWhiteSpace(protection.PasswordHash))
                {
                    workbookProtection.WorkbookPassword = protection.PasswordHash;
                }
            }
            else
            {
                workbookProtection.WorkbookAlgorithmName =
                    DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.ToDescription(
                        protection.Algorithm
                    );
                workbookProtection.WorkbookHashValue = protection.PasswordHash;
                workbookProtection.WorkbookSpinCount = protection.SpinCount;
                workbookProtection.WorkbookSaltValue = protection.Base64EncodedSalt;
            }

            workbookProtection.LockStructure = OpenXmlHelper.GetBooleanValue(
                !protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Structure),
                false
            );
            workbookProtection.LockWindows = OpenXmlHelper.GetBooleanValue(
                !protection.AllowedElements.HasFlag(XLWorkbookProtectionElements.Windows),
                false
            );
        }
        else
        {
            workbook.WorkbookProtection = null;
        }

        #endregion WorkbookProtection

        if (workbook.BookViews == null)
        {
            workbook.BookViews = new BookViews();
        }

        if (workbook.Sheets == null)
        {
            workbook.Sheets = new Sheets();
        }

        XLWorksheets worksheets = xlWorkbook.WorksheetsInternal;
        workbook
            .Sheets.Elements<Sheet>()
            .Where(s => worksheets.Deleted.Contains(s.Id))
            .ToList()
            .ForEach(s => s.Remove());

        foreach (Sheet sheet in workbook.Sheets.Elements<Sheet>())
        {
            int sheetId = (int)sheet.SheetId.Value;

            if (xlWorkbook.WorksheetsInternal.All<XLWorksheet>(w => w.SheetId != sheetId))
            {
                continue;
            }

            XLWorksheet wks = xlWorkbook.WorksheetsInternal.Single<XLWorksheet>(w =>
                w.SheetId == sheetId
            );
            wks.RelId = sheet.Id;
            sheet.Name = wks.Name;
        }

        foreach (
            XLWorksheet xlSheet in xlWorkbook.WorksheetsInternal.OrderBy<XLWorksheet, int>(w =>
                w.Position
            )
        )
        {
            string rId;
            if (string.IsNullOrWhiteSpace(xlSheet.RelId))
            {
                // Sheet isn't from loaded file and hasn't been saved yet.
                rId = xlSheet.RelId = context.RelIdGenerator.GetNext(XLWorkbook.RelType.Workbook);
            }
            else
            {
                // Keep same r:id from previous file
                rId = xlSheet.RelId;
            }

            if (workbook.Sheets.Cast<Sheet>().All(s => s.Id != rId))
            {
                Sheet newSheet = new()
                {
                    Name = xlSheet.Name,
                    Id = rId,
                    SheetId = xlSheet.SheetId,
                };

                workbook.Sheets.AppendChild(newSheet);
            }
        }

        IEnumerable<Sheet> sheetElements =
            from sheet in workbook.Sheets.Elements<Sheet>()
            join worksheet in ((IEnumerable<XLWorksheet>)xlWorkbook.WorksheetsInternal)
                on sheet.Id.Value equals worksheet.RelId
            orderby worksheet.Position
            select sheet;

        uint firstSheetVisible = 0;
        uint activeTab = (
            from us in xlWorkbook.UnsupportedSheets
            where us.IsActive
            select (uint)us.Position - 1
        ).FirstOrDefault();
        bool foundVisible = false;

        int totalSheets = sheetElements.Count() + xlWorkbook.UnsupportedSheets.Count;
        for (int p = 1; p <= totalSheets; p++)
        {
            if (xlWorkbook.UnsupportedSheets.All(us => us.Position != p))
            {
                Sheet sheet = sheetElements.ElementAt(
                    p - xlWorkbook.UnsupportedSheets.Count(us => us.Position <= p) - 1
                );
                workbook.Sheets.RemoveChild(sheet);
                workbook.Sheets.AppendChild(sheet);
                IXLWorksheet xlSheet = xlWorkbook.Worksheet(sheet.Name);
                if (xlSheet.Visibility != XLWorksheetVisibility.Visible)
                {
                    sheet.State = xlSheet.Visibility.ToOpenXml();
                }
                else
                {
                    sheet.State = null;
                }

                if (foundVisible)
                {
                    continue;
                }

                if (sheet.State == null || sheet.State == SheetStateValues.Visible)
                {
                    foundVisible = true;
                }
                else
                {
                    firstSheetVisible++;
                }
            }
            else
            {
                uint sheetId = xlWorkbook.UnsupportedSheets.First(us => us.Position == p).SheetId;
                Sheet sheet = workbook.Sheets.Elements<Sheet>().First(s => s.SheetId == sheetId);
                workbook.Sheets.RemoveChild(sheet);
                workbook.Sheets.AppendChild(sheet);
            }
        }

        WorkbookView workbookView = workbook.BookViews.Elements<WorkbookView>().FirstOrDefault();

        if (activeTab == 0)
        {
            uint? firstActiveTab = null;
            uint? firstSelectedTab = null;
            foreach (XLWorksheet ws in worksheets)
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

        if (workbookView == null)
        {
            workbookView = new WorkbookView
            {
                ActiveTab = activeTab,
                FirstSheet = firstSheetVisible,
            };
            workbook.BookViews.AppendChild(workbookView);
        }
        else
        {
            workbookView.ActiveTab = activeTab;
            workbookView.FirstSheet = firstSheetVisible;
        }

        DefinedNames definedNames = new();
        foreach (XLWorksheet worksheet in xlWorkbook.WorksheetsInternal)
        {
            uint wsSheetId = worksheet.SheetId;
            uint sheetId = 0;
            foreach (
                Sheet s in workbook.Sheets.Elements<Sheet>().TakeWhile(s => s.SheetId != wsSheetId)
            )
            {
                sheetId++;
            }

            if (worksheet.PageSetup.PrintAreas.Any())
            {
                DefinedName definedName = new()
                {
                    Name = "_xlnm.Print_Area",
                    LocalSheetId = sheetId,
                };
                string worksheetName = worksheet.Name;
                string definedNameText = worksheet.PageSetup.PrintAreas.Aggregate(
                    string.Empty,
                    (current, printArea) =>
                        current
                        + (
                            worksheetName.EscapeSheetName()
                            + "!"
                            + printArea.RangeAddress.FirstAddress.ToStringFixed(XLReferenceStyle.A1)
                            + ":"
                            + printArea.RangeAddress.LastAddress.ToStringFixed(XLReferenceStyle.A1)
                            + ","
                        )
                );
                definedName.Text = definedNameText.Substring(0, definedNameText.Length - 1);
                definedNames.AppendChild(definedName);
            }

            if (worksheet.AutoFilter.IsEnabled)
            {
                DefinedName definedName = new()
                {
                    Name = "_xlnm._FilterDatabase",
                    LocalSheetId = sheetId,
                    Text =
                        worksheet.Name.EscapeSheetName()
                        + "!"
                        + worksheet.AutoFilter.Range.RangeAddress.FirstAddress.ToStringFixed(
                            XLReferenceStyle.A1
                        )
                        + ":"
                        + worksheet.AutoFilter.Range.RangeAddress.LastAddress.ToStringFixed(
                            XLReferenceStyle.A1
                        ),
                    Hidden = BooleanValue.FromBoolean(true),
                };
                definedNames.AppendChild(definedName);
            }

            foreach (
                XLDefinedName xlDefinedName in worksheet.DefinedNames.Where<XLDefinedName>(n =>
                    n.Name != "_xlnm._FilterDatabase"
                )
            )
            {
                DefinedName definedName = new()
                {
                    Name = xlDefinedName.Name,
                    LocalSheetId = sheetId,
                    Text = xlDefinedName.ToString(),
                };

                if (!xlDefinedName.Visible)
                {
                    definedName.Hidden = BooleanValue.FromBoolean(true);
                }

                if (!string.IsNullOrWhiteSpace(xlDefinedName.Comment))
                {
                    definedName.Comment = xlDefinedName.Comment;
                }

                definedNames.AppendChild(definedName);
            }

            string definedNameTextRow = string.Empty;
            string definedNameTextColumn = string.Empty;
            if (worksheet.PageSetup.FirstRowToRepeatAtTop > 0)
            {
                definedNameTextRow =
                    worksheet.Name.EscapeSheetName()
                    + "!"
                    + worksheet.PageSetup.FirstRowToRepeatAtTop
                    + ":"
                    + worksheet.PageSetup.LastRowToRepeatAtTop;
            }
            if (worksheet.PageSetup.FirstColumnToRepeatAtLeft > 0)
            {
                int minColumn = worksheet.PageSetup.FirstColumnToRepeatAtLeft;
                int maxColumn = worksheet.PageSetup.LastColumnToRepeatAtLeft;
                definedNameTextColumn =
                    worksheet.Name.EscapeSheetName()
                    + "!"
                    + XlsxSharp.XLHelper.GetColumnLetterFromNumber(minColumn)
                    + ":"
                    + XlsxSharp.XLHelper.GetColumnLetterFromNumber(maxColumn);
            }

            string titles;
            if (definedNameTextColumn.Length > 0)
            {
                titles = definedNameTextColumn;
                if (definedNameTextRow.Length > 0)
                {
                    titles += "," + definedNameTextRow;
                }
            }
            else
            {
                titles = definedNameTextRow;
            }

            if (titles.Length <= 0)
            {
                continue;
            }

            DefinedName definedName2 = new()
            {
                Name = "_xlnm.Print_Titles",
                LocalSheetId = sheetId,
                Text = titles,
            };

            definedNames.AppendChild(definedName2);
        }

        foreach (XLDefinedName xlDefinedName in xlWorkbook.DefinedNamesInternal)
        {
            DefinedName definedName = new()
            {
                Name = xlDefinedName.Name,
                Text = xlDefinedName.RefersTo,
            };

            if (!xlDefinedName.Visible)
            {
                definedName.Hidden = BooleanValue.FromBoolean(true);
            }

            if (!string.IsNullOrWhiteSpace(xlDefinedName.Comment))
            {
                definedName.Comment = xlDefinedName.Comment;
            }

            definedNames.AppendChild(definedName);
        }

        workbook.DefinedNames = definedNames;

        if (workbook.CalculationProperties == null)
        {
            workbook.CalculationProperties = new CalculationProperties { CalculationId = 125725U };
        }

        if (xlWorkbook.CalculateMode == XLCalculateMode.Default)
        {
            workbook.CalculationProperties.CalculationMode = null;
        }
        else
        {
            workbook.CalculationProperties.CalculationMode = xlWorkbook.CalculateMode.ToOpenXml();
        }

        if (xlWorkbook.ReferenceStyle == XLReferenceStyle.Default)
        {
            workbook.CalculationProperties.ReferenceMode = null;
        }
        else
        {
            workbook.CalculationProperties.ReferenceMode = xlWorkbook.ReferenceStyle.ToOpenXml();
        }

        if (xlWorkbook.CalculationOnSave)
        {
            workbook.CalculationProperties.CalculationOnSave = xlWorkbook.CalculationOnSave;
        }

        if (xlWorkbook.ForceFullCalculation)
        {
            workbook.CalculationProperties.ForceFullCalculation = xlWorkbook.ForceFullCalculation;
        }

        if (xlWorkbook.FullCalculationOnLoad)
        {
            workbook.CalculationProperties.FullCalculationOnLoad = xlWorkbook.FullCalculationOnLoad;
        }

        if (xlWorkbook.FullPrecision)
        {
            workbook.CalculationProperties.FullPrecision = xlWorkbook.FullPrecision;
        }
    }
}

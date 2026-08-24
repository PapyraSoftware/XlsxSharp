using System.Diagnostics;
using XlsxSharp.IO.CodeGen.Model;
using XlsxSharp.IO.CodeGen.XsdParser;

namespace XlsxSharp.IO.CodeGen;

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.Error.WriteLine("Usage:");
            Console.Error.WriteLine(
                $"    {Process.GetCurrentProcess().ProcessName}.exe command name-of-ooxml.xsd output.cs"
            );
            Console.Error.WriteLine();
            return;
        }

        string command = args[0];
        string schemaPath = args[1];
        string target = args[2];
        using FileStream fileStream = File.OpenRead(schemaPath);
        using XmlTreeReader reader = new(fileStream, new XsdEnumMapper(), true);
        XsdSchemaParser parser = new();

        Schema schema = parser.ParseSchema(reader);
        switch (command)
        {
            case "styles":
                GenerateStylesReader(schema, target);
                break;

            case "cache-records":
                GenerateCacheRecords(schema, target);
                break;

            default:
                Console.WriteLine($"Unknown command '{command}'");
                break;
        }
    }

    private static void GenerateStylesReader(Schema schema, string target)
    {
        SchemeTypeMap typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes()
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_NumFmtId",
                    CsTypeName = "uint",
                    RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                    OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")",
                }
            )
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_FontId",
                    CsTypeName = "uint",
                    RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                    OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")",
                }
            )
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_FillId",
                    CsTypeName = "uint",
                    RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                    OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")",
                }
            )
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_BorderId",
                    CsTypeName = "uint",
                    RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                    OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")",
                }
            )
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_CellStyleXfId",
                    CsTypeName = "uint",
                    RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                    OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")",
                }
            )
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_TextRotation",
                    CsTypeName = "uint",
                    RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                    OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")",
                }
            )
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_DxfId",
                    CsTypeName = "uint",
                    RequiredTemplate = "_reader.GetUInt(\"{0}\")",
                    OptionalTemplate = "_reader.GetOptionalUInt(\"{0}\")",
                }
            )
            .AddSimpleTypeEnum("ST_PatternType", "XLFillPatternValues")
            .AddSimpleTypeEnum(
                "ST_GradientType",
                "XLGradientType",
                "linear",
                "XLGradientType.Linear"
            )
            .AddSimpleTypeEnum(
                "ST_BorderStyle",
                "XLBorderStyleValues",
                "none",
                "XLBorderStyleValues.None"
            )
            .AddSimpleTypeEnum("ST_HorizontalAlignment", "XLAlignmentHorizontalValues")
            .AddSimpleTypeEnum(
                "ST_VerticalAlignment",
                "XLAlignmentVerticalValues",
                "bottom",
                "XLAlignmentVerticalValues.Bottom"
            )
            .AddSimpleType(
                new SimpleTypeMapping
                {
                    Name = "ST_TableStyleType",
                    CsTypeName = "(XLTableStyleRegionValues?, XLPivotStyleRegionValues?)",
                    RequiredTemplate = "_reader.GetStringMappedValue(\"{0}\", TableStyleTypeMap)",
                }
            )
            .AddComplexTypeMapping("CT_Color", "XLColor")
            .AddComplexTypeMapping("CT_GradientStop", "(FractionOfOne Value, XLColor Color)")
            .AddComplexTypeMapping("CT_Font", "XLDifferentialFontValue")
            .AddComplexTypeMapping("CT_Fill", "XLFillFormatValue")
            .AddComplexTypeMapping("CT_Border", "XLDifferentialBorderValue")
            .AddComplexTypeMapping("CT_BorderPr", "XLBorderLine")
            .AddComplexTypeMapping("CT_PatternFill", "XLFillFormatValue")
            .AddComplexTypeMapping("CT_GradientFill", "XLFillFormatValue")
            .AddComplexTypeMapping("CT_NumFmt", "(int NumFmtId, XLNumberFormat Format)")
            .AddComplexTypeMapping("CT_CellAlignment", "XLDifferentialAlignmentValue")
            .AddComplexTypeMapping("CT_CellProtection", "XLDifferentialProtectionValue")
            .AddComplexTypeMapping("CT_Xf", "(XLCellFormatValue Format, int? CellStyleXfId)")
            .AddComplexTypeMapping(
                "CT_CellXfs",
                "List<(XLCellFormatValue Format, int? CellStyleXfId)>"
            )
            .AddComplexTypeMapping("CT_CellStyle", "(int CellStyleXfId, XLCellStyleValue Style)")
            .AddComplexTypeMapping("CT_CellStyles", "Dictionary<int, XLCellStyleValue>")
            .AddComplexTypeMapping("CT_RgbColor", "uint");

        ParserGenerator stylesReaderGenerator = new ParserGenerator(schema, typeMap, "StylesReader")
            .AddUsing("System.Collections.Generic")
            .AddUsing("XlsxSharp.IO")
            .AddUsing("XlsxSharp.Excel.Formatting")
            //.AddParseMethod("CT_Stylesheet")
            .AddParseMethod("CT_NumFmts")
            .AddParseMethod("CT_NumFmt")
            .AddParseMethod("CT_Fonts")
            // AddParseMethod("CT_Font")
            .AddParseMethod("CT_Fills")
            .AddParseMethod("CT_Fill")
            .AddParseMethod("CT_PatternFill")
            .AddParseMethod("CT_GradientFill")
            .AddParseMethod("CT_GradientStop")
            .AddParseMethod("CT_Borders")
            .AddParseMethod("CT_Border")
            .AddParseMethod("CT_BorderPr")
            .AddParseMethod("CT_CellStyleXfs")
            .AddParseMethod("CT_Xf")
            .AddParseMethod("CT_CellAlignment")
            .AddParseMethod("CT_CellProtection")
            .AddParseMethod("CT_CellXfs")
            .AddParseMethod("CT_CellStyles")
            .AddParseMethod("CT_CellStyle")
            .AddParseMethod("CT_Dxfs")
            .AddParseMethod("CT_Dxf")
            .AddParseMethod("CT_TableStyles")
            .AddParseMethod("CT_TableStyle")
            .AddParseMethod("CT_TableStyleElement")
            .AddParseMethod("CT_Colors")
            .AddParseMethod("CT_IndexedColors")
            .AddParseMethod("CT_MRUColors")
            .AddParseMethod("CT_RgbColor");

        string stylesReaderSource = stylesReaderGenerator.Generate();
        File.WriteAllText(target, stylesReaderSource);
        Console.WriteLine(stylesReaderSource);
    }

    private static void GenerateCacheRecords(Schema schema, string target)
    {
        SchemeTypeMap typeMap = new SchemeTypeMap().AddPrimitiveTypes();

        ParserGenerator cacheRecordsGenerator = new ParserGenerator(
            schema,
            typeMap,
            "PivotCacheRecordsReader"
        )
            .WithNamespace("XlsxSharp.Excel.IO")
            .AddUsing("System.Collections.Generic")
            .AddUsing("XlsxSharp.IO")
            // CT_PivotCacheRecords - hand-coded
            .AddParseMethod("CT_Record")
            .AddParseMethod("CT_Missing")
            .AddParseMethod("CT_Number")
            .AddParseMethod("CT_Boolean")
            .AddParseMethod("CT_Error")
            .AddParseMethod("CT_String")
            .AddParseMethod("CT_DateTime")
            .AddParseMethod("CT_Index")
            .AddParseMethod("CT_X")
            .AddParseMethod("CT_Tuples")
            .AddParseMethod("CT_Tuple");

        string cacheRecordsSource = cacheRecordsGenerator.Generate();
        File.WriteAllText(target, cacheRecordsSource);
        Console.WriteLine(cacheRecordsSource);
    }
}

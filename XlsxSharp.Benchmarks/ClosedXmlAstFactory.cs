using ClosedXML.Parser;

namespace XlsxSharp.Benchmarks;

/// <summary>
/// Builds <see cref="BenchNode"/>s for the real, upstream <c>ClosedXML.Parser</c> NuGet package.
/// Mirrors <see cref="XlsxSharpAstFactory"/> method for method, since the two <c>IAstFactory</c>
/// interfaces are identical in shape (<see cref="XlsxSharp.Parser"/> is a vendored fork of this
/// package) - only the parsing work behind each factory call differs.
/// </summary>
internal sealed class ClosedXmlAstFactory : IAstFactory<BenchNode, BenchNode, object?>
{
    public BenchNode LogicalValue(object? context, SymbolRange range, bool value) => new("LogicalValue");

    public BenchNode NumberValue(object? context, SymbolRange range, double value) => new("NumberValue");

    public BenchNode TextValue(object? context, SymbolRange range, string text) => new("TextValue");

    public BenchNode ErrorValue(object? context, SymbolRange range, ReadOnlySpan<char> error) => new("ErrorValue");

    public BenchNode ArrayNode(object? context, SymbolRange range, int rows, int columns, IReadOnlyList<BenchNode> elements) => new("Array");

    public BenchNode BlankNode(object? context, SymbolRange range) => new("Blank");

    public BenchNode LogicalNode(object? context, SymbolRange range, bool value) => new("Logical");

    public BenchNode ErrorNode(object? context, SymbolRange range, ReadOnlySpan<char> error) => new("Error");

    public BenchNode NumberNode(object? context, SymbolRange range, double value) => new("Number");

    public BenchNode TextNode(object? context, SymbolRange range, string text) => new("Text");

    public BenchNode Reference(object? context, SymbolRange range, ReferenceArea reference) => new("Reference");

    public BenchNode SheetReference(object? context, SymbolRange range, string sheet, ReferenceArea reference) => new("SheetReference");

    public BenchNode BangReference(object? context, SymbolRange range, ReferenceArea reference) => new("BangReference");

    public BenchNode Reference3D(object? context, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference) => new("Reference3D");

    public BenchNode ExternalSheetReference(object? context, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference) => new("ExternalSheetReference");

    public BenchNode ExternalReference3D(object? context, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference) => new("ExternalReference3D");

    public BenchNode Function(object? context, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<BenchNode> arguments) => new("Function");

    public BenchNode Function(object? context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<BenchNode> args) => new("SheetFunction");

    public BenchNode ExternalFunction(object? context, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<BenchNode> arguments) => new("ExternalSheetFunction");

    public BenchNode ExternalFunction(object? context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<BenchNode> arguments) => new("ExternalFunction");

    public BenchNode CellFunction(object? context, SymbolRange range, RowCol cell, IReadOnlyList<BenchNode> arguments) => new("CellFunction");

    public BenchNode StructureReference(object? context, SymbolRange range, StructuredReferenceArea area, string? firstColumn, string? lastColumn) => new("StructureReference");

    public BenchNode StructureReference(object? context, SymbolRange range, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn) => new("TableStructureReference");

    public BenchNode ExternalStructureReference(object? context, SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn) => new("ExternalStructureReference");

    public BenchNode Name(object? context, SymbolRange range, string name) => new("Name");

    public BenchNode SheetName(object? context, SymbolRange range, string sheet, string name) => new("SheetName");

    public BenchNode BangName(object? context, SymbolRange range, string name) => new("BangName");

    public BenchNode ExternalName(object? context, SymbolRange range, int workbookIndex, string name) => new("ExternalName");

    public BenchNode ExternalSheetName(object? context, SymbolRange range, int workbookIndex, string sheet, string name) => new("ExternalSheetName");

    public BenchNode BinaryNode(object? context, SymbolRange range, BinaryOperation operation, BenchNode leftNode, BenchNode rightNode) => new("Binary");

    public BenchNode Unary(object? context, SymbolRange range, UnaryOperation operation, BenchNode node) => new("Unary");

    public BenchNode Nested(object? context, SymbolRange range, BenchNode node) => node;
}

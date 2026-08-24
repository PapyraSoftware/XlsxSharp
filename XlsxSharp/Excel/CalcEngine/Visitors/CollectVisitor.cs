using XlsxSharp.Parser;

namespace XlsxSharp.Excel.CalcEngine.Visitors;

internal abstract class CollectVisitor<TContext> : IAstFactory<object?, object?, TContext>
{
    public virtual object? LogicalValue(TContext context, SymbolRange range, bool value) => default;

    public virtual object? NumberValue(TContext context, SymbolRange range, double value) =>
        default;

    public virtual object? TextValue(TContext context, SymbolRange range, string text) => default;

    public virtual object? ErrorValue(
        TContext context,
        SymbolRange range,
        ReadOnlySpan<char> error
    ) => default;

    public virtual object? ArrayNode(
        TContext context,
        SymbolRange range,
        int rows,
        int columns,
        IReadOnlyList<object?> elements
    ) => default;

    public virtual object? BlankNode(TContext context, SymbolRange range) => default;

    public virtual object? LogicalNode(TContext context, SymbolRange range, bool value) => default;

    public virtual object? ErrorNode(
        TContext context,
        SymbolRange range,
        ReadOnlySpan<char> error
    ) => default;

    public virtual object? NumberNode(TContext context, SymbolRange range, double value) => default;

    public virtual object? TextNode(TContext context, SymbolRange range, string text) => default;

    public virtual object? Reference(
        TContext context,
        SymbolRange range,
        ReferenceArea reference
    ) => default;

    public virtual object? SheetReference(
        TContext context,
        SymbolRange range,
        string sheet,
        ReferenceArea reference
    ) => default;

    public virtual object? BangReference(
        TContext context,
        SymbolRange range,
        ReferenceArea reference
    ) => default;

    public virtual object? Reference3D(
        TContext context,
        SymbolRange range,
        string firstSheet,
        string lastSheet,
        ReferenceArea reference
    ) => default;

    public virtual object? ExternalSheetReference(
        TContext context,
        SymbolRange range,
        int workbookIndex,
        string sheet,
        ReferenceArea reference
    ) => default;

    public virtual object? ExternalReference3D(
        TContext context,
        SymbolRange range,
        int workbookIndex,
        string firstSheet,
        string lastSheet,
        ReferenceArea reference
    ) => default;

    public virtual object? Function(
        TContext context,
        SymbolRange range,
        ReadOnlySpan<char> functionName,
        IReadOnlyList<object?> arguments
    ) => default;

    public virtual object? Function(
        TContext context,
        SymbolRange range,
        string sheetName,
        ReadOnlySpan<char> functionName,
        IReadOnlyList<object?> args
    ) => default;

    public virtual object? ExternalFunction(
        TContext context,
        SymbolRange range,
        int workbookIndex,
        string sheetName,
        ReadOnlySpan<char> functionName,
        IReadOnlyList<object?> arguments
    ) => default;

    public virtual object? ExternalFunction(
        TContext context,
        SymbolRange range,
        int workbookIndex,
        ReadOnlySpan<char> functionName,
        IReadOnlyList<object?> arguments
    ) => default;

    public virtual object? CellFunction(
        TContext context,
        SymbolRange range,
        RowCol cell,
        IReadOnlyList<object?> arguments
    ) => default;

    public virtual object? StructureReference(
        TContext context,
        SymbolRange range,
        StructuredReferenceArea area,
        string? firstColumn,
        string? lastColumn
    ) => default;

    public virtual object? StructureReference(
        TContext context,
        SymbolRange range,
        string table,
        StructuredReferenceArea area,
        string? firstColumn,
        string? lastColumn
    ) => default;

    public virtual object? ExternalStructureReference(
        TContext context,
        SymbolRange range,
        int workbookIndex,
        string table,
        StructuredReferenceArea area,
        string? firstColumn,
        string? lastColumn
    ) => default;

    public virtual object? Name(TContext context, SymbolRange range, string name) => default;

    public virtual object? SheetName(
        TContext context,
        SymbolRange range,
        string sheet,
        string name
    ) => default;

    public virtual object? BangName(TContext context, SymbolRange range, string name) => default;

    public virtual object? ExternalName(
        TContext context,
        SymbolRange range,
        int workbookIndex,
        string name
    ) => default;

    public virtual object? ExternalSheetName(
        TContext context,
        SymbolRange range,
        int workbookIndex,
        string sheet,
        string name
    ) => default;

    public virtual object? BinaryNode(
        TContext context,
        SymbolRange range,
        BinaryOperation operation,
        object? leftNode,
        object? rightNode
    ) => default;

    public virtual object? Unary(
        TContext context,
        SymbolRange range,
        UnaryOperation operation,
        object? node
    ) => default;

    public virtual object? Nested(TContext context, SymbolRange range, object? node) => default;
}

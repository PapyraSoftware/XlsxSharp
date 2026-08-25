using System.Text;
using XlsxSharp.Parser.Pratt.Parselets;

namespace XlsxSharp.Parser;

/// <summary>
/// A visitor that generates the identical formula for the parsed formula based on passed arguments.
/// CopyVisitor doesn't make any judgements if passed arguments have been modified. It just makes
/// a newly allocated copy based on passed values.
/// </summary>
public class CopyVisitor : IAstFactory<TransformedSymbol, TransformedSymbol, ModContext>
{
    // 1 quote on left, 1 quote on right size and at most 4 quotes inside.
    private const int QUOTE_RESERVE = 6;
    private const int SHEET_SEPARATOR_LEN = 1;
    private const int BOOK_PREFIX_LEN = 3;
    private const int MAX_R1_C1_LEN = 20;

    /// <inheritdoc />
    public virtual TransformedSymbol LogicalValue(ModContext ctx, SymbolRange range, bool value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol NumberValue(ModContext ctx, SymbolRange range, double value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol TextValue(ModContext ctx, SymbolRange range, string text)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ErrorValue(ModContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ArrayNode(ModContext ctx, SymbolRange range, int rows, int columns, IReadOnlyList<TransformedSymbol> elements)
    {
        StringBuilder sb = new(2 + elements.Sum(x => x.Length) + elements.Count);
        sb.AppendStartFragment(ctx, range, elements[0]);
        int i = 0;
        sb.Append(elements[i++].AsSpan());
        for (int col = 1; col < columns; ++col)
        {
            sb.AppendMiddleFragment(ctx, elements[i - 1], elements[i]);
            sb.Append(elements[i++].AsSpan());
        }

        for (int row = 1; row < rows; ++row)
        {
            sb.AppendMiddleFragment(ctx, elements[i - 1], elements[i]);
            sb.Append(elements[i++].AsSpan());
            for (int col = 1; col < columns; ++col)
            {
                sb.AppendMiddleFragment(ctx, elements[i - 1], elements[i]);
                sb.Append(elements[i++].AsSpan());
            }
        }

        sb.AppendEndFragment(ctx, range, elements[elements.Count - 1]);
        string nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol BlankNode(ModContext ctx, SymbolRange range)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol LogicalNode(ModContext ctx, SymbolRange range, bool value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ErrorNode(ModContext ctx, SymbolRange range, ReadOnlySpan<char> error)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol NumberNode(ModContext ctx, SymbolRange range, double value)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol TextNode(ModContext ctx, SymbolRange range, string text)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        StringBuilder sb = new(MAX_R1_C1_LEN);
        string nodeText = sb.AppendRef(reference).ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        StringBuilder sb = new(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        string nodeText = sb
            .AppendSheetReference(sheet)
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol BangReference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        StringBuilder sb = new(SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        string nodeText = sb
            .AppendReferenceSeparator()
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Reference3D(ModContext ctx, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        StringBuilder sb = new(firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        if (NameUtils.ShouldQuote(firstSheet.AsSpan()) || NameUtils.ShouldQuote(lastSheet.AsSpan()))
        {
            sb
                .Append('\'')
                .AppendEscapedSheetName(firstSheet)
                .Append(':')
                .AppendEscapedSheetName(lastSheet)
                .Append('\'');
        }
        else
        {
            sb.Append(firstSheet)
                .Append(':')
                .Append(lastSheet);
        }

        string nodeText = sb
            .AppendReferenceSeparator()
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalSheetReference(ModContext ctx, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference)
    {
        StringBuilder sb = new(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        string nodeText = sb
            .AppendExternalSheetReference(workbookIndex, sheet)
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalReference3D(ModContext ctx, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference)
    {
        StringBuilder sb = new(BOOK_PREFIX_LEN + firstSheet.Length + QUOTE_RESERVE + lastSheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + MAX_R1_C1_LEN);
        if (NameUtils.ShouldQuote(firstSheet.AsSpan()) || NameUtils.ShouldQuote(lastSheet.AsSpan()))
        {
            sb
                .Append('\'')
                .AppendBookIndex(workbookIndex)
                .AppendEscapedSheetName(firstSheet)
                .Append(':')
                .AppendEscapedSheetName(lastSheet)
                .Append('\'');
        }
        else
        {
            sb
                .AppendBookIndex(workbookIndex)
                .Append(firstSheet)
                .Append(':')
                .Append(lastSheet);
        }

        string nodeText = sb
            .AppendReferenceSeparator()
            .AppendRef(reference)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Function(ModContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        StringBuilder sb = new(functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        string nodeText = sb.Append(functionName).AppendArguments(ctx, range, arguments).ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalFunction(ModContext ctx, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        StringBuilder sb = new(BOOK_PREFIX_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        string nodeText = sb
            .AppendBookIndex(workbookIndex)
            .AppendReferenceSeparator()
            .AppendFunction(ctx, range, functionName, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Function(ModContext ctx, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        StringBuilder sb = new(sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        string nodeText = sb
            .AppendSheetReference(sheetName)
            .AppendFunction(ctx, range, functionName, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalFunction(ModContext ctx, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        StringBuilder sb = new(BOOK_PREFIX_LEN + sheetName.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + functionName.Length + 2 + arguments.Sum(static x => x.Length) + arguments.Count);
        string nodeText = sb
            .AppendExternalSheetReference(workbookIndex, sheetName)
            .AppendFunction(ctx, range, functionName, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol CellFunction(ModContext ctx, SymbolRange range, RowCol cell, IReadOnlyList<TransformedSymbol> arguments)
    {
        StringBuilder sb = new(MAX_R1_C1_LEN + SHEET_SEPARATOR_LEN + arguments.Sum(static x => x.Length));
        string nodeText = sb
            .AppendRef(cell)
            .AppendArguments(ctx, range, arguments)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol StructureReference(ModContext ctx, SymbolRange range,
        StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        string nodeText = GetIntraTableReference(area, firstColumn, lastColumn);
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol StructureReference(ModContext ctx, SymbolRange range,
        string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        string nodeText = table + GetIntraTableReference(area, firstColumn, lastColumn);
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalStructureReference(ModContext ctx,
        SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn,
        string? lastColumn)
    {
        string nodeText = new StringBuilder()
            .AppendBookIndex(workbookIndex).Append(table)
            .Append(GetIntraTableReference(area, firstColumn, lastColumn))
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Name(ModContext ctx, SymbolRange range, string name)
    {
        return TransformedSymbol.CopyOriginal(ctx.Formula, range);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol SheetName(ModContext ctx, SymbolRange range, string sheet, string name)
    {
        StringBuilder sb = new(sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        string nodeText = sb
            .AppendSheetReference(sheet)
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol BangName(ModContext ctx, SymbolRange range, string name)
    {
        StringBuilder sb = new(SHEET_SEPARATOR_LEN + name.Length);
        string nodeText = sb
            .AppendReferenceSeparator()
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalName(ModContext ctx, SymbolRange range, int workbookIndex, string name)
    {
        StringBuilder sb = new(BOOK_PREFIX_LEN + SHEET_SEPARATOR_LEN + name.Length + QUOTE_RESERVE);
        sb.AppendBookIndex(workbookIndex).AppendReferenceSeparator();
        string nodeText = (
            RequiresQuotingAsExternalName(name.AsSpan())
                ? sb.Append('\'').AppendEscapedSheetName(name).Append('\'')
                : sb.Append(name)
        ).ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <summary>
    /// Does an external defined name need to be re-quoted when rendered unqualified (bare
    /// <c>[n]!name</c>, no sheet)? <see cref="NameUtils.ShouldQuote"/> can't be reused here: that
    /// bitmask encodes Excel's *sheet name* quoting rules, which allow characters (e.g. <c>/</c>)
    /// that NAME's own grammar (<see cref="NameUtils.IsNameValid"/>) doesn't - rendering those
    /// unquoted wouldn't throw, it would silently reparse as something else (e.g. <c>N/A</c> as a
    /// division). <see cref="NameUtils.IsNameValid"/> alone isn't enough either: it accepts
    /// anything shaped like a plain identifier, including one that also happens to look like a
    /// cell reference (A1 or R1C1 style) or <c>TRUE</c>/<c>FALSE</c> - the same shapes
    /// <see cref="Parselets.StructureReferenceParselet{TScalar,T,TContext}"/> excludes from its
    /// unquoted branch (see its <c>ParseExternalReference</c>). Checked against both A1 and R1C1
    /// cell shapes regardless of which style is being rendered, since over-quoting is harmless but
    /// under-quoting produces text that fails to reparse (or reparses into something else).
    /// </summary>
    private static bool RequiresQuotingAsExternalName(ReadOnlySpan<char> name)
    {
        if (!NameUtils.IsNameValid(name) || ParserExtensions.TryGetCellA1(name, out _))
        {
            return true;
        }

        if (name.Length > 0 && name[0] is 'R' or 'r' or 'C' or 'c')
        {
            int i = 0;
            TokenParser.ParseR1C1Reference(name, ref i);
            if (i == name.Length)
            {
                return true;
            }
        }

        return ParserExtensions.EqualCaseInsensitive(name, "TRUE") || ParserExtensions.EqualCaseInsensitive(name, "FALSE");
    }

    /// <inheritdoc />
    public virtual TransformedSymbol ExternalSheetName(ModContext ctx, SymbolRange range, int workbookIndex, string sheet, string name)
    {
        StringBuilder sb = new(BOOK_PREFIX_LEN + sheet.Length + QUOTE_RESERVE + SHEET_SEPARATOR_LEN + name.Length);
        string nodeText = sb
            .AppendExternalSheetReference(workbookIndex, sheet)
            .Append(name)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol BinaryNode(ModContext ctx, SymbolRange range, BinaryOperation operation, TransformedSymbol leftNode, TransformedSymbol rightNode)
    {
        StringBuilder sb = new StringBuilder(leftNode.Length + rightNode.OriginalRange.Start - leftNode.OriginalRange.End + rightNode.Length)
            .AppendStartFragment(ctx, range, leftNode)
            .Append(leftNode.AsSpan())
            .AppendMiddleFragment(ctx, leftNode, rightNode)
            .Append(rightNode.AsSpan())
            .AppendEndFragment(ctx, range, rightNode);

        string nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Unary(ModContext ctx, SymbolRange range, UnaryOperation operation, TransformedSymbol node)
    {
        StringBuilder sb = new StringBuilder(node.Length + 1)
            .AppendStartFragment(ctx, range, node)
            .Append(node.AsSpan())
            .AppendEndFragment(ctx, range, node);

        string nodeText = sb.ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    /// <inheritdoc />
    public virtual TransformedSymbol Nested(ModContext ctx, SymbolRange range, TransformedSymbol node)
    {
        string nodeText = new StringBuilder(node.Length + 2)
            .AppendStartFragment(ctx, range, node)
            .Append(node.AsSpan())
            .AppendEndFragment(ctx, range, node)
            .ToString();
        return TransformedSymbol.ToText(ctx.Formula, range, nodeText);
    }

    private static string GetIntraTableReference(StructuredReferenceArea area, string? firstColumn, string? lastColumn)
    {
        if (firstColumn is null || lastColumn is null)
        {
            // No column

            // Shorthand for full table inside the table.
            if (area == StructuredReferenceArea.None)
            {
                return "[]";
            }

            if (area == (StructuredReferenceArea.Headers | StructuredReferenceArea.Data))
            {
                return "[[#Headers],[#Data]]";
            }

            if (area == (StructuredReferenceArea.Data | StructuredReferenceArea.Totals))
            {
                return "[[#Data],[#Totals]]";
            }

            return Keyword(area);
        }

        if (firstColumn == lastColumn)
        {
            // One column
            if (area == StructuredReferenceArea.None)
            {
                // One column, no keyword
                return new StringBuilder(firstColumn.Length + 2)
                    .Append('[').Append(firstColumn).Append(']')
                    .ToString();
            }

            // One column, keyword
            string keywordList = KeywordList(area);
            return new StringBuilder(keywordList.Length + firstColumn.Length + 5)
                .Append('[')
                .Append(keywordList).Append(',')
                .Append('[').Append(firstColumn).Append(']')
                .Append(']')
                .ToString();
        }
        else
        {
            // Two columns
            string keywordList = KeywordList(area);
            StringBuilder sb = new(firstColumn.Length + lastColumn.Length + keywordList.Length + 8);
            sb.Append('[');
            if (keywordList.Length > 0)
            {
                sb.Append(keywordList).Append(',');
            }

            return sb
                .Append('[').Append(firstColumn).Append(']')
                .Append(':')
                .Append('[').Append(lastColumn).Append(']')
                .Append(']')
                .ToString();
        }

        static string KeywordList(StructuredReferenceArea area)
        {
            return area switch
            {
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data => "[#Headers],[#Data]",
                StructuredReferenceArea.Data | StructuredReferenceArea.Totals => "[#Data],[#Totals]",
                _ => Keyword(area),
            };
        }

        static string Keyword(StructuredReferenceArea area)
        {
            return area switch
            {
                StructuredReferenceArea.None => string.Empty,
                StructuredReferenceArea.Headers => "[#Headers]",
                StructuredReferenceArea.Data => "[#Data]",
                StructuredReferenceArea.Totals => "[#Totals]",
                StructuredReferenceArea.All => "[#All]",
                StructuredReferenceArea.ThisRow => "[#This Row]",
                _ => throw new NotSupportedException(),
            };
        }
    }
}

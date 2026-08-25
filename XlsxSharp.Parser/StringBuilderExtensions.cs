using System.Text;

namespace XlsxSharp.Parser;

/// <summary>
/// Extension methods for building formulas.
/// </summary>
internal static class StringBuilderExtensions
{
    public static StringBuilder AppendSheetReference(this StringBuilder sb, string? sheetName)
    {
        if (sheetName is null)
        {
            return sb.Append("#REF!");
        }

        return NameUtils.EscapeName(sb, sheetName).AppendReferenceSeparator();
    }

    public static StringBuilder AppendExternalSheetReference(this StringBuilder sb, int workbookIndex, string sheetName)
    {
        if (NameUtils.ShouldQuote(sheetName.AsSpan()))
        {
            return sb
                .Append('\'')
                .AppendBookIndex(workbookIndex)
                .AppendEscapedSheetName(sheetName)
                .Append('\'')
                .AppendReferenceSeparator();
        }

        return sb
            .AppendBookIndex(workbookIndex)
            .AppendSheetReference(sheetName);
    }
    public static StringBuilder AppendEscapedSheetName(this StringBuilder sb, string sheetName)
    {
        int startIndex = sb.Length;
        return sb.Append(sheetName).Replace("'", "''", startIndex, sheetName.Length);
    }

    public static StringBuilder AppendReferenceSeparator(this StringBuilder sb)
    {
        return sb.Append('!');
    }

    public static StringBuilder AppendBookIndex(this StringBuilder sb, int bookIndex)
    {
        return sb.Append('[').Append(bookIndex).Append(']');
    }

    public static StringBuilder AppendFunction(this StringBuilder sb, ModContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
    {
        return sb.Append(functionName).AppendArguments(ctx, range, arguments);
    }

    public static StringBuilder AppendRef(this StringBuilder sb, ReferenceArea? reference)
    {
        return reference is null ? sb.Append("#REF!") : reference.Value.Append(sb);
    }

    public static StringBuilder AppendRef(this StringBuilder sb, RowCol? rowCol)
    {
        if (rowCol is null)
        {
            return sb.Append("#REF!");
        }

        rowCol.Value.Append(sb);
        return sb;
    }

    public static StringBuilder AppendStartFragment(this StringBuilder sb, ModContext ctx, SymbolRange symbolRange, TransformedSymbol nestedNode)
    {
        string formula = ctx.Formula;
        for (int i = symbolRange.Start; i < nestedNode.OriginalRange.Start; ++i)
        {
            sb.Append(formula[i]);
        }

        return sb;
    }

    public static StringBuilder AppendMiddleFragment(this StringBuilder sb, ModContext ctx, TransformedSymbol beforeNode, TransformedSymbol afterNode)
    {
        string formula = ctx.Formula;
        for (int i = beforeNode.OriginalRange.End; i < afterNode.OriginalRange.Start; ++i)
        {
            sb.Append(formula[i]);
        }

        return sb;
    }

    public static StringBuilder AppendEndFragment(this StringBuilder sb, ModContext ctx, SymbolRange symbolRange, TransformedSymbol nestedNode)
    {
        string formula = ctx.Formula;
        for (int i = nestedNode.OriginalRange.End; i < symbolRange.End; ++i)
        {
            sb.Append(formula[i]);
        }

        return sb;
    }

    public static StringBuilder AppendArguments(this StringBuilder sb, ModContext ctx, SymbolRange range, IReadOnlyList<TransformedSymbol> arguments)
    {
        if (arguments.Count == 0)
        {
            int braceIdx = GetStartBraceIndex(ctx, range, range.End);
            ReadOnlySpan<char> braces = ctx.Formula.AsSpan().Slice(braceIdx, range.End - braceIdx);
            sb.Append(braces);
        }
        else
        {
            sb
                .AppendStartBrace(ctx, range, arguments[0])
                .AppendArguments(ctx, arguments)
                .AppendEndFragment(ctx, range, arguments[arguments.Count - 1]);
        }

        return sb;
    }

    private static StringBuilder AppendStartBrace(this StringBuilder sb, ModContext ctx, SymbolRange range, TransformedSymbol firstNode)
    {
        int firstNodeStart = firstNode.OriginalRange.Start;
        int braceIdx = GetStartBraceIndex(ctx, range, firstNodeStart);
        for (int j = braceIdx; j < firstNodeStart; ++j)
        {
            sb.Append(ctx.Formula[j]);
        }

        return sb;
    }

    private static int GetStartBraceIndex(ModContext ctx, SymbolRange range, int nodeStart)
    {
        string formula = ctx.Formula;
        int braceIdx = nodeStart - 1;
        for (; braceIdx > range.Start; --braceIdx)
        {
            if (formula[braceIdx] == '(')
            {
                return braceIdx;
            }
        }

        throw new InvalidOperationException("No opening brace found.");
    }

    private static StringBuilder AppendArguments(this StringBuilder sb, ModContext ctx, IReadOnlyList<TransformedSymbol> arguments)
    {
        if (arguments.Count > 0)
        {
            sb.Append(arguments[0].AsSpan());
        }

        for (int i = 1; i < arguments.Count; ++i)
        {
            sb.AppendMiddleFragment(ctx, arguments[i - 1], arguments[i]);
            sb.Append(arguments[i].AsSpan());
        }

        return sb;
    }
}

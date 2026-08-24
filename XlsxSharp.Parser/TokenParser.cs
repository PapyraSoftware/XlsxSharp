using System.Diagnostics;
using static XlsxSharp.Parser.ReferenceAxisType;

namespace XlsxSharp.Parser;

internal static class TokenParser
{
    /// <summary>
    /// Parse <see cref="Token.SINGLE_SHEET_PREFIX"/> token.
    /// </summary>
    internal static void ParseSingleSheetPrefix(ReadOnlySpan<char> input, out int? index, out string sheetName)
    {
        // There can be whitespaces after exclamation mark at the end of a token.
        input = input.TrimEnd();
        bool isEscaped = input[0] == '\'';
        input = isEscaped
            ? input.Slice(1, input.Length - 3) // second sheet name ends with TICK EXCLAMATION_MARK ('!). 
            : input.Slice(0, input.Length - 1); // only strip exclamation mark

        // Parse optional WORKBOOK_INDEX
        input = ExtractWorkbookIndex(input, out index);

        ReadOnlySpan<char> sheetRangeSpan = input;
        if (!isEscaped)
        {
            // SHEET_NAME
            sheetName = sheetRangeSpan.ToString();
            return;
        }

        // The ending '! have been stripped from escape
        sheetName = GetEscapedSheetName(input);
    }

    /// <summary>
    /// Parse token <see cref="Token.SHEET_RANGE_PREFIX"/>
    /// </summary>
    internal static void ParseSheetRangePrefix(ReadOnlySpan<char> input, out int? index, out string firstSheetName, out string secondSheetName)
    {
        bool isEscaped = input[0] == '\'';
        input = isEscaped
            ? input.Slice(1, input.Length - 3) // second sheet name ends with TICK EXCLAMATION_MARK ('!). 
            : input.Slice(0, input.Length - 1); // only strip exclamation mark

        // Parse optional WORKBOOK_INDEX
        input = ExtractWorkbookIndex(input, out index);

        ReadOnlySpan<char> sheetRangeSpan = input;
        if (!isEscaped)
        {
            // SHEET_NAME ':' SHEET_NAME
            int endIndex = sheetRangeSpan.IndexOf(':');
            firstSheetName = sheetRangeSpan.Slice(0, endIndex).ToString();
            secondSheetName = sheetRangeSpan.Slice(endIndex + 1).ToString();
            return;
        }

        // Parse SHEET_NAME_SPECIAL which can contain escaped tick (') as double tick
        firstSheetName = GetEscapedSheetName(ref input, ':'); // Even escaped sheet name can't contain :
        secondSheetName = GetEscapedSheetName(input);
    }

    private static ReadOnlySpan<char> ExtractWorkbookIndex(ReadOnlySpan<char> input, out int? wbIndex)
    {
        if (input[0] != '[')
        {
            wbIndex = null;
            return input;
        }

        int i = 0;
        int number = 0;
        char c = input[++i];
        do
        {
            number = number * 10 + c - '0';
            c = input[++i];
        } while (c != ']');

        wbIndex = number;
        return input.Slice(i + 1);
    }

    private static string GetEscapedSheetName(ref ReadOnlySpan<char> input, char endChar)
    {
        Span<char> buffer = stackalloc char[input.Length];
        int bufferIdx = 0;
        int inputIdx = 0;
        do
        {
            if (input[inputIdx] == '\'')
            {
                inputIdx++;
            }

            buffer[bufferIdx++] = input[inputIdx++];
        } while (input[inputIdx] != endChar);

        input = input.Slice(inputIdx + 1);
        return buffer.Slice(0, bufferIdx).ToString();
    }

    private static string GetEscapedSheetName(ReadOnlySpan<char> input)
    {
        Span<char> buffer = stackalloc char[input.Length];
        int bufferIdx = 0;
        int inputIdx = 0;
        do
        {
            if (input[inputIdx] == '\'')
            {
                inputIdx++;
            }

            buffer[bufferIdx++] = input[inputIdx++];
        } while (input.Length > inputIdx);

        return buffer.Slice(0, bufferIdx).ToString();
    }

    internal static RowCol ExtractCellFunction(ReadOnlySpan<char> cellFunctionToken)
    {
        int i = 0;
        return ReadA1Cell(cellFunctionToken, ref i);
    }

    internal static ReadOnlySpan<char> ExtractLocalFunctionName(ReadOnlySpan<char> functionNameWithBrace)
    {
        // In most cases, there won't be any whitespace
        int endPosition = functionNameWithBrace[functionNameWithBrace.Length - 1] == '('
            ? functionNameWithBrace.Length - 1
            : functionNameWithBrace.LastIndexOf('(');
        ReadOnlySpan<char> functionName = functionNameWithBrace.Slice(0, endPosition);
        return functionName;
    }

    internal static ReferenceArea ParseReference(ReadOnlySpan<char> input, bool isA1)
    {
        return isA1 ? ParseA1Reference(input) : ParseR1C1Reference(input);
    }

    /// <summary>
    /// Parse <c>A1_REFERENCE</c> token in R1C1 mode.
    /// </summary>
    /// <param name="token">The span of a token.</param>
    private static ReferenceArea ParseR1C1Reference(ReadOnlySpan<char> token)
    {
        int i = 0;
        RowCol rowCol1 = ParseR1C1Reference(token, ref i);
        if (i == token.Length)
        {
            return new ReferenceArea(rowCol1, rowCol1);
        }

        if (token[i++] != ':')
        {
            throw Bug();
        }

        RowCol rowCol2 = ParseR1C1Reference(token, ref i);
        return new ReferenceArea(rowCol1, rowCol2);
    }

    private static RowCol ParseR1C1Reference(ReadOnlySpan<char> token, ref int i)
    {
        if (token[i] is 'C' or 'c')
        {
            // Token is COLUMN. ROW must be after column, so there can't be one.
            (ReferenceAxisType Type, int Value) loneCol = ReadR1C1Axis(token, ref i);
            return new RowCol(None, 0, loneCol.Type, loneCol.Value, R1C1);
        }

        // It must be a row.
        if (token[i] is not ('R' or 'r'))
        {
            throw Bug();
        }

        (ReferenceAxisType Type, int Value) row = ReadR1C1Axis(token, ref i);

        // Token is ROW. Either it has ended or it is followed by :
        if (i == token.Length || token[i] is not ('C' or 'c'))
        {
            return new RowCol(row.Type, row.Value, None, 0, R1C1);
        }

        // Token is ROW COLUMN
        (ReferenceAxisType Type, int Value) col = ReadR1C1Axis(token, ref i);
        return new RowCol(row.Type, row.Value, col.Type, col.Value, R1C1);
    }

    /// <summary>
    /// Read the axis value. Can work for row or column.
    /// </summary>
    /// <param name="token">The span of a token.</param>
    /// <param name="currentIdx">Index where is <c>C</c>/<c>R</c>.</param>
    private static (ReferenceAxisType Type, int Value) ReadR1C1Axis(ReadOnlySpan<char> token, ref int currentIdx)
    {
        // There are three possibilities: C only, C[-14] and C123
        int i = currentIdx + 1;
        if (token.Length == i)
        {
            // We are at the end of a formula and the only thing that was left was C/R, an alias for C[0]/R[0]
            currentIdx = i;
            return (Relative, 0);
        }

        if (token[i] == '[')
        {
            // Axis is relative
            ++i;
            bool isNegative = token[i] == '-';
            if (isNegative)
            {
                ++i; // Skip sign character
            }

            // Axis is relative and thus must have a position. No need to check
            // length in the loop, because corresponding there must be ]
            int position = 0;
            do
            {
                position = position * 10 + (token[i++] - '0');
            } while (token[i] >= '0' && token[i] <= '9');

            // Index is at the last character that has to be ']'
            currentIdx = ++i;
            return (Relative, isNegative ? -position : position);
        }

        // Axis is absolute or relative [0] without explicit number.
        int absoluteNumber = 0;
        while (i < token.Length && token[i] >= '0' && token[i] <= '9')
        {
            absoluteNumber = absoluteNumber * 10 + (token[i++] - '0');
        }

        currentIdx = i;

        // There is no number after 'C'/'R' => it's a shorthand for `C[0]`/`R[0]`
        if (absoluteNumber == 0)
        {
            return (Relative, 0);
        }

        return (Absolute, absoluteNumber);
    }

    /// <summary>
    /// Extract info about cell reference from a <c>A1_REFERENCE</c> token.
    /// </summary>
    private static ReferenceArea ParseA1Reference(ReadOnlySpan<char> input)
    {
        // The point of this method is to be fast, not pretty. It assumes that input has
        // already been checked by lexer and thus will never be incorrect.
        int i = 0;
        bool abs1 = IsAbsolute(input, i);
        if (abs1)
        {
            i++;
        }

        bool colStart = IsLetter(input[i]);
        if (!colStart)
        {
            // A1_ROW ':' A1_ROW
            int row1 = ReadRow(input, ref i);
            i++; // Skip ':'
            bool absRow2 = IsAbsolute(input, i);
            if (absRow2)
            {
                i++; // Skip '$'
            }

            int row2 = ReadRow(input, ref i);
            return new ReferenceArea(
                new RowCol(abs1 ? Absolute : Relative, row1, None, 0, A1),
                new RowCol(absRow2 ? Absolute : Relative, row2, None, 0, A1));
        }

        int col = ReadColumn(input, ref i);
        if (input[i] == ':')
        {
            // A1_COLUMN ':' A1_COLUMN
            i++; // Skip ':'
            bool absCol2 = IsAbsolute(input, i);
            if (absCol2)
            {
                i++;
            }

            int col2 = ReadColumn(input, ref i);
            return new ReferenceArea(
                new RowCol(None, 0, abs1 ? Absolute : Relative, col, A1),
                new RowCol(None, 0, absCol2 ? Absolute : Relative, col2, A1));
        }

        bool secondAbsolute = IsAbsolute(input, i);
        if (secondAbsolute)
        {
            // Skip $
            i++;
        }

        // A1_CELL | A1_AREA : A1_CELL ':' A1_CELL
        int row = ReadRow(input, ref i);

        RowCol cell = new(secondAbsolute, row, abs1, col, A1);
        if (i == input.Length)
        {
            // A1_CELL
            return new ReferenceArea(cell, cell);
        }

        // A1_AREA
        i++; // Skip ':'
        RowCol secondCell = ReadA1Cell(input, ref i);
        return new ReferenceArea(cell, secondCell);
    }

    private static RowCol ReadA1Cell(ReadOnlySpan<char> input, ref int i)
    {
        bool colAbs = IsAbsolute(input, i);
        if (colAbs)
        {
            i++;
        }

        int col = ReadColumn(input, ref i);
        bool rowAbs = IsAbsolute(input, i);
        if (rowAbs)
        {
            i++;
        }

        int row = ReadRow(input, ref i);
        return new RowCol(rowAbs, row, colAbs, col, A1);
    }

    private static bool IsAbsolute(ReadOnlySpan<char> input, int startIdx) => input[startIdx] == '$';

    // Call only when first char is column
    private static int ReadColumn(ReadOnlySpan<char> input, ref int startIdx)
    {
        int column = 0;
        int i = startIdx;

        do
        {
            char c = input[i];
            int letter = c < 'a' // A is before a 
                ? c - 'A' + 1
                : c - 'a' + 1;
            column = column * 26 + letter;
            i++;
        } while (i < input.Length && IsLetter(input[i]));

        startIdx = i;
        return column;
    }

    private static int ReadRow(ReadOnlySpan<char> input, ref int startIdx)
    {
        int row = 0;
        int i = startIdx;
        do
        {
            int digit = input[i] - '0';
            row = row * 10 + digit;
            i++;
        } while (i < input.Length && input[i] >= '0' && input[i] <= '9');

        startIdx = i;
        return row;
    }

    internal static void ParseIntraTableReference(ReadOnlySpan<char> input, out StructuredReferenceArea area, out string? firstColumn, out string? lastColumn)
    {
        // Skip first char, it's always '['
        int i = 1;
        if (input[i] == '#')
        {
            // Pattern is a KEYWORD
            area = GetArea(input, i);
            firstColumn = null;
            lastColumn = null;
            return;
        }

        if (input[i] != '[' && input[i] != ' ')
        {
            // Pattern is '[]', '[First]' or '[First:Last]' or '[First:[Last]]'
            // because simple column can't start with a space.
            area = StructuredReferenceArea.None;
            if (input[i] == ']')
            {
                // Pattern is '[]', i.e. whole table.
                firstColumn = null;
                lastColumn = null;
                return;
            }

            // Read simple column
            i = GetStructuredName(input, i, out firstColumn);
            if (i < input.Length && input[i] == ':')
            {
                GetStructuredName(input, i + 1, out lastColumn);
            }
            else
            {
                lastColumn = null;
            }

            return;
        }

        // Pattern is SPACED_LBRACKET INNER_REFERENCE SPACED_RBRACKET

        // Skip potential whitespaces at the beginning of a structured reference (SPACED_LBRACKET)
        i = SkipWhitespaces(input, i);
        area = StructuredReferenceArea.None;
        if (input[i + 1] == '#')
        {
            // Inner reference contains a keyword.
            StructuredReferenceArea listItem = GetArea(input, ++i);
            i += GetLength(listItem) + 1;
            area |= listItem;

            i = SkipComma(input, i);
        }

        if (input[i + 1] == '#')
        {
            // Item is a keyword list, either
            // * '[#Headers]' SPACED_COMMA '[#Data]'
            // * '[#Data]' SPACED_COMMA '[#Totals]'
            StructuredReferenceArea listItem = GetArea(input, ++i);
            i += GetLength(listItem) + 1;
            area |= listItem;

            i = SkipComma(input, i);
        }

        // KEYWORD_LIST can contain at most two item specifiers.
        // After keyword list, we get either a COLUMN or a COLUMN:COLUMN
        i = GetStructuredName(input, i, out firstColumn);
        if (i < input.Length && input[i] == ':')
        {
            GetStructuredName(input, i + 1, out lastColumn);
        }
        else
        {
            lastColumn = null;
        }
    }

    private static int SkipComma(ReadOnlySpan<char> input, int i)
    {
        // comma might be wrapped in whitespaces.
        i = SkipWhitespaces(input, i);

        Debug.Assert(input[i] == ',');
        i++;
        i = SkipWhitespaces(input, i);
        return i;
    }

    private static int SkipWhitespaces(ReadOnlySpan<char> input, int i)
    {
        for (; i < input.Length; i++)
        {
            if (!IsWhiteSpace(input[i]))
            {
                break;
            }
        }

        return i;
    }

    /// <summary>
    /// Read a structured name until the end bracket or column
    /// </summary>
    /// <param name="input">Input span.</param>
    /// <param name="startIdx">First index of expected name. It will either contain a bracket or first letter of column name.</param>
    /// <param name="columnName">Parsed name.</param>
    private static int GetStructuredName(ReadOnlySpan<char> input, int startIdx, out string columnName)
    {
        Span<char> buffer = stackalloc char[input.Length];
        int bufferIdx = 0;
        int i = startIdx + (input[startIdx] == '[' ? 1 : 0);
        char c = input[i];
        for (; c is not ']' and not ':'; c = input[++i])
        {
            if (c == '\'')
            {
                c = input[++i];
            }

            buffer[bufferIdx++] = c;
        }

        columnName = buffer.Slice(0, bufferIdx).ToString();
        return i + (c == ']' ? 1 : 0); // char after last bracket
    }

    private static StructuredReferenceArea GetArea(ReadOnlySpan<char> input, int i)
    {
        // Tokenizer has taken care that input can contain only valid values = only first two chars is enough.
        StructuredReferenceArea item = input[i + 1] switch
        {
            'A' => StructuredReferenceArea.All,
            'a' => StructuredReferenceArea.All,
            'D' => StructuredReferenceArea.Data,
            'd' => StructuredReferenceArea.Data,
            'H' => StructuredReferenceArea.Headers,
            'h' => StructuredReferenceArea.Headers,
            'T' => input[i + 2] switch
            {
                'O' => StructuredReferenceArea.Totals,
                'o' => StructuredReferenceArea.Totals,
                'H' => StructuredReferenceArea.ThisRow,
                'h' => StructuredReferenceArea.ThisRow,
                _ => throw new NotSupportedException()
            },
            _ => throw new NotSupportedException()
        };
        return item;
    }

    private static int GetLength(StructuredReferenceArea item)
    {
        return item switch
        {
            StructuredReferenceArea.All => 4,
            StructuredReferenceArea.Data => 5,
            StructuredReferenceArea.Headers => 8,
            StructuredReferenceArea.ThisRow => 9,
            StructuredReferenceArea.Totals => 7,
            _ => throw new InvalidOperationException()
        };
    }

    private static bool IsWhiteSpace(char c)
    {
        return c is ' ' or '\n' or '\r';
    }

    internal static int ParseBookPrefix(ReadOnlySpan<char> input)
    {
        ExtractWorkbookIndex(input, out int? index);
        Debug.Assert(index.HasValue);
        return index.Value;
    }

    private static bool IsLetter(char c) => (c is >= 'A' and <= 'Z') || (c is >= 'a' and <= 'z');

    private static Exception Bug()
    {
        throw new InvalidOperationException("Bug in token parser. Token doesn't have expected format.");
    }
}
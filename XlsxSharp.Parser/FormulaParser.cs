using System.Globalization;
using JetBrains.Annotations;
using XlsxSharp.Parser.Rolex;

namespace XlsxSharp.Parser;

/// <summary>
/// A parser of Excel formulas, with main purpose of creating an abstract syntax tree.
/// </summary>
/// <remarks>
/// An implementation is a recursive descent parser, based on the ANTLR grammar.
/// </remarks>
/// <typeparam name="TScalarValue">Type of a scalar value used across expressions.</typeparam>
/// <typeparam name="TNode">Type of a node used in the AST.</typeparam>
/// <typeparam name="TContext">A context of the parsing. It's passed to every factory method and can contain global info that doesn't belong individual nodes.</typeparam>
[PublicAPI]
public class FormulaParser<TScalarValue, TNode, TContext>
{
    private const string REF_ERROR = "#REF!";
    private readonly string _input;
    private readonly List<Token> _tokens;
    private readonly IAstFactory<TScalarValue, TNode, TContext> _factory;
    private readonly TContext _context;

    /// <summary>
    /// Is parser in A1 mode (true) or R1C1 mode (false)?
    /// </summary>
    private readonly bool _a1Mode;
    private Token _tokenSource;
    private int _tokenIndex = -1;

    // Current lookahead token index
    private int _la;

    private FormulaParser(string formula, TContext context, IAstFactory<TScalarValue, TNode, TContext> factory, bool a1Mode)
    {
        // Trim the end, so ref_intersection_expression that tried to parse SPACE as an operator
        // doesn't recognize spaces at the end of formula as operators. The control tokens of
        // the formula have whitespaces around them (unlike params), so the whitespaces should
        // be consumed by control tokens (e.g. ` IF ( A1 ) ` will be split into `IF ( `, `A1` and ` ) `)
        // but to avoid the whitespace at the end, trim it.
        ReadOnlySpan<char> trimmedFormula = formula.AsSpan().TrimEnd();
        this._input = formula;
        this._context = context;
        this._tokens = a1Mode
            ? RolexLexer.GetTokensA1(trimmedFormula)
            : RolexLexer.GetTokensR1C1(trimmedFormula);
        this._factory = factory;
        this._a1Mode = a1Mode;
        this.Consume();
    }

    /// <summary>
    /// Parse a formula using A1 semantic for references.
    /// </summary>
    /// <param name="formula">Formula text that will be parsed.</param>
    /// <param name="context">Context that is going to be passed to every method of the <paramref name="factory"/>.</param>
    /// <param name="factory">Factory to create nodes of AST tree.</param>
    /// <exception cref="ParsingException">If the formula doesn't satisfy the grammar.</exception>
    public static TNode CellFormulaA1(string formula, TContext context, IAstFactory<TScalarValue, TNode, TContext> factory)
    {
        FormulaParser<TScalarValue, TNode, TContext> parser = new(formula, context, factory, true);
        return parser.Formula();
    }

    /// <summary>
    /// Parse a formula using R1C1 semantic for references.
    /// </summary>
    /// <param name="formula">Formula text that will be parsed.</param>
    /// <param name="context">Context that is going to be passed to every method of the <paramref name="factory"/>.</param>
    /// <param name="factory">Factory to create nodes of AST tree.</param>
    /// <exception cref="ParsingException">If the formula doesn't satisfy the grammar.</exception>
    public static TNode CellFormulaR1C1(string formula, TContext context, IAstFactory<TScalarValue, TNode, TContext> factory)
    {
        FormulaParser<TScalarValue, TNode, TContext> parser = new(formula, context, factory, false);
        return parser.Formula();
    }

    private TNode Formula()
    {
        if (this._tokens[this._tokens.Count - 1].SymbolId == Token.ErrorSymbolId)
        {
            throw new ParsingException($"Unable to determine token for '{this._input}' at index {this._tokens[this._tokens.Count - 1].StartIndex}.");
        }

        if (this._la == Token.SPACE)
        {
            this.Consume();
        }

        TNode expression = this.Expression(false, out _);
        if (this._la != Token.EofSymbolId)
        {
            string parsedPart = this._input.Substring(0, this._tokenSource.StartIndex);
            string remainder = this._input.Substring(this._tokenSource.StartIndex);
            throw new ParsingException($"The formula `{this._input}` wasn't parsed correctly. The expression `{parsedPart}` was parsed, but the rest `{remainder}` wasn't.");
        }

        return expression;
    }

    private TNode Expression(bool skipRangeUnion, out bool isPureRef)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.ConcatExpression(skipRangeUnion, out isPureRef);
        while (true)
        {
            BinaryOperation cmpOp;
            switch (this._la)
            {
                case Token.GREATER_OR_EQUAL_THAN:
                    cmpOp = BinaryOperation.GreaterOrEqualThan;
                    break;
                case Token.LESS_OR_EQUAL_THAN:
                    cmpOp = BinaryOperation.LessOrEqualThan;
                    break;
                case Token.LESS_THAN:
                    cmpOp = BinaryOperation.LessThan;
                    break;
                case Token.GREATER_THAN:
                    cmpOp = BinaryOperation.GreaterThan;
                    break;
                case Token.NOT_EQUAL:
                    cmpOp = BinaryOperation.NotEqual;
                    break;
                case Token.EQUAL:
                    cmpOp = BinaryOperation.Equal;
                    break;
                default:
                    return leftNode;
            }

            this.Consume();
            isPureRef = false;

            TNode rightNode = this.ConcatExpression(skipRangeUnion, out _);
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), cmpOp, leftNode, rightNode);
        }
    }

    private TNode ConcatExpression(bool skipRangeUnion, out bool isPureRef)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.AdditiveExpression(skipRangeUnion, out isPureRef);
        while (this._la == Token.CONCAT)
        {
            this.Consume();
            isPureRef = false;
            TNode rightNode = this.AdditiveExpression(skipRangeUnion, out _);
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), BinaryOperation.Concat, leftNode, rightNode);
        }

        return leftNode;
    }
    private TNode AdditiveExpression(bool skipRangeUnion, out bool isPureRef)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.MultiplyingExpression(skipRangeUnion, out isPureRef);
        while (true)
        {
            BinaryOperation op;
            switch (this._la)
            {
                case Token.PLUS:
                    op = BinaryOperation.Addition;
                    break;
                case Token.MINUS:
                    op = BinaryOperation.Subtraction;
                    break;
                default:
                    return leftNode;
            }

            this.Consume();
            isPureRef = false;
            TNode rightNode = this.MultiplyingExpression(skipRangeUnion, out _);
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), op, leftNode, rightNode);
        }
    }

    private TNode MultiplyingExpression(bool skipRangeUnion, out bool isPureRef)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.PowExpression(skipRangeUnion, out isPureRef);
        while (true)
        {
            BinaryOperation op;
            switch (this._la)
            {
                case Token.MULT:
                    op = BinaryOperation.Multiplication;
                    break;
                case Token.DIV:
                    op = BinaryOperation.Division;
                    break;
                default:
                    return leftNode;
            }

            this.Consume();
            isPureRef = false;
            TNode appendNode = this.PowExpression(skipRangeUnion, out _);
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), op, leftNode, appendNode);
        }
    }

    private TNode PowExpression(bool skipRangeUnion, out bool isPureRef)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.PercentExpression(skipRangeUnion, out isPureRef);
        while (this._la == Token.POW)
        {
            this.Consume();
            isPureRef = false;
            TNode rightNode = this.PercentExpression(skipRangeUnion, out _);
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), BinaryOperation.Power, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode PercentExpression(bool skipRangeUnion, out bool isPureRef)
    {
        int start = this._tokenSource.StartIndex;
        TNode prefixAtomNode = this.PrefixAtomExpression(skipRangeUnion, out isPureRef);
        TNode percentNode = prefixAtomNode;
        while (this._la == Token.PERCENT)
        {
            this.Consume();
            isPureRef = false;
            percentNode = this._factory.Unary(this._context, new SymbolRange(start, this._tokenSource.StartIndex), UnaryOperation.Percent, percentNode);
        }

        return percentNode;
    }

    /// <summary>
    /// Parser for two rules unified into a single method.
    /// <para>
    /// <c>
    /// prefix_atom_expression
    ///     : (PLUS | MINUS) prefix_atom_expression
    ///     | atom_expression
    ///     ;
    /// </c>
    ///
    /// <c>
    /// arg_prefix_atom_expression
    ///     : (PLUS | MINUS) arg_prefix_atom_expression
    ///     | arg_atom_expression
    ///     ;
    /// </c>
    /// </para>
    /// </summary>
    /// <param name="skipRangeUnion">Does the method represent <c>prefix_atom_expression</c> (<c>false</c>) or <c>arg_prefix_atom_expression</c> (<c>true</c>)</param>
    /// <param name="isPureRef">Is the expression of the node a reference expression?</param>
    /// <returns></returns>
    private TNode PrefixAtomExpression(bool skipRangeUnion, out bool isPureRef)
    {
        int start = this._tokenSource.StartIndex;
        UnaryOperation op;
        switch (this._la)
        {
            case Token.PLUS:
                op = UnaryOperation.Plus;
                break;
            case Token.MINUS:
                op = UnaryOperation.Minus;
                break;
            default:
                return this.AtomExpression(skipRangeUnion, out isPureRef);
        }

        this.Consume();
        TNode neutralAtom = this.PrefixAtomExpression(skipRangeUnion, out _);
        isPureRef = false;
        return this._factory.Unary(this._context, new SymbolRange(start, this._tokenSource.StartIndex), op, neutralAtom);
    }

    private TNode AtomExpression(bool skipRangeUnion, out bool isPureRef)
    {
        switch (this._la)
        {
            // Constant
            case Token.NONREF_ERRORS:
            case Token.LOGICAL_CONSTANT:
            case Token.NUMERICAL_CONSTANT:
            case Token.STRING_CONSTANT:
            case Token.OPEN_CURLY:
                isPureRef = false;
                TNode constantNode = this.Constant();
                return constantNode;

            // '(' expression ')'
            case Token.OPEN_BRACE:
                {
                    int start = this._tokenSource.StartIndex;
                    this.Consume();
                    TNode expression = this.Expression(false, out isPureRef);
                    this.Match(Token.CLOSE_BRACE);
                    TNode nestedNode = this._factory.Nested(this._context, new SymbolRange(start, this._tokenSource.StartIndex), expression);

                    // This is the point of an ambiguity. Atom should be a value, but it can
                    // be determined by calling an expression inside the braces or
                    // it can go through ref_expression path that also has an ref_expression
                    // in braces. The second option should be seriously rare, so parser
                    // tracks whether expression is a ref_expression and if it is, we 'backtrack'
                    // through patching to the correct path of the 'ref_expression' below.
                    // Example: '(A1) A1:B2' <- the '(A1)' should be detected as ref_expression
                    // and thus the ' ' intersection operator be valid. Of course, braces can be
                    // very nested '(((A1))) A1:B2' and when we are entering the brace, there is
                    // no way to detect whether it is ref_expression or expression.
                    if (isPureRef)
                    {
                        // Incorrect expectation, backtrack to the ref_expression
                        // note the passed true argument for 'replaceFirstAtom'
                        if (skipRangeUnion)
                        {
                            return this.RefImplicitExpression(true, nestedNode);
                        }

                        return this.RefExpression(true, nestedNode);
                    }

                    return nestedNode;
                }

            // function_call
            case Token.CELL_FUNCTION_LIST:
                {
                    isPureRef = false;
                    int start = this._tokenSource.StartIndex;
                    RowCol cellReference = TokenParser.ExtractCellFunction(this.GetCurrentToken());
                    this.Consume();
                    IReadOnlyList<TNode> args = this.ArgumentList();
                    SymbolRange range = new(start, this._tokenSource.StartIndex);
                    return this._factory.CellFunction(this._context, range, cellReference, args);
                }

            case Token.USER_DEFINED_FUNCTION_NAME:
                isPureRef = false;
                return this.LocalFunctionCall();

            default:
                // function_call : SINGLE_SHEET_PREFIX USER_DEFINED_FUNCTION_NAME argument_list
                if (this._la == Token.SINGLE_SHEET_PREFIX && this.LL(1) == Token.USER_DEFINED_FUNCTION_NAME)
                {
                    isPureRef = false;
                    int start = this._tokenSource.StartIndex;
                    TokenParser.ParseSingleSheetPrefix(this.GetCurrentToken(), out int? wbIndex, out string sheetName);
                    this.Consume();
                    ReadOnlySpan<char> functionName = TokenParser.ExtractLocalFunctionName(this.GetCurrentToken());
                    this.Consume();
                    IReadOnlyList<TNode> args = this.ArgumentList();
                    SymbolRange range = new(start, this._tokenSource.StartIndex);
                    return wbIndex is null
                        ? this._factory.Function(this._context, range, sheetName, functionName, args)
                        : this._factory.ExternalFunction(this._context, range, wbIndex.Value, sheetName, functionName, args);
                }

                // function_call : BOOK_PREFIX USER_DEFINED_FUNCTION_NAME argument_list
                if (this._la == Token.BOOK_PREFIX && this.LL(1) == Token.USER_DEFINED_FUNCTION_NAME)
                {
                    isPureRef = false;
                    int start = this._tokenSource.StartIndex;
                    int wbIndex = TokenParser.ParseBookPrefix(this.GetCurrentToken());
                    this.Consume();
                    ReadOnlySpan<char> functionName = TokenParser.ExtractLocalFunctionName(this.GetCurrentToken());
                    this.Consume();
                    IReadOnlyList<TNode> args = this.ArgumentList();
                    SymbolRange range = new(start, this._tokenSource.StartIndex);
                    return this._factory.ExternalFunction(this._context, range, wbIndex, functionName, args);
                }

                // ref_expression
                if (skipRangeUnion)
                {
                    isPureRef = true;
                    return this.RefIntersectionExpression();
                }

                isPureRef = true;
                return this.RefExpression();
        }
    }

    private TNode RefExpression(bool replaceFirstAtom = false, TNode? refAtom = default)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.RefImplicitExpression(replaceFirstAtom, refAtom);
        while (this._la == Token.COMMA)
        {
            this.Consume();
            TNode rightNode = this.RefImplicitExpression();
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), BinaryOperation.Union, leftNode, rightNode);
        }

        return leftNode;
    }

    /// <summary>
    /// <code>
    /// ref_implicit_expression
    ///        : INTERSECT ref_implicit_expression
    ///        | ref_intersection_expression
    ///        ;
    /// </code>
    /// </summary>
    private TNode RefImplicitExpression(bool replaceFirstAtom = false, TNode? refAtom = default)
    {
        int start = this._tokenSource.StartIndex;
        if (this._la == Token.INTERSECT)
        {
            this.Consume();
            TNode refNode = this.RefImplicitExpression(replaceFirstAtom, refAtom);
            return this._factory.Unary(this._context, new SymbolRange(start, this._tokenSource.StartIndex), UnaryOperation.ImplicitIntersection, refNode);
        }

        return this.RefIntersectionExpression(replaceFirstAtom, refAtom);
    }

    private TNode RefIntersectionExpression(bool replaceFirstAtom = false, TNode? refAtom = default)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.RefRangeExpression(replaceFirstAtom, refAtom);
        while (this._la == Token.SPACE)
        {
            this.Consume();
            TNode rightNode = this.RefRangeExpression();
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), BinaryOperation.Intersection, leftNode, rightNode);
        }

        return leftNode;
    }

    private TNode RefRangeExpression(bool replaceFirstAtom = false, TNode? refAtom = default)
    {
        int start = this._tokenSource.StartIndex;
        TNode leftNode = this.RefSpillExpression(replaceFirstAtom, refAtom);
        while (this._la == Token.COLON)
        {
            this.Consume();
            TNode rightNode = this.RefSpillExpression();
            leftNode = this._factory.BinaryNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), BinaryOperation.Range, leftNode, rightNode);
        }

        return leftNode;
    }

    /// <summary>
    /// Parser of the following node.
    /// <c>
    /// ref_spill_expression
    ///     : ref_atom_expression SPILL?
    ///     ;
    /// </c>
    /// </summary>
    private TNode RefSpillExpression(bool replaceFirstAtom = false, TNode? refAtom = default)
    {
        int start = this._tokenSource.StartIndex;
        TNode refAtomNode = this.RefAtomExpression(replaceFirstAtom, refAtom);
        if (this._la == Token.SPILL)
        {
            this.Consume();
            return this._factory.Unary(this._context, new SymbolRange(start, this._tokenSource.StartIndex), UnaryOperation.SpillRange, refAtomNode);
        }

        return refAtomNode;
    }

    private TNode RefAtomExpression(bool replaceFirstAtom = false, TNode? refAtom = default)
    {
        // A backtracking of an incorrect detection whether an expression in a braces is value expression or ref expression.
        if (replaceFirstAtom)
        {
            return refAtom!;
        }

        switch (this._la)
        {
            // REF_CONSTANT (a1_reference | REF_CONSTANT)?
            // -> REF_CONSTANT
            // -> REF_CONSTANT REF_CONSTANT
            // -> REF_CONSTANT A1_CELL
            // -> REF_CONSTANT A1_CELL COLON A1_CELL
            // -> REF_CONSTANT A1_SPAN_REFERENCE
            // Happens when sheet is deleted, e.g. `#REF!A1`. Note that #REF is actually a valid
            // name of a sheet, but it must be escaped to be usable ('#REF'!B3) because of '#'.
            case Token.REF_CONSTANT:
                {
                    // In all cases, it is a #REF! error from AST PoV, just with weird tokens.
                    int start = this._tokenSource.StartIndex;
                    ReadOnlySpan<char> errorToken = this.GetCurrentToken();
                    Span<char> normalizedError = stackalloc char[errorToken.Length];
                    errorToken.ToUpperInvariant(normalizedError);
                    this.Match(Token.REF_CONSTANT);

                    if (this._la == Token.REF_CONSTANT)
                    {
                        // -> REF_CONSTANT REF_CONSTANT
                        this.Match(Token.REF_CONSTANT);
                    }
                    else if (this._la == Token.A1_CELL)
                    {
                        // -> REF_CONSTANT A1_CELL
                        this.Match(Token.A1_CELL);
                        if (this._la == Token.COLON && this.LL(1) == Token.A1_CELL)
                        {
                            // -> REF_CONSTANT A1_CELL COLON A1_CELL
                            this.Match(Token.COLON);
                            this.Match(Token.A1_CELL);
                        }
                    }
                    else if (this._la == Token.A1_SPAN_REFERENCE)
                    {
                        // -> REF_CONSTANT A1_SPAN_REFERENCE
                        this.Match(Token.A1_SPAN_REFERENCE);
                    }

                    return this._factory.ErrorNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), normalizedError);
                }

            case Token.OPEN_BRACE:
                {
                    int start = this._tokenSource.StartIndex;
                    this.Consume();
                    TNode refExpression = this.RefExpression();
                    this.Match(Token.CLOSE_BRACE);
                    return this._factory.Nested(this._context, new SymbolRange(start, this._tokenSource.StartIndex), refExpression);
                }
            // cell_reference has been inlined into this switch

            // cell_reference:
            //     (A1_CELL | A1_CELL COLON A1_CELL) -- inlined a1_reference
            case Token.A1_CELL:
                {
                    int startIdx = this._tokenSource.StartIndex;
                    ReferenceArea area = TokenParser.ParseReference(this.GetCurrentToken(), this._a1Mode);
                    this.Consume();
                    if (this._la == Token.COLON && this.LL(1) == Token.A1_CELL)
                    {
                        this.Consume();
                        ReferenceArea secondCell = TokenParser.ParseReference(this.GetCurrentToken(), this._a1Mode);
                        this.Consume();
                        area = new ReferenceArea(area.First, secondCell.First);
                    }

                    int endIdx = this._tokenSource.StartIndex;
                    TNode reference = this._factory.Reference(this._context, new SymbolRange(startIdx, endIdx), area);
                    return reference;
                }

            // cell_reference:
            //     (A1_SPAN_REFERENCE)  -- inlined a1_reference
            case Token.A1_SPAN_REFERENCE:
                {
                    int start = this._tokenSource.StartIndex;
                    ReferenceArea area = TokenParser.ParseReference(this.GetCurrentToken(), this._a1Mode);
                    this.Consume();
                    int end = this._tokenSource.StartIndex;
                    TNode reference = this._factory.Reference(this._context, new SymbolRange(start, end), area);
                    return reference;
                }

            // cell_reference:
            //     BANG_REFERENCE
            case Token.BANG_REFERENCE:
                {
                    // Slice away '!' from the bang reference so it can be parsed.
                    ReadOnlySpan<char> referenceToken = this.GetCurrentToken().Slice(1);
                    int start = this._tokenSource.StartIndex;
                    if (referenceToken.Equals(REF_ERROR.AsSpan(), StringComparison.OrdinalIgnoreCase))
                    {
                        this.Consume();
                        return this._factory.ErrorNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), REF_ERROR.AsSpan());
                    }

                    ReferenceArea reference = TokenParser.ParseReference(referenceToken, this._a1Mode);
                    this.Consume();
                    return this._factory.BangReference(this._context, new SymbolRange(start, this._tokenSource.StartIndex), reference);
                }

            // external_cell_reference: SHEET_RANGE_PREFIX (A1_CELL | A1_CELL COLON A1_CELL | A1_SPAN_REFERENCE)
            case Token.SHEET_RANGE_PREFIX:
                {
                    int start = this._tokenSource.StartIndex;
                    ReadOnlySpan<char> sheetRangePrefixToken = this.GetCurrentToken();
                    TokenParser.ParseSheetRangePrefix(sheetRangePrefixToken, out int? wbIdx, out string firstName,
                        out string secondName);
                    this.Consume();

                    ReferenceArea? area = this.A1Reference();
                    if (area is null)
                    {
                        throw this.UnexpectedTokenError(Token.A1_CELL, Token.A1_SPAN_REFERENCE);
                    }

                    int end = this._tokenSource.StartIndex;
                    return wbIdx is not null
                        ? this._factory.ExternalReference3D(this._context, new SymbolRange(start, end), wbIdx.Value, firstName, secondName, area.Value)
                        : this._factory.Reference3D(this._context, new SymbolRange(start, end), firstName, secondName, area.Value);
                }

            // ref_function_call
            case Token.REF_FUNCTION_LIST:
                return this.LocalFunctionCall();

            // name_reference | structure_reference - all variants are expanded from the grammar.

            // Either defined name or table name for a structure reference
            case Token.NAME:
                {
                    int start = this._tokenSource.StartIndex;
                    ReadOnlySpan<char> localName = this.GetCurrentToken();
                    this.Consume();
                    if (this._la == Token.INTRA_TABLE_REFERENCE)
                    {
                        TokenParser.ParseIntraTableReference(this.GetCurrentToken(), out StructuredReferenceArea specifics, out string? firstColumn, out string? lastColumn);
                        this.Consume();
                        SymbolRange range = new(start, this._tokenSource.StartIndex);
                        return this._factory.StructureReference(this._context, range, localName.ToString(), specifics, firstColumn, lastColumn ?? firstColumn);
                    }

                    // 3D reference
                    if (this._la == Token.COLON && this.LL(1) == Token.SINGLE_SHEET_PREFIX)
                    {
                        string firstSheetName = localName.ToString();
                        this.Consume(); // COLON

                        // TODO: Decouple book prefix from single sheet prefix
                        TokenParser.ParseSingleSheetPrefix(this.GetCurrentToken(), out int? wbIdx, out string lastSheetName);
                        if (wbIdx is not null)
                        {
                            throw this.Error("External workbook not expected.");
                        }

                        this.Consume(); // SINGLE_SHEET_PREFIX

                        // After prefix, there must be A1Reference
                        ReferenceArea? area = this.A1Reference();
                        if (area is null)
                        {
                            throw this.UnexpectedTokenError(Token.A1_CELL, Token.A1_SPAN_REFERENCE);
                        }

                        int end = this._tokenSource.StartIndex;
                        return this._factory.Reference3D(this._context, new SymbolRange(start, end), firstSheetName, lastSheetName, area.Value);
                    }

                    return this._factory.Name(this._context, new SymbolRange(start, this._tokenSource.StartIndex), localName.ToString());
                }

            // reference to another workbook
            case Token.BOOK_PREFIX:
                {
                    int start = this._tokenSource.StartIndex;
                    int bookPrefix = TokenParser.ParseBookPrefix(this.GetCurrentToken());
                    this.Consume();
                    ReadOnlySpan<char> externalName = this.GetCurrentToken();
                    this.Match(Token.NAME);
                    if (this._la == Token.INTRA_TABLE_REFERENCE)
                    {
                        TokenParser.ParseIntraTableReference(this.GetCurrentToken(), out StructuredReferenceArea specifics, out string? firstColumn, out string? lastColumn);
                        this.Consume();
                        SymbolRange range = new(start, this._tokenSource.StartIndex);
                        return this._factory.ExternalStructureReference(this._context, range, bookPrefix, externalName.ToString(), specifics, firstColumn, lastColumn ?? firstColumn);
                    }

                    return this._factory.ExternalName(this._context, new SymbolRange(start, this._tokenSource.StartIndex), bookPrefix, externalName.ToString());
                }
            // name_reference: SINGLE_SHEET_PREFIX NAME
            // external_cell_reference: SINGLE_SHEET_PREFIX (A1_CELL | A1_CELL COLON A1_CELL | A1_SPAN_REFERENCE | REF_CONSTANT)
            case Token.SINGLE_SHEET_PREFIX:
                {
                    int start = this._tokenSource.StartIndex;
                    ReadOnlySpan<char> sheetPrefix = this.GetCurrentToken();
                    TokenParser.ParseSingleSheetPrefix(sheetPrefix, out int? wbIdx, out string sheetName);
                    this.Consume();

                    ReferenceArea? area = this.A1Reference();
                    if (area is not null)
                    {
                        int end = this._tokenSource.StartIndex;
                        return wbIdx is null
                            ? this._factory.SheetReference(this._context, new SymbolRange(start, end), sheetName, area.Value)
                            : this._factory.ExternalSheetReference(this._context, new SymbolRange(start, end), wbIdx.Value, sheetName, area.Value);
                    }

                    if (this._la == Token.REF_CONSTANT)
                    {
                        ReadOnlySpan<char> error = this.GetCurrentToken(); // Sheet1!#REF! is a valid
                        this.Consume();
                        return this._factory.ErrorNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), error);
                    }

                    // name_reference
                    ReadOnlySpan<char> name = this.GetCurrentToken();
                    this.Match(Token.NAME);
                    SymbolRange range = new(start, this._tokenSource.StartIndex);
                    return wbIdx is null
                        ? this._factory.SheetName(this._context, range, sheetName, name.ToString())
                        : this._factory.ExternalSheetName(this._context, range, wbIdx.Value, sheetName, name.ToString());
                }

            // structure_reference - only for formulas directly in the table, e.g. totals row.
            case Token.INTRA_TABLE_REFERENCE:
                {
                    int start = this._tokenSource.StartIndex;
                    ReadOnlySpan<char> localTableReference = this.GetCurrentToken();
                    TokenParser.ParseIntraTableReference(localTableReference, out StructuredReferenceArea specifics, out string? firstColumn, out string? lastColumn);
                    this.Consume();
                    SymbolRange range = new(start, this._tokenSource.StartIndex);
                    return this._factory.StructureReference(this._context, range, specifics, firstColumn, lastColumn ?? firstColumn);
                }
        }

        throw this.UnexpectedTokenError();
    }

    /// <summary>
    /// <code>
    /// a1_reference
    ///     : A1_CELL
    ///     | A1_CELL COLON A1_CELL
    ///     | A1_SPAN_REFERENCE
    ///     ;
    /// </code>
    /// </summary>
    private ReferenceArea? A1Reference()
    {
        if (this._la == Token.A1_CELL)
        {
            ReadOnlySpan<char> cellToken = this.GetCurrentToken();
            ReferenceArea cell = TokenParser.ParseReference(cellToken, this._a1Mode);
            this.Consume();
            ReferenceArea area = cell;
            if (this._la == Token.COLON && this.LL(1) == Token.A1_CELL)
            {
                this.Consume();
                ReferenceArea secondCell = TokenParser.ParseReference(this.GetCurrentToken(), this._a1Mode);
                area = new ReferenceArea(cell.First, secondCell.First);
                this.Consume();
            }

            return area;
        }

        if (this._la == Token.A1_SPAN_REFERENCE)
        {
            ReferenceArea area = TokenParser.ParseReference(this.GetCurrentToken(), this._a1Mode);
            this.Consume();
            return area;
        }

        return null;
    }

    private TNode ErrorNode()
    {
        int start = this._tokenSource.StartIndex;
        ReadOnlySpan<char> errorToken = this.GetCurrentToken();
        Span<char> normalizedError = stackalloc char[errorToken.Length];
        errorToken.ToUpperInvariant(normalizedError);
        this.Consume();
        return this._factory.ErrorNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), normalizedError);
    }

    private TNode Constant()
    {
        int start = this._tokenSource.StartIndex;
        switch (this._la)
        {
            case Token.NONREF_ERRORS:
                return this.ErrorNode();

            case Token.LOGICAL_CONSTANT:
                bool logical = this.ConvertLogical();
                this.Consume();
                return this._factory.LogicalNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), logical);

            case Token.NUMERICAL_CONSTANT:
                double number = this.ConvertNumber();
                this.Consume();
                return this._factory.NumberNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), number);

            case Token.STRING_CONSTANT:
                string text = this.ConvertText();
                this.Consume();
                return this._factory.TextNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), text);

            case Token.OPEN_CURLY:
                this.Consume();
                List<TScalarValue> arrayElements = this.ConstantListRows(out int rows, out int columns);
                this.Match(Token.CLOSE_CURLY);
                return this._factory.ArrayNode(this._context, new SymbolRange(start, this._tokenSource.StartIndex), rows, columns, arrayElements);

            default:
                throw this.UnexpectedTokenError();
        }
    }

    private List<TScalarValue> ConstantListRows(out int rows, out int columns)
    {
        // First use list with doubling strategy
        List<TScalarValue> elements = new();
        int rowSize = this.ConstantListRow(elements);
        int height = 1;
        while (this._la == Token.SEMICOLON)
        {
            this.Consume();
            int nextRowSize = this.ConstantListRow(elements);
            if (nextRowSize != rowSize)
            {
                throw this.Error("Rows of an array don't have same size.");
            }

            height++;
        }

        rows = height;
        columns = rowSize;
        return elements;
    }

    private int ConstantListRow(List<TScalarValue> arrayElements)
    {
        int origSize = arrayElements.Count;
        TScalarValue arrayElement = this.ArrayConstant();
        arrayElements.Add(arrayElement);
        while (this._la == Token.COMMA)
        {
            this.Consume();
            TScalarValue nextElement = this.ArrayConstant();
            arrayElements.Add(nextElement);
        }

        return arrayElements.Count - origSize;
    }

    private TScalarValue ArrayConstant()
    {
        TScalarValue value;
        SymbolRange symbolRange;
        int start = this._tokenSource.StartIndex;
        switch (this._la)
        {
            case Token.REF_CONSTANT:
            case Token.NONREF_ERRORS:
                // Convert to upper case on stack, because length of an error is limited to ~20
                ReadOnlySpan<char> errorToken = this.GetCurrentToken();
                Span<char> normalizedError = stackalloc char[errorToken.Length];
                errorToken.ToUpperInvariant(normalizedError);
                this.Consume();
                symbolRange = new SymbolRange(start, this._tokenSource.StartIndex);
                value = this._factory.ErrorValue(this._context, symbolRange, normalizedError);
                break;

            case Token.LOGICAL_CONSTANT:
                bool logicalValue = this.GetTokenLogicalValue();
                this.Consume();
                symbolRange = new SymbolRange(start, this._tokenSource.StartIndex);
                value = this._factory.LogicalValue(this._context, symbolRange, logicalValue);
                break;
            case Token.MINUS:
                this.Consume();
                if (this._la != Token.NUMERICAL_CONSTANT)
                {
                    throw this.UnexpectedTokenError(Token.NUMERICAL_CONSTANT);
                }

                double negativeNumberValue = -ParseNumber(this.GetCurrentToken());
                this.Consume();
                symbolRange = new SymbolRange(start, this._tokenSource.StartIndex);
                value = this._factory.NumberValue(this._context, symbolRange, negativeNumberValue);
                break;

            case Token.PLUS:
                this.Consume();
                if (this._la != Token.NUMERICAL_CONSTANT)
                {
                    throw this.UnexpectedTokenError(Token.NUMERICAL_CONSTANT);
                }

                double positiveNumberValue = ParseNumber(this.GetCurrentToken());
                this.Consume();
                symbolRange = new SymbolRange(start, this._tokenSource.StartIndex);
                value = this._factory.NumberValue(this._context, symbolRange, positiveNumberValue);
                break;

            case Token.NUMERICAL_CONSTANT:
                double numberValue = ParseNumber(this.GetCurrentToken());
                this.Consume();
                symbolRange = new SymbolRange(start, this._tokenSource.StartIndex);
                value = this._factory.NumberValue(this._context, symbolRange, numberValue);
                break;

            case Token.STRING_CONSTANT:
                ReadOnlySpan<char> token = this.GetCurrentToken();
                Span<char> buffer = stackalloc char[token.Length];
                this.Consume();
                symbolRange = new SymbolRange(start, this._tokenSource.StartIndex);
                value = ConvertTextValue(token, out ReadOnlySpan<char> slice, ref buffer)
                    ? this._factory.TextValue(this._context, symbolRange, slice.ToString())
                    : this._factory.TextValue(this._context, symbolRange, buffer.ToString());
                break;

            default:
                throw this.UnexpectedTokenError();
        }

        return value;
    }

    private IReadOnlyList<TNode> ArgumentList()
    {
        // A special case, there are no arguments
        if (this._la == Token.CLOSE_BRACE)
        {
            this.Consume();
            return Array.Empty<TNode>();
        }

        List<TNode> args = new();
        while (true)
        {
            // At the start of the loop, previous argument
            // should have been consumed with a comma.
            if (this._la == Token.COMMA)
            {
                // If there is a comma, it means there are
                // two commas in a row and thus a blank argument.
                int start = this._tokenSource.StartIndex;
                this.Consume();
                args.Add(this._factory.BlankNode(this._context, new SymbolRange(start, start)));
            }
            else if (this._la == Token.CLOSE_BRACE)
            {
                // if there is a brace, it means the previous
                // comma is immediately followed by a brace `,)`
                // thus there is a blank node and end of args.
                int start = this._tokenSource.StartIndex;
                this.Consume();
                args.Add(this._factory.BlankNode(this._context, new SymbolRange(start, start)));
                return args;
            }
            else
            {
                // Path for a non-blank argument.
                TNode arg = this.Expression(true, out _);
                args.Add(arg);
                if (this._la == Token.CLOSE_BRACE)
                {
                    this.Consume();
                    return args;
                }

                // Each argument must be followed by a comma.
                this.Match(Token.COMMA);
            }
        }
    }

    private void Match(int expected)
    {
        if (this._la != expected)
        {
            throw this.UnexpectedTokenError(expected);
        }

        this.Consume();
    }

    private void Consume()
    {
        this._tokenSource = this._tokens[++this._tokenIndex];
        this._la = this._tokenSource.SymbolId;
    }

    private int LL(int lookAhead)
    {
        int idx = this._tokenIndex + lookAhead;
        return idx < this._tokens.Count ? this._tokens[idx].SymbolId : Token.EofSymbolId;
    }

    private static double ParseNumber(ReadOnlySpan<char> number)
    {
        return double.Parse(
            number,
            NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent,
            CultureInfo.InvariantCulture);
    }

    private bool ConvertLogical()
    {
        return this.GetTokenLogicalValue();
    }

    private bool GetTokenLogicalValue()
    {
        return this._input[this._tokenSource.StartIndex] is 'T' or 't';
    }

    private double ConvertNumber()
    {
        return ParseNumber(this.GetCurrentToken());
    }

    private string ConvertText()
    {
        ReadOnlySpan<char> token = this.GetCurrentToken();
        Span<char> buffer = stackalloc char[token.Length];
        return ConvertTextValue(token, out ReadOnlySpan<char> slice, ref buffer)
            ? slice.ToString()
            : buffer.ToString();
    }

    private static bool ConvertTextValue(ReadOnlySpan<char> token, out ReadOnlySpan<char> copy, ref Span<char> buffer)
    {
        ReadOnlySpan<char> text = token.Slice(1, token.Length - 2);
        int indexOfDQuote = text.IndexOf('"');
        bool textMustBeUnescaped = indexOfDQuote >= 0;
        if (!textMustBeUnescaped)
        {
            copy = text;
            return true;
        }

        Span<char> unescaped = buffer;
        Span<char> tail = unescaped;
        int quoteCount = 0;
        do
        {
            ReadOnlySpan<char> quoteText = text.Slice(0, indexOfDQuote + 1);
            quoteText.CopyTo(tail);
            tail = tail.Slice(indexOfDQuote + 1);
            text = text.Slice(indexOfDQuote + 2);
            indexOfDQuote = text.IndexOf('"');
            quoteCount++;
        } while (indexOfDQuote >= 0);

        text.CopyTo(tail);
        buffer = unescaped.Slice(0, token.Length - 2 - quoteCount);
        copy = default;
        return false;
    }

    private TNode LocalFunctionCall()
    {
        int start = this._tokenSource.StartIndex;
        ReadOnlySpan<char> functionName = TokenParser.ExtractLocalFunctionName(this.GetCurrentToken());
        this.Consume();
        IReadOnlyList<TNode> args = this.ArgumentList();
        SymbolRange range = new(start, this._tokenSource.StartIndex);
        return this._factory.Function(this._context, range, functionName, args);
    }

    private Exception UnexpectedTokenError(params int[] expectedToken)
    {
        return this.Error($"Unexpected token {this.GetLaTokenName()}, expected one of {string.Join(",", expectedToken.Select(GetTokenName))}.");
    }

    private Exception UnexpectedTokenError()
    {
        return this.Error($"Unexpected token {this.GetLaTokenName()}.");
    }

    private Exception Error(string message)
    {
        return new ParsingException($"Error at char {this._tokenSource.StartIndex} of '{this._input}': {message}");
    }

    private ReadOnlySpan<char> GetCurrentToken()
    {
        return this._input.AsSpan(this._tokenSource.StartIndex, this._tokenSource.Length);
    }

    private static string GetTokenName(int tokenType) => Token.GetSymbolName(tokenType);

    private string GetLaTokenName() => GetTokenName(this._la);
}

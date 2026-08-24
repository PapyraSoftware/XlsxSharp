namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// Function definition class (keeps function name, parameter counts, and delegate).
/// </summary>
internal class FunctionDefinition
{
    private readonly CalcEngineFunction _function;

    private readonly FunctionFlags _flags;

    private readonly AllowRange _allowRanges;

    /// <summary>
    /// Which parameters of the function are marked. The values are indexes of the function parameters, starting from 0.
    /// Used to determine which arguments allow ranges and which don't.
    /// </summary>
    private readonly IReadOnlyCollection<int> _markedParams;

    public FunctionDefinition(
        int minParams,
        int maxParams,
        CalcEngineFunction function,
        FunctionFlags flags,
        AllowRange allowRanges,
        IReadOnlyCollection<int> markedParams
    )
    {
        if (allowRanges == AllowRange.None && markedParams.Any())
        {
            throw new ArgumentException(nameof(markedParams));
        }

        this.MinParams = minParams;
        this.MaxParams = maxParams;
        this._allowRanges = allowRanges;
        this._markedParams = markedParams;
        this._function = function;
        this._flags = flags;
    }

    public int MinParams { get; }

    public int MaxParams { get; }

    public AnyValue CallFunction(CalcContext ctx, Span<AnyValue> args)
    {
        if (CalcContext.UseImplicitIntersection)
        {
            this.IntersectArguments(ctx, args);
        }

        return this._function(ctx, args);
    }

    /// <summary>
    /// Evaluate the function with array formula semantic.
    /// </summary>
    public AnyValue CallAsArray(CalcContext ctx, Span<AnyValue> args)
    {
        if (this._flags.HasFlag(FunctionFlags.ReturnsArray) && this._allowRanges == AllowRange.All)
        {
            return this._function!(ctx, args);
        }

        // Step 1: For scalar parameters of function, determine maximum size of scalar
        // parameters from argument arrays
        (int totalRows, int totalColumns) = this.GetScalarArgsMaxSize(args);

        // Step 2: Normalize arguments. Single params are converted to array of same size, multi params are converted from scalars
        for (int i = 0; i < args.Length; ++i)
        {
            ref AnyValue arg = ref args[i];
            bool argIsSingle = arg.TryPickSingleOrMultiValue(
                out ScalarValue single,
                out Array multi,
                ctx
            );
            if (this.IsParameterSingleValue(i))
            {
                arg = argIsSingle
                    ? new ScalarArray(single, totalColumns, totalRows)
                    : multi.Broadcast(totalRows, totalColumns);
            }
            else
            {
                // 18.17.2.4 When a function expects a multi-valued argument but a single-valued
                // expression is passed, that single-valued argument is treated as a 1x1 array.
                // If there is an error as a single value, e.g. reference to a single cell, the SUMIF behaves
                // as it was converted to 1x1 array and doesn't return error, just because it found an error.
                // Ergo: for ranges, we don't immediately return error, just because range parameter contains an error
                arg = argIsSingle ? new ScalarArray(single, 1, 1) : multi;
            }
        }

        // Step 3: For each item in total array, calculate function
        ScalarValue[,] result = new ScalarValue[totalRows, totalColumns];
        for (int row = 0; row < totalRows; ++row)
        {
            for (int column = 0; column < totalColumns; ++column)
            {
                AnyValue[] itemArg = new AnyValue[args.Length];
                for (int i = 0; i < itemArg.Length; ++i)
                {
                    ref AnyValue arg = ref args[i];
                    itemArg[i] = this.IsParameterSingleValue(i)
                        ? arg.GetArray()[row, column].ToAnyValue()
                        : arg;
                }

                AnyValue itemResult = this._function(ctx, args);

                // Even if function returns an array, only the top-left value of array is used
                // as a result for the item, per tests with FILTERXML.
                result[row, column] = itemResult.TryPickSingleOrMultiValue(
                    out ScalarValue scalarResult,
                    out Array arrayResult,
                    ctx
                )
                    ? scalarResult
                    : arrayResult[0, 0];
            }
        }

        return new ConstArray(result);
    }

    private void IntersectArguments(CalcContext ctx, Span<AnyValue> args)
    {
        for (int i = 0; i < args.Length; ++i)
        {
            bool intersectArgument = this._allowRanges switch
            {
                AllowRange.None => true,
                AllowRange.Except => this._markedParams.Contains(i),
                AllowRange.Only => !this._markedParams.Contains(i),
                AllowRange.All => false,
                _ => throw new InvalidOperationException($"Unexpected value {this._allowRanges}"),
            };
            if (intersectArgument)
            {
                args[i] = args[i].ImplicitIntersection(ctx);
            }
        }
    }

    private (int Rows, int Columns) GetScalarArgsMaxSize(Span<AnyValue> args)
    {
        int maxRows = 1;
        int maxColumns = 1;
        for (int i = 0; i < args.Length; ++i)
        {
            ref AnyValue arg = ref args[i];
            if (this.IsParameterSingleValue(i))
            {
                (int argRows, int argColumns) = arg.GetArraySize();
                maxRows = Math.Max(maxRows, argRows);
                maxColumns = Math.Max(maxColumns, argColumns);
            }
        }

        return (maxRows, maxColumns);
    }

    private bool IsParameterSingleValue(int paramIndex)
    {
        bool paramAllowsMultiValues = this._allowRanges switch
        {
            AllowRange.None => false,
            AllowRange.Except => !this._markedParams.Contains(paramIndex),
            AllowRange.Only => this._markedParams.Contains(paramIndex),
            AllowRange.All => true,
            _ => throw new NotSupportedException($"Unexpected value {this._allowRanges}"),
        };
        return !paramAllowsMultiValues;
    }
}

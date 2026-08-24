namespace XlsxSharp.Excel.CalcEngine.Functions;

internal class TallyNumbers : ITally
{
    private readonly bool _ignoreScalarBlank;
    private readonly bool _ignoreErrors;
    private readonly Func<CalcContext, Reference, IEnumerable<ScalarValue>> _getNonBlankValues;

    /// <summary>
    /// Tally numbers.
    /// </summary>
    internal static readonly TallyNumbers Default = new();

    /// <summary>
    /// Ignore blank from scalar values. Basically used for <c>PRODUCT</c> function, so it doesn't end up with 0.
    /// </summary>
    internal static readonly TallyNumbers WithoutScalarBlank = new(ignoreScalarBlank: true);

    /// <summary>
    /// Tally algorithm for <c>SUBTOTAL</c> functions 1..11.
    /// </summary>
    internal static readonly TallyNumbers Subtotal10 = new(
        static (ctx, reference) => ctx.GetFilteredNonBlankValues(reference, "SUBTOTAL")
    );

    /// <summary>
    /// Tally algorithm for <c>SUBTOTAL</c> functions 101..111.
    /// </summary>
    internal static readonly TallyNumbers Subtotal100 = new(
        static (ctx, reference) =>
            ctx.GetFilteredNonBlankValues(reference, "SUBTOTAL", skipHiddenRows: true)
    );

    /// <summary>
    /// Tally numbers. Any error (including conversion), logical, text is ignored and not tallied.
    /// </summary>
    internal static readonly TallyNumbers IgnoreErrors = new(ignoreErrors: true);

    private TallyNumbers(
        Func<CalcContext, Reference, IEnumerable<ScalarValue>>? getNonBlankValues = null,
        bool ignoreScalarBlank = false,
        bool ignoreErrors = false
    )
    {
        this._ignoreScalarBlank = ignoreScalarBlank;
        this._ignoreErrors = ignoreErrors;
        this._getNonBlankValues =
            getNonBlankValues ?? (static (ctx, reference) => ctx.GetNonBlankValues(reference));
    }

    /// <summary>
    /// The method tries to convert scalar arguments to numbers, but ignores non-numbers in
    /// reference/array. Any error found is propagated to the result.
    /// </summary>
    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>
    {
        T tally = initialState;
        foreach (AnyValue arg in args)
        {
            if (arg.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
            {
                ctx.ThrowIfCancelled();
                if (this._ignoreScalarBlank && scalar.IsBlank)
                {
                    continue;
                }

                // Scalars are converted to number.
                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out double number, out XLError error))
                {
                    if (this._ignoreErrors)
                    {
                        continue;
                    }

                    return error;
                }

                tally = tally.Tally(number);
            }
            else
            {
                IEnumerable<ScalarValue> valuesIterator = !collection.TryPickT0(
                    out Array? array,
                    out Reference? reference
                )
                    ? this._getNonBlankValues(ctx, reference)
                    : array;
                foreach (ScalarValue value in valuesIterator)
                {
                    ctx.ThrowIfCancelled();
                    if (value.TryPickError(out XLError error))
                    {
                        if (this._ignoreErrors)
                        {
                            continue;
                        }

                        return error;
                    }

                    // For arrays and references, only the number type is used. Other types are ignored.
                    if (value.TryPickNumber(out double number))
                    {
                        tally = tally.Tally(number);
                    }
                }
            }
        }

        return tally;
    }
}

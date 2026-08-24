namespace XlsxSharp.Excel.CalcEngine.Functions;

internal interface ITally
{
    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>;
}

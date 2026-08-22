namespace XlsxSharp.Excel.CalcEngine.Functions;

internal interface ITallyState<out TState>
{
    public TState Tally(double number);
}

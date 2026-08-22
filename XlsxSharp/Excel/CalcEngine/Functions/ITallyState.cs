namespace XlsxSharp.Excel.CalcEngine.Functions;

internal interface ITallyState<out TState>
{
    TState Tally(double number);
}

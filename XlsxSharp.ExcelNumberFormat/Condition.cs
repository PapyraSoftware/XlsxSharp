namespace XlsxSharp.ExcelNumberFormat;

internal class Condition
{
    public string Operator { get; set; }
    public double Value { get; set; }

    public bool Evaluate(double lhs)
    {
        switch (this.Operator)
        {
            case "<":
                return lhs < this.Value;
            case "<=":
                return lhs <= this.Value;
            case ">":
                return lhs > this.Value;
            case ">=":
                return lhs >= this.Value;
            case "<>":
                return lhs != this.Value;
            case "=":
                return lhs == this.Value;
        }

        return false;
    }
}

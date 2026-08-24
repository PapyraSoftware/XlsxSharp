#nullable disable

using XlsxSharp.Extensions;
using XlsxSharp.Parser;

namespace XlsxSharp.Excel.Misc;

public class XLFormula
{
    public XLFormula() { }

    public XLFormula(XLFormula defaultFormula)
    {
        this._value = defaultFormula._value;
        this.IsFormula = defaultFormula.IsFormula;
    }

    public XLFormula(string value) => this.Value = value;

    public XLFormula(double value) => this.Value = value.ToInvariantString();

    public XLFormula(int value) => this.Value = value.ToInvariantString();

    internal string _value;
    public string Value
    {
        get => this._value;
        set
        {
            if (value == null)
            {
                this._value = string.Empty;
            }
            else
            {
                this._value = value.Trim();
                this.IsFormula =
                    !string.IsNullOrWhiteSpace(this._value) && this._value.TrimStart()[0] == '=';
                if (this.IsFormula)
                {
                    this._value = this._value.Substring(1);
                }
            }
        }
    }

    public bool IsFormula { get; internal set; }

    internal XLFormula GetAdjustedCopy(Point sourceAnchor, Point targetAnchor)
    {
        if (!this.IsFormula)
        {
            return new XLFormula(this);
        }

        string formulaR1C1 = FormulaConverter.ToR1C1(
            this.Value,
            sourceAnchor.Row,
            sourceAnchor.Column
        );
        string formulaA1 = FormulaConverter.ToA1(
            formulaR1C1,
            targetAnchor.Row,
            targetAnchor.Column
        );
        return new XLFormula { _value = formulaA1, IsFormula = true };
    }
}

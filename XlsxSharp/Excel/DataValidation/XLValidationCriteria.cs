#nullable disable

namespace XlsxSharp.Excel.DataValidation;

public abstract class XLValidationCriteria : IXLValidationCriteria
{
    protected IXLDataValidation dataValidation;

    protected XLValidationCriteria(IXLDataValidation dataValidation) =>
        this.dataValidation = dataValidation;

    #region IXLValidationCriteria Members

    public void Between(string minValue, string maxValue)
    {
        this.dataValidation.MinValue = minValue;
        this.dataValidation.MaxValue = maxValue;
        this.dataValidation.Operator = XLOperator.Between;
    }

    public void Between(IXLCell minValue, IXLCell maxValue)
    {
        this.dataValidation.MinValue = minValue.Address.ToStringFixed();
        this.dataValidation.MaxValue = maxValue.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.Between;
    }

    public void EqualOrGreaterThan(string value)
    {
        this.dataValidation.Value = value;
        this.dataValidation.Operator = XLOperator.EqualOrGreaterThan;
    }

    public void EqualOrGreaterThan(IXLCell cell)
    {
        this.dataValidation.Value = cell.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.EqualOrGreaterThan;
    }

    public void EqualOrLessThan(string value)
    {
        this.dataValidation.Value = value;
        this.dataValidation.Operator = XLOperator.EqualOrLessThan;
    }

    public void EqualOrLessThan(IXLCell cell)
    {
        this.dataValidation.Value = cell.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.EqualOrLessThan;
    }

    public void EqualTo(string value)
    {
        this.dataValidation.Value = value;
        this.dataValidation.Operator = XLOperator.EqualTo;
    }

    public void EqualTo(IXLCell cell)
    {
        this.dataValidation.Value = cell.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.EqualTo;
    }

    public void GreaterThan(string value)
    {
        this.dataValidation.Value = value;
        this.dataValidation.Operator = XLOperator.GreaterThan;
    }

    public void GreaterThan(IXLCell cell)
    {
        this.dataValidation.Value = cell.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.GreaterThan;
    }

    public void LessThan(string value)
    {
        this.dataValidation.Value = value;
        this.dataValidation.Operator = XLOperator.LessThan;
    }

    public void LessThan(IXLCell cell)
    {
        this.dataValidation.Value = cell.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.LessThan;
    }

    public void NotBetween(string minValue, string maxValue)
    {
        this.dataValidation.MinValue = minValue;
        this.dataValidation.MaxValue = maxValue;
        this.dataValidation.Operator = XLOperator.NotBetween;
    }

    public void NotBetween(IXLCell minValue, IXLCell maxValue)
    {
        this.dataValidation.MinValue = minValue.Address.ToStringFixed();
        this.dataValidation.MaxValue = maxValue.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.NotBetween;
    }

    public void NotEqualTo(string value)
    {
        this.dataValidation.Value = value;
        this.dataValidation.Operator = XLOperator.NotEqualTo;
    }

    public void NotEqualTo(IXLCell cell)
    {
        this.dataValidation.Value = cell.Address.ToStringFixed();
        this.dataValidation.Operator = XLOperator.NotEqualTo;
    }

    #endregion IXLValidationCriteria Members
}

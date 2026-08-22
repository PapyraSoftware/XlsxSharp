#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.DataValidation;

public class XLDateCriteria : XLValidationCriteria
{
    public XLDateCriteria(IXLDataValidation dataValidation)
        : base(dataValidation) { }

    public void Between(DateTime minValue, DateTime maxValue)
    {
        this.dataValidation.MinValue = minValue.ToOADate().ToInvariantString();
        this.dataValidation.MaxValue = maxValue.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.Between;
    }

    public void EqualOrGreaterThan(DateTime value)
    {
        this.dataValidation.Value = value.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.EqualOrGreaterThan;
    }

    public void EqualOrLessThan(DateTime value)
    {
        this.dataValidation.Value = value.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.EqualOrLessThan;
    }

    public void EqualTo(DateTime value)
    {
        this.dataValidation.Value = value.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.EqualTo;
    }

    public void GreaterThan(DateTime value)
    {
        this.dataValidation.Value = value.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.GreaterThan;
    }

    public void LessThan(DateTime value)
    {
        this.dataValidation.Value = value.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.LessThan;
    }

    public void NotBetween(DateTime minValue, DateTime maxValue)
    {
        this.dataValidation.MinValue = minValue.ToOADate().ToInvariantString();
        this.dataValidation.MaxValue = maxValue.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.NotBetween;
    }

    public void NotEqualTo(DateTime value)
    {
        this.dataValidation.Value = value.ToOADate().ToInvariantString();
        this.dataValidation.Operator = XLOperator.NotEqualTo;
    }
}

#nullable disable

namespace XlsxSharp.Excel.DataValidation;

public interface IXLValidationCriteria
{
    public void Between(string minValue, string maxValue);

    public void Between(IXLCell minValue, IXLCell maxValue);

    public void EqualOrGreaterThan(string value);

    public void EqualOrGreaterThan(IXLCell cell);

    public void EqualOrLessThan(string value);

    public void EqualOrLessThan(IXLCell cell);

    public void EqualTo(string value);

    public void EqualTo(IXLCell cell);

    public void GreaterThan(string value);

    public void GreaterThan(IXLCell cell);

    public void LessThan(string value);

    public void LessThan(IXLCell cell);

    public void NotBetween(string minValue, string maxValue);

    public void NotBetween(IXLCell minValue, IXLCell maxValue);

    public void NotEqualTo(string value);

    public void NotEqualTo(IXLCell cell);
}

#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel.DataValidation;

public enum XLAllowedValues
{
    AnyValue,
    WholeNumber,
    Decimal,
    Date,
    Time,
    TextLength,
    List,
    Custom,
}

public enum XLErrorStyle
{
    Stop,
    Warning,
    Information,
}

public enum XLOperator
{
    EqualTo,
    NotEqualTo,
    GreaterThan,
    LessThan,
    EqualOrGreaterThan,
    EqualOrLessThan,
    Between,
    NotBetween,
}

public interface IXLDataValidation
{
    public XLAllowedValues AllowedValues { get; set; }

    public XLDateCriteria Date { get; }

    public XLDecimalCriteria Decimal { get; }

    public string ErrorMessage { get; set; }

    public XLErrorStyle ErrorStyle { get; set; }

    public string ErrorTitle { get; set; }

    public bool IgnoreBlanks { get; set; }

    public bool InCellDropdown { get; set; }

    public string InputMessage { get; set; }

    public string InputTitle { get; set; }

    public string MaxValue { get; set; }

    public string MinValue { get; set; }

    public XLOperator Operator { get; set; }

    /// <summary>
    /// A collection of ranges the data validation rule applies too.
    /// </summary>
    public IEnumerable<IXLRange> Ranges { get; }

    public bool ShowErrorMessage { get; set; }

    //void Delete();
    //void CopyFrom(IXLDataValidation dataValidation);
    public bool ShowInputMessage { get; set; }

    public XLTextLengthCriteria TextLength { get; }

    public XLTimeCriteria Time { get; }

    public string Value { get; set; }

    public XLWholeNumberCriteria WholeNumber { get; }

    /// <summary>
    /// Add a range to the collection of ranges this rule applies to.
    /// If the specified range does not belong to the worksheet of the data validation
    /// rule it is transferred to the target worksheet.
    /// </summary>
    /// <param name="range">A range to add.</param>
    public void AddRange(IXLRange range);

    /// <summary>
    /// Add a collection of ranges to the collection of ranges this rule applies to.
    /// Ranges that do not belong to the worksheet of the data validation
    /// rule are transferred to the target worksheet.
    /// </summary>
    /// <param name="ranges">Ranges to add.</param>
    public void AddRanges(IEnumerable<IXLRange> ranges);

    public void Clear();

    /// <summary>
    /// Detach data validation rule of all ranges it applies to.
    /// </summary>
    public void ClearRanges();

    public void Custom(string customValidation);

    public bool IsDirty();

    public void List(string list);

    public void List(string list, bool inCellDropdown);

    public void List(IXLRange range);

    public void List(IXLRange range, bool inCellDropdown);

    /// <summary>
    /// Remove the specified range from the collection of range this rule applies to.
    /// </summary>
    /// <param name="range">A range to remove.</param>
    public bool RemoveRange(IXLRange range);
}

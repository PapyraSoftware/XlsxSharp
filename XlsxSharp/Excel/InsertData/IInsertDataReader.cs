namespace XlsxSharp.Excel.InsertData;

/// <summary>
/// A universal interface for different data readers used in InsertData logic.
/// </summary>
internal interface IInsertDataReader
{
    /// <summary>
    /// Get a collection of records, each as a collection of values, extracted from a source.
    /// </summary>
    public IEnumerable<IEnumerable<XLCellValue>> GetRecords();

    /// <summary>
    /// Get the number of properties to use as a table with.
    /// Actual number of may vary in different records.
    /// </summary>
    public int GetPropertiesCount();

    /// <summary>
    /// Get the title of the property with the specified index.
    /// </summary>
    public string? GetPropertyName(int propertyIndex);
}

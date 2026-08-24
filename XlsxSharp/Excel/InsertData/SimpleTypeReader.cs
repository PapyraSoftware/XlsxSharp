#nullable disable

using System.Collections;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.InsertData;

internal class SimpleTypeReader : IInsertDataReader
{
    private readonly IEnumerable<object> _data;
    private readonly Type _itemType;

    public SimpleTypeReader(IEnumerable data)
    {
        this._data = data?.Cast<object>() ?? throw new ArgumentNullException(nameof(data));
        this._itemType = data.GetItemType();
    }

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords() =>
        this._data.Select(item => new[] { item }.Select(XLCellValue.FromInsertedObject));

    public int GetPropertiesCount() => 1;

    public string GetPropertyName(int propertyIndex = 0)
    {
        if (propertyIndex != 0)
        {
            throw new ArgumentException("SimpleTypeReader supports only a single property");
        }

        return this._itemType.Name;
    }
}

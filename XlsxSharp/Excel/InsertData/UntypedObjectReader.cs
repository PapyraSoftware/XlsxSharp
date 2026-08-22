#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel.InsertData;

internal class UntypedObjectReader : IInsertDataReader
{
    private readonly IEnumerable<object> _data;
    private readonly IEnumerable<IInsertDataReader> _readers;

    public UntypedObjectReader(IEnumerable data)
    {
        this._data = (data ?? new object[0]).Cast<object>();
        this._readers = CreateReaders().ToList();

        IEnumerable<IInsertDataReader> CreateReaders()
        {
            if (!this._data.Any())
            {
                yield break;
            }

            List<object> itemsOfSameType = [];
            Type previousType = null;

            foreach (object item in this._data)
            {
                Type currentType = item?.GetType();

                if (previousType != currentType && itemsOfSameType.Count > 0)
                {
                    yield return CreateReader(itemsOfSameType, previousType);
                    itemsOfSameType.Clear();
                }
                itemsOfSameType.Add(item);
                previousType = currentType;
            }

            if (itemsOfSameType.Count > 0)
            {
                yield return CreateReader(itemsOfSameType, previousType);
            }
        }

        IInsertDataReader CreateReader(List<object> itemsOfSameType, Type itemType)
        {
            if (itemType == null)
            {
                return new NullDataReader(itemsOfSameType);
            }

            Array items = Array.CreateInstance(itemType, itemsOfSameType.Count);
            Array.Copy(itemsOfSameType.ToArray(), items, items.Length);

            return InsertDataReaderFactory.Instance.CreateReader(items);
        }
    }

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
    {
        foreach (IInsertDataReader reader in this._readers)
        {
            foreach (IEnumerable<XLCellValue> item in reader.GetRecords())
            {
                yield return item;
            }
        }
    }

    public int GetPropertiesCount() => this.GetFirstNonNullReader()?.GetPropertiesCount() ?? 0;

    public string GetPropertyName(int propertyIndex) =>
        this.GetFirstNonNullReader()?.GetPropertyName(propertyIndex);

    private IInsertDataReader GetFirstNonNullReader() =>
        this._readers.FirstOrDefault(r => !(r is NullDataReader));
}

// Keep this file CodeMaid organised and cleaned

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using XlsxSharp.Attributes;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.InsertData;

internal class ObjectReader : IInsertDataReader
{
    private const BindingFlags MemberBindingFlags =
        BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;

    private readonly IEnumerable<object> _data;
    private readonly MemberInfo[] _members;
    private readonly bool[] _staticMembers;

    public ObjectReader(IEnumerable data)
    {
        this._data = data.Cast<object>();

        Type itemType = data.GetItemType()!;
        if (itemType.IsNullableType())
        {
            itemType = itemType.GetUnderlyingType();
        }

        this._members =
        [
            .. itemType
                .GetFields(MemberBindingFlags)
                .Cast<MemberInfo>()
                .Concat(
                    itemType
                        .GetProperties(MemberBindingFlags)
                        .Where(pi => !pi.GetIndexParameters().Any())
                )
                .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                .OrderBy(XLColumnAttribute.GetOrder),
        ];

        this._staticMembers = [.. this._members.Select(ReflectionExtensions.IsStatic)];
    }

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords() =>
        this._data.Select(item => this.GetItemData(item).Select(XLCellValue.FromInsertedObject));

    public int GetPropertiesCount() => this._members.Length;

    public string? GetPropertyName(int propertyIndex)
    {
        if (propertyIndex < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(propertyIndex),
                "Property index must be non-negative"
            );
        }

        if (propertyIndex >= this.GetPropertiesCount())
        {
            throw new ArgumentOutOfRangeException(
                $"{propertyIndex} exceeds the number of the object properties"
            );
        }

        MemberInfo memberInfo = this._members[propertyIndex];
        string? fieldName = XLColumnAttribute.GetHeader(memberInfo);
        if (string.IsNullOrWhiteSpace(fieldName))
        {
            fieldName = memberInfo.Name;
        }

        return fieldName;
    }

    private IEnumerable<object?> GetItemData(object item)
    {
        for (int i = 0; i < this._members.Length; i++)
        {
            if (item == null)
            {
                yield return null;
                continue;
            }

            MemberInfo memberInfo = this._members[i];
            switch (memberInfo)
            {
                case PropertyInfo propertyInfo when this._staticMembers[i]:
                    yield return propertyInfo.GetValue(null, null);
                    break;

                case PropertyInfo propertyInfo:
                    yield return propertyInfo.GetValue(item, null);
                    break;

                case FieldInfo fieldInfo when this._staticMembers[i]:
                    yield return fieldInfo.GetValue(null);
                    break;

                case FieldInfo fieldInfo:
                    yield return fieldInfo.GetValue(item);
                    break;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using XlsxSharp.IO.CodeGen.Model;

namespace XlsxSharp.IO.CodeGen.XsdParser;

/// <summary>
/// Mapper for enums found in XSD of OOXML.
/// </summary>
public class XsdEnumMapper : IEnumMapper
{
    private readonly Dictionary<Type, object> _textToEnumMaps = new();

    public XsdEnumMapper() => this.AddMaps();

    public bool TryGetEnum<TEnum>(string text, out TEnum enumValue)
        where TEnum : struct, Enum
    {
        Dictionary<string, TEnum> enumMap =
            (Dictionary<string, TEnum>)this._textToEnumMaps[typeof(TEnum)];
        return enumMap.TryGetValue(text, out enumValue);
    }

    public bool TryGetText<TEnum>(TEnum enumValue, out string text)
        where TEnum : struct, Enum => throw new NotSupportedException();

    private void AddMaps()
    {
        this.AddMap(
            new Dictionary<string, AttributeUseType>
            {
                { "optional", AttributeUseType.Optional },
                { "required", AttributeUseType.Required },
            }
        );
        this.AddMap(
            new Dictionary<string, ProcessContents>
            {
                { "strict", ProcessContents.Strict },
                { "lax", ProcessContents.Lax },
            }
        );
    }

    private void AddMap<TEnum>(Dictionary<string, TEnum> enumMap)
        where TEnum : struct, Enum => this._textToEnumMaps.Add(typeof(TEnum), enumMap);
}

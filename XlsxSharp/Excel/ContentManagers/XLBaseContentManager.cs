using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace XlsxSharp.Excel.ContentManagers;

internal abstract class XLBaseContentManager<T>
    where T : struct, Enum
{
    protected readonly Dictionary<T, OpenXmlElement?> contents = new();

    public OpenXmlElement? GetPreviousElementFor(T content)
    {
        // JIT will recognize the conversion for identity and removes it (for int enums).
        int i = (int)(ValueType)content;

        IEnumerable<KeyValuePair<T, OpenXmlElement?>> previousElements = this.contents.Where(kv =>
            (int)(ValueType)kv.Key < i && kv.Value is not null
        );

        // If there is no previous element, return null.
        OpenXmlElement? previousElement = previousElements
            .DefaultIfEmpty(new KeyValuePair<T, OpenXmlElement?>(default, null))
            .MaxBy(kv => kv.Key)
            .Value;
        return previousElement;
    }

    public void SetElement(T content, OpenXmlElement? element)
    {
        this.contents[content] = element;
    }
}

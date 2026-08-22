using System;
using System.Globalization;

namespace XlsxSharp.Excel.IO;

internal class SequentialNameGenerator
{
    private readonly string _prefix;
    private int _nextNumber;

    internal SequentialNameGenerator(string prefix, int nextNumber)
    {
        this._prefix = prefix;
        this._nextNumber = nextNumber;
    }

    internal void AddName(string name)
    {
        if (!name.StartsWith(this._prefix))
        {
            return;
        }

        if (
            !int.TryParse(
                name[this._prefix.Length..],
                NumberStyles.None,
                XlsxSharp.XLHelper.ParseCulture,
                out int styleNumber
            )
        )
        {
            return;
        }

        this._nextNumber = Math.Max(styleNumber + 1, this._nextNumber);
    }

    internal string NextUnusedStyleName()
    {
        return $"{this._prefix}{this._nextNumber++}";
    }
}

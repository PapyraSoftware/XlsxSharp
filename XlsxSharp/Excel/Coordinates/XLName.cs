using System;
using System.Linq;

namespace XlsxSharp.Excel;

/// <summary>
/// A name in a worksheet. Unlike <see cref="IXLDefinedName"/>, this is basically only a reference.
/// The actual
/// </summary>
internal readonly struct XLName : IEquatable<XLName>
{
    /// <summary>
    /// Name of a sheet. If null, the scope is a workbook. The sheet might not exist, e.g. it
    /// is only in a formula. The name of a sheet is not escaped.
    /// </summary>
    public string? SheetName { get; }

    /// <summary>
    /// The defined name in the scope. Case insensitive during comparisons.
    /// </summary>
    public string Name { get; }

    public XLName(string sheetName, string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(sheetName);

        if (name.Any(char.IsWhiteSpace))
        {
            throw new ArgumentException("Name can't contain whitespace.");
        }

        this.SheetName = sheetName;
        this.Name = name;
    }

    public XLName(string name)
    {
        if (name.Any(char.IsWhiteSpace))
        {
            throw new ArgumentException("Name can't contain whitespace.");
        }

        this.SheetName = null;
        this.Name = name;
    }

    public bool Equals(XLName other)
    {
        bool differentScope = this.SheetName is null ^ other.SheetName is null;
        if (differentScope)
        {
            return false;
        }

        bool bothWorkbookScope = this.SheetName is null && other.SheetName is null;
        if (bothWorkbookScope)
        {
            return XlsxSharp.XLHelper.NameComparer.Equals(this.Name, other.Name);
        }

        return XlsxSharp.XLHelper.NameComparer.Equals(this.Name, other.Name)
            && XlsxSharp.XLHelper.SheetComparer.Equals(this.SheetName, other.SheetName);
    }

    public override bool Equals(object? obj) => obj is XLName other && this.Equals(other);

    public override int GetHashCode()
    {
        // Both parts are hashed through their comparer so that they match the case insensitive
        // Equals.
        int sheetHashCode =
            this.SheetName is not null
                ? XlsxSharp.XLHelper.SheetComparer.GetHashCode(this.SheetName)
                : 0;
        return HashCode.Combine(sheetHashCode, XlsxSharp.XLHelper.NameComparer.GetHashCode(this.Name));
    }

    public override string ToString()
    {
        bool isWorkbookScoped = this.SheetName is null;
        return isWorkbookScoped ? this.Name : $"{this.SheetName}!{this.Name}";
    }
}

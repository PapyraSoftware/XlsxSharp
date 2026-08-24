using System.Diagnostics.CodeAnalysis;

namespace XlsxSharp.Excel;

public interface IXLDefinedNames : IEnumerable<IXLDefinedName>
{
    /// <inheritdoc cref="DefinedName"/>
    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    public IXLDefinedName NamedRange(string name);

    /// <summary>
    /// Gets the specified defined name.
    /// </summary>
    /// <param name="name">Name identifier.</param>
    /// <exception cref="KeyNotFoundException">Name wasn't found.</exception>
    public IXLDefinedName DefinedName(string name);

    /// <summary>
    /// Adds a new defined name.
    /// </summary>
    /// <param name="name">Name identifier to add.</param>
    /// <param name="rangeAddress">The range address to add.</param>
    /// <exception cref="ArgumentException">The name or address is invalid.</exception>
    public IXLDefinedName Add(string name, string rangeAddress);

    /// <summary>
    /// Adds a new defined name.
    /// </summary>
    /// <param name="name">Name identifier to add.</param>
    /// <param name="range">The range to add.</param>
    /// <exception cref="ArgumentException">The name is invalid.</exception>
    public IXLDefinedName Add(string name, IXLRange range);

    /// <summary>
    /// Adds a new defined name.
    /// </summary>
    /// <param name="name">Name identifier to add.</param>
    /// <param name="ranges">The ranges to add.</param>
    /// <exception cref="ArgumentException">The name is invalid.</exception>
    public IXLDefinedName Add(string name, IXLRanges ranges);

    /// <summary>
    /// Adds a new defined name.
    /// </summary>
    /// <param name="name">Name identifier to add.</param>
    /// <param name="rangeAddress">The range address to add.</param>
    /// <param name="comment">The comment for the new named range.</param>
    /// <exception cref="ArgumentException">The range name or address is invalid.</exception>
    public IXLDefinedName Add(string name, string rangeAddress, string? comment);

    /// <summary>
    /// Adds a new defined name.
    /// </summary>
    /// <param name="name">Name identifier to add.</param>
    /// <param name="range">The range to add.</param>
    /// <param name="comment">The comment for the new named range.</param>
    /// <exception cref="ArgumentException">The range name is invalid.</exception>
    public IXLDefinedName Add(string name, IXLRange range, string? comment);

    /// <summary>
    /// Adds a new defined name.
    /// </summary>
    /// <param name="name">Name identifier to add.</param>
    /// <param name="ranges">The ranges to add.</param>
    /// <param name="comment">The comment for the new named range.</param>
    /// <exception cref="ArgumentException">The range name is invalid.</exception>
    public IXLDefinedName Add(string name, IXLRanges ranges, string? comment);

    /// <summary>
    /// Deletes the specified defined name.  Deleting defined name doesn't delete referenced
    /// cells.
    /// </summary>
    /// <param name="name">Name identifier to delete.</param>
    public void Delete(string name);

    /// <summary>
    /// Deletes the specified defined name's index. Deleting defined name doesn't delete
    /// referenced cells.
    /// </summary>
    /// <param name="index">Index of the defined name to delete.</param>
    /// <exception cref="ArgumentOutOfRangeException">The index is outside of named ranges array.</exception>
    public void Delete(int index);

    /// <summary>
    /// Deletes all defined names of this collection, i.e. a workbook or a sheet. Deleting
    /// defined name doesn't delete referenced cells.
    /// </summary>
    public void DeleteAll();

    public bool TryGetValue(string name, [NotNullWhen(true)] out IXLDefinedName? range);

    public bool Contains(string name);

    /// <summary>
    /// Returns a subset of defined names that do not have invalid references.
    /// </summary>
    public IEnumerable<IXLDefinedName> ValidNamedRanges();

    /// <summary>
    /// Returns a subset of defined names that do have invalid references.
    /// </summary>
    public IEnumerable<IXLDefinedName> InvalidNamedRanges();
}

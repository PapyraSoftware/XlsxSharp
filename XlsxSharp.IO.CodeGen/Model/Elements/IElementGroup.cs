using System.Collections.Generic;

namespace XlsxSharp.IO.CodeGen.Model.Elements;

/// <summary>
/// A node in a complex type element tree.
/// </summary>
public interface IElementGroup
{
    /// <summary>
    /// Children elements.
    /// </summary>
    public List<IElementGroup> Children { get; }
}

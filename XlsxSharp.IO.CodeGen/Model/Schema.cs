using System.Diagnostics.CodeAnalysis;
using XlsxSharp.IO.CodeGen.Model.TopLevel;

namespace XlsxSharp.IO.CodeGen.Model;

/// <summary>
/// A representation of a one XSD file.
/// </summary>
public class Schema
{
    /// <summary>
    /// Imports in the file.
    /// </summary>
    public List<ImportElement> Imports { get; } = [];

    /// <summary>
    /// One of <c>xsd:attributeGroup</c>, <c>xsd:complexType</c>, <c>xsd:element</c>, <c>xsd:group</c> or <c>xsd:simpleType</c>.
    /// </summary>
    public List<object> Entries { get; } = [];

    internal bool TryGetParslet(ParsletName parsletName, [NotNullWhen(true)] out IParslet? parslet)
    {
        parslet = this.Entries.OfType<IParslet>().SingleOrDefault(x => x.Name == parsletName);
        return parslet is not null;
    }
}

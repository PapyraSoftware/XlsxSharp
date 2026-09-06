using System.Xml.Linq;

namespace XlsxSharp.Excel.IO.Schemas;

/// <summary>
/// Removes the parts of a document that ECMA-376 Part 3 (Markup Compatibility and Extensibility)
/// says a processor may ignore, so that schema validation - which knows nothing about MCE - sees
/// only content the base ISO/IEC 29500 schemas actually define.
/// </summary>
/// <remarks>
/// A document mixes in exactly two things MCE defines: an element carrying an
/// <c>mc:Ignorable</c> attribute lists namespace prefixes whose elements and attributes, anywhere
/// in its subtree, a processor that does not understand them is free to ignore - which XlsxSharp's
/// own writers use for `x14ac:dyDescent`, and a loaded file can use for anything else. An
/// <c>mc:AlternateContent</c> wrapper offers a choice between alternatives for different
/// consumers, of which only the <c>mc:Fallback</c> - the one meant for a processor that
/// understands none of the extensions offered as `mc:Choice` - is something the base schemas
/// describe.
/// </remarks>
internal static class MarkupCompatibility
{
    private static readonly XNamespace Ns =
        "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>Strips <paramref name="root"/> and its descendants in place.</summary>
    internal static void Strip(XElement root) => Strip(root, []);

    private static void Strip(XElement element, HashSet<string> ignoredNamespaces)
    {
        if (element.Attribute(Ns + "Ignorable") is { } ignorable)
        {
            ignoredNamespaces = [.. ignoredNamespaces];
            foreach (
                string prefix in ignorable.Value.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            )
            {
                if (element.GetNamespaceOfPrefix(prefix) is { } ns)
                {
                    ignoredNamespaces.Add(ns.NamespaceName);
                }
            }

            ignorable.Remove();
        }

        element
            .Attributes()
            .Where(a =>
                !a.IsNamespaceDeclaration && ignoredNamespaces.Contains(a.Name.NamespaceName)
            )
            .ToList()
            .ForEach(a => a.Remove());

        foreach (XElement child in element.Elements().ToList())
        {
            if (child.Name == Ns + "AlternateContent")
            {
                StripAlternateContent(child, ignoredNamespaces);
                continue;
            }

            if (ignoredNamespaces.Contains(child.Name.NamespaceName))
            {
                child.Remove();
                continue;
            }

            Strip(child, ignoredNamespaces);

            // A CT_Extension <ext uri="..."> exists only to hold one extension element the base
            // schemas do not themselves define - its content model is a single required xsd:any -
            // so stripping its one child because it turned out to be ignorable leaves an <ext>
            // the schema now considers invalid for having no content at all, even though the
            // exact same file was schema-valid before an ignorable namespace's content was
            // removed from it. Dropping the now-empty wrapper too is what a processor that
            // recognises none of its content would actually end up with. The "uri" attribute is
            // what tells this apart from an unrelated element that also happens to be named
            // "ext", such as DrawingML's <xdr:ext cx="..." cy="..."/> anchor size, which is
            // legitimately childless by its own type and must not be removed.
            if (
                child.Name.LocalName == "ext"
                && !child.HasElements
                && child.Attribute("uri") is not null
            )
            {
                child.Remove();
            }
        }
    }

    /// <summary>
    /// A processor that understands none of the extensions <c>mc:Choice</c> offers - which, not
    /// modelling any of them, this validator never does - falls through to <c>mc:Fallback</c>: the
    /// wrapper is replaced by that element's own children, stripped the same way, or dropped
    /// entirely when there is no fallback to fall through to.
    /// </summary>
    private static void StripAlternateContent(
        XElement alternateContent,
        HashSet<string> ignoredNamespaces
    )
    {
        XElement? fallback = alternateContent.Element(Ns + "Fallback");
        if (fallback is null)
        {
            alternateContent.Remove();
            return;
        }

        List<XElement> children = [.. fallback.Elements()];
        alternateContent.ReplaceWith(children);
        children.ForEach(child => Strip(child, ignoredNamespaces));
    }
}

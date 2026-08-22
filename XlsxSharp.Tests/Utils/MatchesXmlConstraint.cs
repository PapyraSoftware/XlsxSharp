using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using NUnit.Framework.Constraints;

namespace XlsxSharp.Tests.Utils;

/// <summary>
/// Compare an element in an <see cref="XDocument"/> with the supplied XML.
/// </summary>
internal class MatchesXmlConstraint(string xml) : Constraint
{
    public override string Description => $"XML should semantically match {xml}.";

    public override ConstraintResult ApplyTo<TActual>(TActual actual)
    {
        if (actual is not IEnumerable<XElement> elements)
        {
            return new ConstraintResult(this, actual, ConstraintStatus.Failure);
        }

        XElement element = elements.Single();
        XDocument expected = XDocument.Load(new StringReader(xml));
        bool xmlEqual = element.SemanticallyEqual(expected.Root);
        return new ConstraintResult(
            this,
            element,
            xmlEqual ? ConstraintStatus.Success : ConstraintStatus.Error
        );
    }
}

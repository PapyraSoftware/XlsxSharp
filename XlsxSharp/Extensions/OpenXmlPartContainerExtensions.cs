#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace XlsxSharp.Extensions;

internal static class OpenXmlPartContainerExtensions
{
    public static Boolean HasPartWithId(this OpenXmlPartContainer container, String relId) =>
        container.Parts.Any(p => p.RelationshipId.Equals(relId));
}

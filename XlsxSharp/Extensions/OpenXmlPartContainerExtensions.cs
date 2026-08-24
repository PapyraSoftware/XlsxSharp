#nullable disable

using DocumentFormat.OpenXml.Packaging;

namespace XlsxSharp.Extensions;

internal static class OpenXmlPartContainerExtensions
{
    public static bool HasPartWithId(this OpenXmlPartContainer container, string relId) =>
        container.Parts.Any(p => p.RelationshipId.Equals(relId));
}

#nullable enable
using XlsxSharp.IO;

namespace XlsxSharp.Tests.IO;

internal interface IExpectedXmlNode
{
    void AssertMatches(IXmlReader reader);
}

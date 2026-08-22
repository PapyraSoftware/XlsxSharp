#nullable enable
using XlsxSharp.IO;

namespace XlsxSharp.Tests.IO;

internal interface IExpectedXmlNode
{
    public void AssertMatches(IXmlReader reader);
}

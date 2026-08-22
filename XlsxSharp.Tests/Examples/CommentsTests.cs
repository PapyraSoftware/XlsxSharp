using NUnit.Framework;
using XlsxSharp.Examples.Comments;

namespace XlsxSharp.Tests.Examples;

[TestFixture]
public class CommentsTests
{
    [Test]
    public void AddingComments()
    {
        TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx");
    }
}

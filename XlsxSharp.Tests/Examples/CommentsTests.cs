using XlsxSharp.Examples.Comments;

namespace XlsxSharp.Tests.Examples;

public class CommentsTests
{
    [Test]
    public void AddingComments() =>
        TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx");
}

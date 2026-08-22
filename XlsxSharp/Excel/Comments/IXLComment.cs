#nullable disable

using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.RichText;

namespace XlsxSharp.Excel.Comments;

public interface IXLComment : IXLFormattedText<IXLComment>, IXLDrawing<IXLComment>
{
    /// <summary>
    /// Gets or sets this comment's author's name
    /// </summary>
    public string Author { get; set; }

    /// <summary>
    /// Sets the name of the comment's author
    /// </summary>
    /// <param name="value">Author's name</param>
    public IXLComment SetAuthor(string value);

    /// <summary>
    /// Adds a bolded line with the author's name
    /// </summary>
    public IXLRichString AddSignature();

    public void Delete();
}

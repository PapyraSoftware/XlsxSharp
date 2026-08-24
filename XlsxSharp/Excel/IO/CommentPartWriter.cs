#nullable disable

using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Excel.Comments;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.IO.OpenXmlConst;

namespace XlsxSharp.Excel.IO;

internal class CommentPartWriter
{
    internal static void GenerateWorksheetCommentsPartContent(
        WorksheetCommentsPart worksheetCommentsPart,
        XLWorksheet xlWorksheet
    )
    {
        XmlWriterSettings settings = new()
        {
            CloseOutput = true,
            Encoding = XlsxSharp.XLHelper.NoBomUTF8,
        };
        Stream partStream = worksheetCommentsPart.GetStream(FileMode.Create);
        using XmlWriter xml = XmlWriter.Create(partStream, settings);

        List<XLCell> commentCells = [];
        Dictionary<string, int> authorsDict = new();
        xml.WriteStartElement("x", "comments", Main2006SsNs);
        foreach (XLCell c in xlWorksheet.Internals.CellsCollection.GetCells(c => c.HasComment))
        {
            string authorName = c.GetComment().Author;

            if (!authorsDict.TryGetValue(authorName, out int authorId))
            {
                authorId = authorsDict.Count;
                authorsDict.Add(authorName, authorId);
            }

            commentCells.Add(c);
        }

        xml.WriteStartElement("authors", Main2006SsNs);
        foreach (KeyValuePair<string, int> author in authorsDict)
        {
            xml.WriteElementString("author", Main2006SsNs, author.Key);
        }

        xml.WriteEndElement(); // authors

        char[] refBuffer = new char[10];
        xml.WriteStartElement("commentList", Main2006SsNs);
        foreach (XLCell commentCell in commentCells)
        {
            XLComment comment = commentCell.GetComment();
            xml.WriteStartElement("comment", Main2006SsNs);

            int refLen = commentCell.Point.Format(refBuffer);
            xml.WriteStartAttribute("ref");
            xml.WriteRaw(refBuffer, 0, refLen);
            xml.WriteEndAttribute(); // ref

            int authorId = authorsDict[comment.Author];
            xml.WriteAttribute("authorId", authorId);

            // Excel specifies @guid is optional if the workbook is not shared
            // Excel ignores the shapeId attribute.

            xml.WriteStartElement("text", Main2006SsNs);
            XLImmutableRichText richText = XLImmutableRichText.Create(comment);
            foreach (XLImmutableRichText.RichTextRun run in richText.Runs)
            {
                TextSerializer.WriteRun(xml, richText, run);
            }

            xml.WriteEndElement(); // text
            xml.WriteEndElement(); // comment
        }

        xml.WriteEndElement(); // commentList
        xml.WriteEndElement(); // comments

        xml.Close();
    }
}

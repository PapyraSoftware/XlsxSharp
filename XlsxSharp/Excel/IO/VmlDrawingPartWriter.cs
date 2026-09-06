#nullable disable

using System.Text;
using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Excel.Comments;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Drawings.Style;
using XlsxSharp.Extensions;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Writes the legacy VML that carries the shapes of cell comments.
/// </summary>
/// <remarks>
/// The namespace declarations are repeated on every shape instead of sitting on the root, and the
/// attributes come in the order the VML schema declares rather than the order they are set here.
/// Both are what the SDK produced and what Excel has been reading from XlsxSharp all along, so
/// the writer works with namespace handling switched off and writes the qualified names itself.
/// <see cref="XlsxSharp.Tests"/> records the output of this class; see CommentVmlOutputTests.
/// </remarks>
internal class VmlDrawingPartWriter
{
    private const string VmlNs = "urn:schemas-microsoft-com:vml";
    private const string OfficeNs = "urn:schemas-microsoft-com:office:office";
    private const string ExcelNs = "urn:schemas-microsoft-com:office:excel";

    internal static bool GenerateContent(OpcPart vmlDrawingPart, XLWorksheet xlWorksheet)
    {
        using MemoryStream ms = new();
        using (Stream readStream = vmlDrawingPart.GetReadStream())
        {
            XLWorkbook.CopyStream(readStream, ms);
        }

        using Stream stream = vmlDrawingPart.GetWriteStream();

        // Namespaces off: the qualified names and the xmlns attributes are written by hand,
        // so that they land in the same places the SDK put them.
        XmlTextWriter writer = new(stream, Encoding.UTF8) { Namespaces = false };

        writer.WriteStartElement("xml");

        WriteShapeType(writer);

        IEnumerable<XLCell> cellWithComments = xlWorksheet.Internals.CellsCollection.GetCells(c =>
            c.HasComment
        );

        bool hasAnyVmlElements = false;

        foreach (XLCell c in cellWithComments)
        {
            WriteCommentShape(writer, c);
            hasAnyVmlElements |= true;
        }

        if (ms.Length > 0)
        {
            ms.Position = 0;
            XDocument xdoc = XDocumentExtensions.Load(ms);
            xdoc.Root.Elements().ForEach(e => writer.WriteRaw(e.ToXmlString()));
            hasAnyVmlElements |= xdoc.Root.HasElements;
        }

        writer.WriteEndElement();
        writer.Flush();
        writer.Close();

        return hasAnyVmlElements;
    }

    /// <summary>
    /// The shape template every comment shape refers to through its type attribute. See
    /// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.shapetype - a
    /// shapetype is a shape that cannot itself reference another shapetype, and positioning
    /// attributes are not inherited from it.
    /// </summary>
    private static void WriteShapeType(XmlWriter writer)
    {
        writer.WriteStartElement("v:shapetype");
        writer.WriteAttributeString("id", XLConstants.Comment.ShapeTypeId);
        writer.WriteAttributeString("coordsize", "21600,21600");
        writer.WriteAttributeString("o:spt", "202");
        writer.WriteAttributeString("path", "m,l,21600r21600,l21600,xe");
        WriteVmlNamespaces(writer);

        writer.WriteStartElement("v:stroke");
        writer.WriteAttributeString("joinstyle", "miter");
        writer.WriteEndElement();

        writer.WriteStartElement("v:path");
        writer.WriteAttributeString("gradientshapeok", "true");
        writer.WriteAttributeString("o:connecttype", "rect");
        writer.WriteEndElement();

        writer.WriteEndElement();
    }

    private static void WriteCommentShape(XmlWriter writer, XLCell c)
    {
        XLComment comment = c.GetComment();
        IXLDrawingColorsAndLines colors = comment.Style.ColorsAndLines;

        writer.WriteStartElement("v:shape");
        writer.WriteAttributeString("id", string.Concat("_x0000_s", comment.ShapeId));
        writer.WriteAttributeString("style", GetCommentStyle(c));
        if (!string.IsNullOrWhiteSpace(comment.Style.Web.AlternateText))
        {
            writer.WriteAttributeString("alt", comment.Style.Web.AlternateText);
        }

        writer.WriteAttributeString(
            "o:insetmode",
            comment.Style.Margins.Automatic ? "auto" : "custom"
        );

        writer.WriteAttributeString("fillcolor", Hex(colors.FillColor));
        writer.WriteAttributeString("strokecolor", Hex(colors.LineColor));
        writer.WriteAttributeString(
            "strokeweight",
            string.Concat(colors.LineWeight.ToInvariantString(), "pt")
        );

        writer.WriteAttributeString("type", "#" + XLConstants.Comment.ShapeTypeId);
        WriteVmlNamespaces(writer);

        WriteFill(writer, colors);
        WriteStroke(writer, colors);

        writer.WriteStartElement("v:shadow");
        writer.WriteAttributeString("obscured", "true");
        writer.WriteAttributeString("color", "black");
        writer.WriteEndElement();

        writer.WriteStartElement("v:path");
        writer.WriteAttributeString("o:connecttype", "none");
        writer.WriteEndElement();

        WriteTextBox(writer, comment.Style);
        WriteClientData(writer, c, comment);

        writer.WriteEndElement();
    }

    private static void WriteFill(XmlWriter writer, IXLDrawingColorsAndLines colors)
    {
        writer.WriteStartElement("v:fill");
        if (colors.FillTransparency < 1)
        {
            writer.WriteAttributeString("opacity", Opacity(colors.FillTransparency));
        }

        writer.WriteAttributeString("color2", Hex(colors.FillColor));
        writer.WriteEndElement();
    }

    private static void WriteStroke(XmlWriter writer, IXLDrawingColorsAndLines colors)
    {
        XLDashStyle lineDash = colors.LineDash;

        writer.WriteStartElement("v:stroke");
        if (colors.LineTransparency < 1)
        {
            writer.WriteAttributeString("opacity", Opacity(colors.LineTransparency));
        }

        writer.WriteAttributeString("linestyle", LineStyle(colors.LineStyle));
        if (lineDash == XLDashStyle.RoundDot)
        {
            writer.WriteAttributeString("endcap", "round");
        }

        // Both dotted styles are the same dash pattern in VML.
        writer.WriteAttributeString(
            "dashstyle",
            lineDash is XLDashStyle.RoundDot or XLDashStyle.SquareDot
                ? "shortDot"
                : lineDash.ToString().ToCamel()
        );

        writer.WriteEndElement();
    }

    private static void WriteTextBox(XmlWriter writer, IXLDrawingStyle ds)
    {
        StringBuilder sb = new();
        IXLDrawingAlignment a = ds.Alignment;

        if (a.Direction == XLDrawingTextDirection.Context)
        {
            sb.Append("mso-direction-alt:auto;");
        }
        else if (a.Direction == XLDrawingTextDirection.RightToLeft)
        {
            sb.Append("direction:RTL;");
        }

        if (a.Orientation != XLDrawingTextOrientation.LeftToRight)
        {
            sb.Append("layout-flow:vertical;");
            if (a.Orientation == XLDrawingTextOrientation.BottomToTop)
            {
                sb.Append("mso-layout-flow-alt:bottom-to-top;");
            }
            else if (a.Orientation == XLDrawingTextOrientation.Vertical)
            {
                sb.Append("mso-layout-flow-alt:top-to-bottom;");
            }
        }

        if (a.AutomaticSize)
        {
            sb.Append("mso-fit-shape-to-text:t;");
        }

        writer.WriteStartElement("v:textbox");
        if (sb.Length > 0)
        {
            writer.WriteAttributeString("style", sb.ToString());
        }

        IXLDrawingMargins dm = ds.Margins;
        if (!dm.Automatic)
        {
            writer.WriteAttributeString(
                "inset",
                string.Concat(
                    dm.Left.ToInvariantString(),
                    "in,",
                    dm.Top.ToInvariantString(),
                    "in,",
                    dm.Right.ToInvariantString(),
                    "in,",
                    dm.Bottom.ToInvariantString(),
                    "in"
                )
            );
        }

        writer.WriteEndElement();
    }

    private static void WriteClientData(XmlWriter writer, XLCell c, XLComment comment)
    {
        writer.WriteStartElement("xvml:ClientData");
        writer.WriteAttributeString("ObjectType", "Note");
        writer.WriteAttributeString("xmlns:xvml", ExcelNs);

        // Both of these read backwards, and did before this was rewritten: an absolutely
        // positioned comment is the one that says it moves with cells.
        WriteText(
            writer,
            "xvml:MoveWithCells",
            comment.Style.Properties.Positioning == XLDrawingAnchor.Absolute ? "True" : "False"
        );

        WriteText(
            writer,
            "xvml:SizeWithCells",
            comment.Style.Properties.Positioning == XLDrawingAnchor.MoveAndSizeWithCells
                ? "False"
                : "True"
        );

        WriteText(writer, "xvml:Anchor", GetAnchor(c));
        WriteText(
            writer,
            "xvml:TextHAlign",
            comment.Style.Alignment.Horizontal.ToString().ToCamel()
        );

        WriteText(writer, "xvml:TextVAlign", comment.Style.Alignment.Vertical.ToString().ToCamel());
        WriteText(writer, "xvml:AutoFill", "False");
        WriteText(writer, "xvml:Row", (c.Address.RowNumber - 1).ToInvariantString());
        WriteText(writer, "xvml:Column", (c.Address.ColumnNumber - 1).ToInvariantString());
        WriteText(writer, "xvml:Locked", comment.Style.Protection.Locked ? "True" : "False");
        WriteText(writer, "xvml:LockText", comment.Style.Protection.LockText ? "True" : "False");
        WriteText(writer, "xvml:Visible", comment.Visible ? "True" : "False");

        writer.WriteEndElement();
    }

    private static void WriteText(XmlWriter writer, string qualifiedName, string text)
    {
        writer.WriteStartElement(qualifiedName);
        writer.WriteString(text);
        writer.WriteEndElement();
    }

    /// <summary>
    /// The two prefixes go on every shape, after its own attributes, because each shape used to
    /// be serialised as a document of its own.
    /// </summary>
    private static void WriteVmlNamespaces(XmlWriter writer)
    {
        writer.WriteAttributeString("xmlns:o", OfficeNs);
        writer.WriteAttributeString("xmlns:v", VmlNs);
    }

    private static string Hex(XLColor color) => "#" + color.Color.ToHex().Substring(2);

    private static string Opacity(double transparency) =>
        Math.Round(Convert.ToDouble(transparency), 2).ToInvariantString();

    private static string LineStyle(XLLineStyle lineStyle) =>
        lineStyle switch
        {
            XLLineStyle.Single => "single",
            XLLineStyle.ThinThin => "thinThin",
            XLLineStyle.ThinThick => "thinThick",
            XLLineStyle.ThickThin => "thickThin",
            XLLineStyle.ThickBetweenThin => "thickBetweenThin",
            _ => throw new ArgumentOutOfRangeException(nameof(lineStyle)),
        };

    private static string GetAnchor(XLCell cell)
    {
        XLComment c = cell.GetComment();
        double cWidth = c.Style.Size.Width;
        int fcNumber = c.Position.Column - 1;
        int fcOffset = Convert.ToInt32(c.Position.ColumnOffset * 7.5);
        double widthFromColumns =
            cell.Worksheet.Column(c.Position.Column).Width - c.Position.ColumnOffset;
        XLCell lastCell = cell.CellRight(c.Position.Column - cell.Address.ColumnNumber);
        while (widthFromColumns <= cWidth)
        {
            lastCell = lastCell.CellRight();
            widthFromColumns += lastCell.WorksheetColumn().Width;
        }

        int lcNumber = lastCell.WorksheetColumn().ColumnNumber() - 1;
        int lcOffset = Convert.ToInt32(
            (lastCell.WorksheetColumn().Width - (widthFromColumns - cWidth)) * 7.5
        );

        double cHeight = c.Style.Size.Height;
        int frNumber = c.Position.Row - 1;
        int frOffset = Convert.ToInt32(c.Position.RowOffset);
        double heightFromRows = cell.Worksheet.Row(c.Position.Row).Height - c.Position.RowOffset;
        lastCell = cell.CellBelow(c.Position.Row - cell.Address.RowNumber);
        while (heightFromRows <= cHeight)
        {
            lastCell = lastCell.CellBelow();
            heightFromRows += lastCell.WorksheetRow().Height;
        }

        int lrNumber = lastCell.WorksheetRow().RowNumber() - 1;
        int lrOffset = Convert.ToInt32(lastCell.WorksheetRow().Height - (heightFromRows - cHeight));

        return string.Concat(
            fcNumber,
            ", ",
            fcOffset,
            ", ",
            frNumber,
            ", ",
            frOffset,
            ", ",
            lcNumber,
            ", ",
            lcOffset,
            ", ",
            lrNumber,
            ", ",
            lrOffset
        );
    }

    private static string GetCommentStyle(XLCell cell)
    {
        XLComment c = cell.GetComment();
        StringBuilder sb = new("position:absolute; ");

        sb.Append("visibility:");
        sb.Append(c.Visible ? "visible" : "hidden");
        sb.Append(';');

        sb.Append("width:");
        sb.Append(Math.Round(c.Style.Size.Width * 7.5, 2).ToInvariantString());
        sb.Append("pt;");
        sb.Append("height:");
        sb.Append(Math.Round(c.Style.Size.Height, 2).ToInvariantString());
        sb.Append("pt;");

        sb.Append("z-index:");
        sb.Append(c.ZOrder.ToInvariantString());

        return sb.ToString();
    }
}

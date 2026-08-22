#nullable disable

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using XlsxSharp.Excel.Comments;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Drawings.Style;
using XlsxSharp.Extensions;
using Anchor = DocumentFormat.OpenXml.Vml.Spreadsheet.Anchor;
using Locked = DocumentFormat.OpenXml.Vml.Spreadsheet.Locked;
using Vml = DocumentFormat.OpenXml.Vml;

namespace XlsxSharp.Excel.IO;

internal class VmlDrawingPartWriter
{
    // Generates content of vmlDrawingPart1.
    internal static bool GenerateContent(VmlDrawingPart vmlDrawingPart, XLWorksheet xlWorksheet)
    {
        using (MemoryStream ms = new())
        using (Stream stream = vmlDrawingPart.GetStream(FileMode.OpenOrCreate))
        {
            XLWorkbook.CopyStream(stream, ms);
            stream.Position = 0;
            XmlTextWriter writer = new(stream, Encoding.UTF8);

            writer.WriteStartElement("xml");

            // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.shapetype?view=openxml-2.8.1#remarks
            // This element defines a shape template that can be used to create other shapes.
            // Shapetype is identical to the shape element(§14.1.2.19) except it cannot reference another shapetype element.
            // The type attribute shall not be used with shapetype.
            // Attributes defined in the shape override any that appear in the shapetype positioning attributes
            // (such as top, width, z-index, rotation, flip) are not passed to a shape from a shapetype.
            // To use this element, create a shapetype with a specific id attribute.
            // Then create a shape and reference the shapetype's id using the type attribute.
            new Vml.Shapetype(
                new Vml.Stroke { JoinStyle = Vml.StrokeJoinStyleValues.Miter },
                new Vml.Path
                {
                    AllowGradientShape = true,
                    ConnectionPointType = ConnectValues.Rectangle,
                }
            )
            {
                Id = XLConstants.Comment.ShapeTypeId,
                CoordinateSize = "21600,21600",
                OptionalNumber = 202,
                EdgePath = "m,l,21600r21600,l21600,xe",
            }.WriteTo(writer);

            IEnumerable<XLCell> cellWithComments = xlWorksheet.Internals.CellsCollection.GetCells(
                c => c.HasComment
            );

            bool hasAnyVmlElements = false;

            foreach (XLCell c in cellWithComments)
            {
                GenerateCommentShape(c).WriteTo(writer);
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
    }

    // VML Shape for Comment
    private static Vml.Shape GenerateCommentShape(XLCell c)
    {
        int rowNumber = c.Address.RowNumber;
        int columnNumber = c.Address.ColumnNumber;

        XLComment comment = c.GetComment();
        string shapeId = string.Concat("_x0000_s", comment.ShapeId);
        // Unique per cell (workbook?), e.g.: "_x0000_s1026"
        Anchor anchor = GetAnchor(c);
        Vml.TextBox textBox = GetTextBox(comment.Style);
        Vml.Fill fill = new()
        {
            Color2 = "#" + comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2),
        };
        if (comment.Style.ColorsAndLines.FillTransparency < 1)
        {
            fill.Opacity = Math.Round(
                    Convert.ToDouble(comment.Style.ColorsAndLines.FillTransparency),
                    2
                )
                .ToInvariantString();
        }

        Vml.Stroke stroke = GetStroke(c);
        Vml.Shape shape = new(
            fill,
            stroke,
            new Vml.Shadow { Color = "black", Obscured = true },
            new Vml.Path { ConnectionPointType = ConnectValues.None },
            textBox,
            new ClientData(
                new MoveWithCells(
                    comment.Style.Properties.Positioning == XLDrawingAnchor.Absolute
                        ? "True"
                        : "False"
                ), // Counterintuitive
                new ResizeWithCells(
                    comment.Style.Properties.Positioning == XLDrawingAnchor.MoveAndSizeWithCells
                        ? "False"
                        : "True"
                ), // Counterintuitive
                anchor,
                new HorizontalTextAlignment(
                    comment.Style.Alignment.Horizontal.ToString().ToCamel()
                ),
                new Vml.Spreadsheet.VerticalTextAlignment(
                    comment.Style.Alignment.Vertical.ToString().ToCamel()
                ),
                new AutoFill("False"),
                new CommentRowTarget { Text = (rowNumber - 1).ToInvariantString() },
                new CommentColumnTarget { Text = (columnNumber - 1).ToInvariantString() },
                new Locked(comment.Style.Protection.Locked ? "True" : "False"),
                new LockText(comment.Style.Protection.LockText ? "True" : "False"),
                new Visible(comment.Visible ? "True" : "False")
            )
            {
                ObjectType = ObjectValues.Note,
            }
        )
        {
            Id = shapeId,
            Type = "#" + XLConstants.Comment.ShapeTypeId,
            Style = GetCommentStyle(c),
            FillColor = "#" + comment.Style.ColorsAndLines.FillColor.Color.ToHex().Substring(2),
            StrokeColor = "#" + comment.Style.ColorsAndLines.LineColor.Color.ToHex().Substring(2),
            StrokeWeight = string.Concat(
                comment.Style.ColorsAndLines.LineWeight.ToInvariantString(),
                "pt"
            ),
            InsetMode = comment.Style.Margins.Automatic
                ? InsetMarginValues.Auto
                : InsetMarginValues.Custom,
        };
        if (!string.IsNullOrWhiteSpace(comment.Style.Web.AlternateText))
        {
            shape.Alternate = comment.Style.Web.AlternateText;
        }

        return shape;
    }

    private static Vml.Stroke GetStroke(XLCell c)
    {
        XLDashStyle lineDash = c.GetComment().Style.ColorsAndLines.LineDash;
        Vml.Stroke stroke = new()
        {
            LineStyle = c.GetComment().Style.ColorsAndLines.LineStyle.ToOpenXml(),
            DashStyle =
                lineDash == XLDashStyle.RoundDot || lineDash == XLDashStyle.SquareDot
                    ? "shortDot"
                    : lineDash.ToString().ToCamel(),
        };
        if (lineDash == XLDashStyle.RoundDot)
        {
            stroke.EndCap = Vml.StrokeEndCapValues.Round;
        }

        if (c.GetComment().Style.ColorsAndLines.LineTransparency < 1)
        {
            stroke.Opacity = Math.Round(
                    Convert.ToDouble(c.GetComment().Style.ColorsAndLines.LineTransparency),
                    2
                )
                .ToInvariantString();
        }

        return stroke;
    }

    private static Vml.TextBox GetTextBox(IXLDrawingStyle ds)
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

        Vml.TextBox tb = new();

        if (sb.Length > 0)
        {
            tb.Style = sb.ToString();
        }

        IXLDrawingMargins dm = ds.Margins;
        if (!dm.Automatic)
        {
            tb.Inset = string.Concat(
                dm.Left.ToInvariantString(),
                "in,",
                dm.Top.ToInvariantString(),
                "in,",
                dm.Right.ToInvariantString(),
                "in,",
                dm.Bottom.ToInvariantString(),
                "in"
            );
        }

        return tb;
    }

    private static Anchor GetAnchor(XLCell cell)
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

        double cHeight = c.Style.Size.Height; //c.Style.Size.Height * 72.0;
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
        return new Anchor
        {
            Text = string.Concat(
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
            ),
        };
    }

    private static StringValue GetCommentStyle(XLCell cell)
    {
        XLComment c = cell.GetComment();
        StringBuilder sb = new("position:absolute; ");

        sb.Append("visibility:");
        sb.Append(c.Visible ? "visible" : "hidden");
        sb.Append(";");

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

#nullable disable

using System.Xml.Linq;
using XlsxSharp.Excel.IO;
using XlsxSharp.IO;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel;

public partial class XLWorkbook
{
    /// <summary>
    /// Reading and writing <c>xl/drawings/drawingN.xml</c>, shared between load and save because
    /// both sides walk the same three anchor shapes.
    /// </summary>
    internal static class DrawingXml
    {
        internal static readonly XNamespace Xdr =
            "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        internal static readonly XNamespace A =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

        internal static XElement Read(OpcPart part)
        {
            using Stream stream = part.GetReadStream();
            return XDocument.Load(stream).Root
                ?? throw PartStructureException.ExpectedElementNotFound("wsDr");
        }

        /// <summary>
        /// The relationship id of the picture a <c>pic</c> anchor embeds, or null for any other
        /// kind of anchor (a shape, a text box, a connector) - the drawing part holds more than
        /// XlsxSharp's own picture model does, and only the anchor that carries a picture is
        /// this model's to touch.
        /// </summary>
        internal static string PictureRelId(XElement anchor) =>
            anchor
                .Element(Xdr + "pic")
                ?.Element(Xdr + "blipFill")
                ?.Element(A + "blip")
                ?.Attribute(SpreadsheetXml.Rel + "embed")
                ?.Value;

        internal static XElement PictureProperties(XElement anchor) =>
            anchor.Element(Xdr + "pic")?.Element(Xdr + "nvPicPr")?.Element(Xdr + "cNvPr");

        internal static XElement Extents(XElement anchor) =>
            anchor
                .Element(Xdr + "pic")
                ?.Element(Xdr + "spPr")
                ?.Element(A + "xfrm")
                ?.Element(A + "ext");
    }
}

#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XlsxSharp.Excel;

public partial class XLWorkbook
{
    internal static OpenXmlElement GetAnchorFromImageId(DrawingsPart drawingsPart, string relId)
    {
        IEnumerable<OpenXmlElement> matchingAnchor = drawingsPart.WorksheetDrawing.Where(wsdr =>
            wsdr.Descendants<Xdr.BlipFill>().Any(x => x?.Blip?.Embed?.Value.Equals(relId) ?? false)
        );
        return matchingAnchor.FirstOrDefault();
    }

    internal static NonVisualDrawingProperties GetPropertiesFromAnchor(OpenXmlElement anchor)
    {
        if (!IsAllowedAnchor(anchor))
        {
            return null;
        }

        // Maybe we should not restrict here, and just search for all NonVisualDrawingProperties in an anchor?
        OpenXmlCompositeElement shape =
            anchor.Descendants<Xdr.Picture>().Cast<OpenXmlCompositeElement>().FirstOrDefault()
            ?? anchor
                .Descendants<Xdr.ConnectionShape>()
                .Cast<OpenXmlCompositeElement>()
                .FirstOrDefault();

        if (shape == null)
        {
            return null;
        }

        return shape.Descendants<Xdr.NonVisualDrawingProperties>().FirstOrDefault();
    }

    internal static String GetImageRelIdFromAnchor(OpenXmlElement anchor)
    {
        if (!IsAllowedAnchor(anchor))
        {
            return null;
        }

        BlipFill blipFill = anchor.Descendants<Xdr.BlipFill>().FirstOrDefault();
        return blipFill?.Blip?.Embed?.Value;
    }

    private static bool IsAllowedAnchor(OpenXmlElement anchor)
    {
        Type[] allowedAnchorTypes =
        [
            typeof(AbsoluteAnchor),
            typeof(OneCellAnchor),
            typeof(TwoCellAnchor),
        ];
        return (allowedAnchorTypes.Any(t => t == anchor.GetType()));
    }
}

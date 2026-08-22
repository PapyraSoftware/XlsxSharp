#nullable disable

using System;
using System.Drawing;
using System.IO;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Utils;

namespace XlsxSharp.Graphics;

/// <summary>
/// Read info about PCX picture.
/// https://moddingwiki.shikadi.net/wiki/PCX_Format
/// </summary>
internal class PcxInfoReader : ImageInfoReader
{
    protected override bool CheckHeader(Stream stream)
    {
        Span<byte> header = stackalloc byte[3];
        if (stream.Read(header) != header.Length)
        {
            return false;
        }

        return header[0] == 0xA
            && header[1] <= 5
            && // version must be 0..5
            header[2] <= 1; // encoding, nearly always should be 1
    }

    protected override XLPictureInfo ReadInfo(Stream stream)
    {
        stream.Position += 4;
        ushort winXMin = stream.ReadU16LE();
        ushort winYMin = stream.ReadU16LE();
        ushort winXMax = stream.ReadU16LE();
        ushort winYMax = stream.ReadU16LE();
        ushort dpiX = stream.ReadU16LE();
        ushort dpiY = stream.ReadU16LE();

        int widthPx = winXMax - winXMin + 1;
        int heightPx = winYMax - winYMin + 1;
        return new XLPictureInfo(
            XLPictureFormat.Pcx,
            new Size(widthPx, heightPx),
            Size.Empty,
            dpiX,
            dpiY
        );
    }
}

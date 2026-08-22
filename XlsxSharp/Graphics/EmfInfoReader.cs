#nullable disable

using System.Drawing;
using System.IO;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Utils;

namespace XlsxSharp.Graphics;

/// <summary>
/// Metadata read of a vector EMF file. Specification: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/
/// </summary>
internal class EmfInfoReader : ImageInfoReader
{
    private const uint EmfSignature = 0x464D4520; // ' EMF'

    protected override bool CheckHeader(Stream stream)
    {
        if (!stream.TryReadU32LE(out uint type) || type != 0x1)
        {
            return false;
        }

        stream.Position += 36;
        if (!stream.TryReadU32LE(out uint signature) || signature != EmfSignature)
        {
            return false;
        }

        stream.Position += 14;
        if (!stream.TryReadU16LE(out ushort reserved) || reserved != 0x0)
        {
            return false;
        }

        return true;
    }

    protected override XLPictureInfo ReadInfo(Stream stream)
    {
        stream.Position += 24;
        Rectangle frame = ReadRectL(stream);
        return new XLPictureInfo(XLPictureFormat.Emf, Size.Empty, frame.Size);
    }

    private static Rectangle ReadRectL(Stream stream)
    {
        int left = stream.ReadS32LE();
        int top = stream.ReadS32LE();
        int right = stream.ReadS32LE();
        int bottom = stream.ReadS32LE();
        return new Rectangle(left, top, right - left, bottom - top);
    }
}

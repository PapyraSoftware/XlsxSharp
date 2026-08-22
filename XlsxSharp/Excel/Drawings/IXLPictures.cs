#nullable disable

using System.Collections.Generic;
using System.IO;

namespace XlsxSharp.Excel.Drawings;

public interface IXLPictures : IEnumerable<IXLPicture>
{
    public int Count { get; }

    public IXLPicture Add(Stream stream);

    public IXLPicture Add(Stream stream, string name);

    public IXLPicture Add(Stream stream, XLPictureFormat format);

    public IXLPicture Add(Stream stream, XLPictureFormat format, string name);

    public IXLPicture Add(string imageFile);

    public IXLPicture Add(string imageFile, string name);

    public bool Contains(string pictureName);

    public void Delete(string pictureName);

    public void Delete(IXLPicture picture);

    public IXLPicture Picture(string pictureName);

    public bool TryGetPicture(string pictureName, out IXLPicture picture);
}

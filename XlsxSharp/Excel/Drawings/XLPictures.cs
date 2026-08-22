using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace XlsxSharp.Excel.Drawings;

internal class XLPictures : IXLPictures, IEnumerable<XLPicture>
{
    private readonly List<XLPicture> _pictures = [];
    private readonly XLWorksheet _worksheet;

    public XLPictures(XLWorksheet worksheet)
    {
        this._worksheet = worksheet;
        this.Deleted = (HashSet<string>)[];
    }

    public int Count
    {
        [DebuggerStepThrough]
        get => this._pictures.Count;
    }

    internal ICollection<String> Deleted { get; private set; }

    public IXLPicture Add(Stream stream)
    {
        XLPicture picture = new(this._worksheet, stream);
        this._pictures.Add(picture);
        picture.Name = this.GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(Stream stream, string name)
    {
        IXLPicture picture = this.Add(stream);
        picture.Name = name;
        return picture;
    }

    public IXLPicture Add(Stream stream, XLPictureFormat format)
    {
        XLPicture picture = new(this._worksheet, stream, format);
        this._pictures.Add(picture);
        picture.Name = this.GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(Stream stream, XLPictureFormat format, string name)
    {
        IXLPicture picture = this.Add(stream, format);
        picture.Name = name;
        return picture;
    }

    public IXLPicture Add(string imageFile)
    {
        using (FileStream fs = File.OpenRead(imageFile))
        {
            XLPicture picture = new(this._worksheet, fs);
            this._pictures.Add(picture);
            picture.Name = this.GetNextPictureName();
            return picture;
        }
    }

    public IXLPicture Add(string imageFile, string name)
    {
        IXLPicture picture = this.Add(imageFile);
        picture.Name = name;
        return picture;
    }

    public bool Contains(string pictureName) =>
        this._pictures.Any(p =>
            string.Equals(p.Name, pictureName, StringComparison.OrdinalIgnoreCase)
        );

    public void Delete(IXLPicture picture) => this.Delete(picture.Name);

    public void Delete(string pictureName)
    {
        List<XLPicture> picturesToDelete =
        [
            .. this._pictures.Where(picture =>
                picture.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase)
            ),
        ];

        if (!picturesToDelete.Any())
        {
            throw new ArgumentOutOfRangeException(
                nameof(pictureName),
                $"Picture {pictureName} was not found."
            );
        }

        foreach (XLPicture picture in picturesToDelete)
        {
            if (!string.IsNullOrEmpty(picture.RelId))
            {
                this.Deleted.Add(picture.RelId);
            }

            this._pictures.Remove(picture);
        }
    }

    IEnumerator<IXLPicture> IEnumerable<IXLPicture>.GetEnumerator() =>
        this._pictures.Cast<IXLPicture>().GetEnumerator();

    public IEnumerator<XLPicture> GetEnumerator() =>
        ((IEnumerable<XLPicture>)this._pictures).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IXLPicture Picture(string pictureName)
    {
        if (this.TryGetPicture(pictureName, out IXLPicture? p))
        {
            return p!;
        }

        throw new ArgumentOutOfRangeException(
            nameof(pictureName),
            $"Picture {pictureName} was not found."
        );
    }

    public bool TryGetPicture(string pictureName, out IXLPicture? picture)
    {
        XLPicture? match = this._pictures.FirstOrDefault(p =>
            p.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase)
        );
        if (match is not null)
        {
            picture = match;
            return true;
        }
        picture = null;
        return false;
    }

    internal IXLPicture Add(Stream stream, string name, int Id)
    {
        XLPicture picture = (XLPicture)this.Add(stream);
        picture.SetName(name);
        picture.Id = Id;
        return picture;
    }

    private String GetNextPictureName()
    {
        int pictureNumber = this.Count;
        while (this._pictures.Any(p => p.Name == $"Picture {pictureNumber}"))
        {
            pictureNumber++;
        }
        return $"Picture {pictureNumber}";
    }
}

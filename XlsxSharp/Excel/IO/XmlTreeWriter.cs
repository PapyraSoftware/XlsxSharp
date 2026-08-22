using System;
using System.Xml;
using XlsxSharp.Extensions;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

internal sealed class XmlTreeWriter : IDisposable
{
    private readonly XmlWriter _xml;
    private readonly IEnumMapper _enumMapper;
    private bool _disposed;

    internal XmlTreeWriter(XmlWriter xml, IEnumMapper enumMapper)
    {
        this._xml = xml;
        this._enumMapper = enumMapper;
    }

    public void WriteStartDocument(string rootElementName, string ns)
    {
        this.ThrowIfDisposed();

        // No part should rely on external DTD, plus Excel also writes standalone="yes"
        this._xml.WriteStartDocument(standalone: true);

        // Make root element ns a default namespace to avoid prefix if possible
        this._xml.WriteStartElement(rootElementName, ns);
        this._xml.WriteAttributeString("xmlns", ns);
    }

    public void WriteStartElement(string localName, string ns)
    {
        this.ThrowIfDisposed();
        this._xml.WriteStartElement(localName, ns);
    }

    public void WriteStartExtension(string extUri, string defaultNs, string nsPrefix, string extNs)
    {
        this.ThrowIfDisposed();
        this.WriteStartElement("ext", defaultNs);
        this.WriteAttribute("uri", extUri);
        this.WriteNsPrefix(nsPrefix, extNs);
    }

    public void WriteNsPrefix(string prefix, string ns)
    {
        this.ThrowIfDisposed();
        this._xml.WriteAttributeString("xmlns", prefix, null, ns);
    }

    public void WriteAttribute(string attributeName, int value)
    {
        this.ThrowIfDisposed();
        this._xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, bool value)
    {
        this.ThrowIfDisposed();
        this._xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, string value)
    {
        this.ThrowIfDisposed();
        this._xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute(string attributeName, double value)
    {
        this.ThrowIfDisposed();
        this._xml.WriteAttribute(attributeName, value);
    }

    public void WriteAttribute<TEnum>(string attributeName, TEnum value)
        where TEnum : struct, Enum
    {
        this.ThrowIfDisposed();
        if (!this._enumMapper.TryGetText(value, out string text))
        {
            throw new InvalidOperationException(
                $"Missing mapping for enum {value} ({typeof(TEnum).Name})."
            );
        }

        this._xml.WriteAttribute(attributeName, text);
    }

    public void WriteEndElement()
    {
        this.ThrowIfDisposed();
        this._xml.WriteEndElement();
    }

    public void WriteEndDocument()
    {
        this.ThrowIfDisposed();
        this._xml.WriteEndElement();
        this._xml.WriteEndDocument();
    }

    public void Dispose()
    {
        this._xml.Dispose();
        this._disposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (this._disposed)
        {
            throw new ObjectDisposedException(nameof(XmlTreeWriter));
        }
    }
}

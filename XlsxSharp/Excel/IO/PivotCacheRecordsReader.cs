using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

internal partial class PivotCacheRecordsReader
{
    private readonly string _ns = OpenXmlConst.Main2006SsNs;
    private readonly XmlTreeReader _reader;
    private readonly XLPivotCache _pivotCache;

    /// <summary>
    /// Index of current field that is read from the <c>r</c> element.
    /// </summary>
    private int _fieldIdx;

    public PivotCacheRecordsReader(XmlTreeReader reader, XLPivotCache pivotCache)
    {
        this._reader = reader;
        this._pivotCache = pivotCache;
    }

    internal void ReadRecordsToCache()
    {
        // Don't add values to the shared items of a cache when record value is added, because we want 1:1
        // read/write. Read them from definition. Whatever is in shared items now should be written out,
        // unless there is a cache refresh. Basically trust the author of the workbook that it is valid.
        this._reader.Open("pivotCacheRecords", this._ns);
        int recordCount = this._reader.GetCount();
        this._pivotCache.AllocateRecordCapacity(recordCount);

        while (ParseRecord("r", this._ns) is { IsSuccess: true })
        {
            ;
        }

        if (this._reader.TryOpen("extLst", this._ns))
        {
            this._reader.Skip("extLst");
        }

        this._reader.Close("pivotCacheRecords", this._ns);
    }

    partial void OnRecordParsed()
    {
        // Each record should have element for each field
        int fieldsCount = this._pivotCache.FieldCount;
        if (this._fieldIdx != fieldsCount)
        {
            throw PartStructureException.IncorrectElementsCount();
        }

        // Record was read, reset field index for next record.
        this._fieldIdx = 0;
    }

    partial void OnMissingParsed(
        bool? u,
        bool? f,
        string? c,
        uint? cp,
        uint? @in,
        uint? bc,
        uint? fc,
        bool i,
        bool un,
        bool st,
        bool b
    )
    {
        XLPivotCacheValues fieldValues = this.GetFieldValues();
        fieldValues.AddMissing();
    }

    partial void OnNumberParsed(
        double v,
        bool? u,
        bool? f,
        string? c,
        uint? cp,
        uint? @in,
        uint? bc,
        uint? fc,
        bool i,
        bool un,
        bool st,
        bool b
    )
    {
        XLPivotCacheValues fieldValues = this.GetFieldValues();
        fieldValues.AddNumber(v);
    }

    partial void OnBooleanParsed(bool v, bool? u, bool? f, string? c, uint? cp)
    {
        XLPivotCacheValues fieldValues = this.GetFieldValues();
        fieldValues.AddBoolean(v);
    }

    partial void OnErrorParsed(
        string v,
        bool? u,
        bool? f,
        string? c,
        uint? cp,
        uint? @in,
        uint? bc,
        uint? fc,
        bool i,
        bool un,
        bool st,
        bool b
    )
    {
        XLPivotCacheValues fieldValues = this.GetFieldValues();
        if (!XLErrorParser.TryParseError(v, out XLError error))
        {
            throw PartStructureException.InvalidAttributeFormat();
        }

        fieldValues.AddError(error);
    }

    partial void OnStringParsed(
        string v,
        bool? u,
        bool? f,
        string? c,
        uint? cp,
        uint? @in,
        uint? bc,
        uint? fc,
        bool i,
        bool un,
        bool st,
        bool b
    )
    {
        XLPivotCacheValues fieldValues = this.GetFieldValues();
        fieldValues.AddString(v);
    }

    partial void OnDateTimeParsed(DateTime v, bool? u, bool? f, string? c, uint? cp)
    {
        XLPivotCacheValues fieldValues = this.GetFieldValues();
        fieldValues.AddDateTime(v);
    }

    partial void OnIndexParsed(uint v)
    {
        XLPivotCacheValues fieldValues = this.GetFieldValues();
        if (v >= fieldValues.SharedCount)
        {
            throw PartStructureException.InvalidAttributeValue();
        }

        fieldValues.AddIndex(v);
    }

    private XLPivotCacheValues GetFieldValues()
    {
        if (this._fieldIdx >= this._pivotCache.FieldCount)
        {
            throw PartStructureException.IncorrectElementsCount();
        }

        return this._pivotCache.GetFieldValues(this._fieldIdx++);
    }
}

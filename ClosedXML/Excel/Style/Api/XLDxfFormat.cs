using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

internal partial class XLDxFormat
{
    private readonly XLWorkbookStyles _styles;
    private readonly IXLDxfContainer _container;

    internal XLDxFormat(XLWorkbookStyles styles, IXLDxfContainer container)
    {
        _styles = styles;
        _container = container;
    }

    private XLDxfValue Dxf => _container.FormatValue ?? XLDxfValue.Empty;

    private XLDxfNumberFormat NumberFormat => new(this);

    private XLDxfAlignmentFormat Alignment => new(this);

    private XLDxfFontFormat Font => new(this);

    private XLDxfFillFormat Fill => new(this);

    private XLDxfBorderFormat Border => new(this);

    internal TProperty? Resolve<TComponent, TProperty>(Func<XLDxfValue, TComponent> getComponent, Func<TComponent, TProperty?> getProperty)
        where TProperty : struct
    {
        var component = getComponent(Dxf);
        return getProperty(component);
    }

    internal TProperty? Resolve<TComponent, TProperty>(Func<XLDxfValue, TComponent> getComponent, Func<TComponent, TProperty?> getProperty)
        where TProperty : class
    {
        var component = getComponent(Dxf);
        return getProperty(component);
    }

    internal void ModifyNumberFormat(XLNumberFormat numberFormat)
    {
        _container.FormatValue = _styles.RegisterDxFormat(Dxf with { NumberFormat = numberFormat }); ;
    }

    internal void ModifyFont<T>(Func<XLDifferentialFontValue, T, XLDifferentialFontValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedFont = modify(dxf.Font, value);
            var modifiedDxf = dxf with { Font = modifiedFont };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void ModifyFill<T>(Func<XLDifferentialFillValue, T, XLDifferentialFillValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedFill = modify(dxf.Fill, value);
            var modifiedDxf = dxf with { Fill = modifiedFill };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void ModifyAlignment<T>(Func<XLDifferentialAlignmentValue, T, XLDifferentialAlignmentValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedAlignment = modify(dxf.Alignment, value);
            var modifiedDxf = dxf with { Alignment = modifiedAlignment };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void ModifyBorder<T>(Func<XLDifferentialBorderValue, T, XLDifferentialBorderValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedBorder = modify(dxf.Border, value);
            var modifiedDxf = dxf with { Border = modifiedBorder };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    /// <summary>
    /// A helper method that is used when a style if copied from one object to another.
    /// For example, <c>conditionaFormat.Style = someOtherApi.Style</c>.
    /// </summary>
    internal void SetStyle(IXLStyle other)
    {
        if (other is not XLDxFormat otherFormat)
            throw new NotSupportedException("Can only copy from dxf.");

        _container.FormatValue = _styles.RegisterDxFormat(otherFormat.Dxf);
    }
}

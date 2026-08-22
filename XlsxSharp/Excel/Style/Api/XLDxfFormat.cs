using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal partial class XLDxFormat
{
    private readonly XLWorkbookStyles _styles;
    private readonly IXLDxfContainer _container;

    internal XLDxFormat(XLWorkbookStyles styles, IXLDxfContainer container)
    {
        this._styles = styles;
        this._container = container;
    }

    private XLDxfValue Dxf => this._container.FormatValue ?? XLDxfValue.Empty;

    private XLDxfNumberFormat NumberFormat => new(this);

    private XLDxfAlignmentFormat Alignment => new(this);

    private XLDxfFontFormat Font => new(this);

    private XLDxfFillFormat Fill => new(this);

    private XLDxfBorderFormat Border => new(this);

    private XLDxfProtectionFormat Protection => new(this);

    internal TProperty? Resolve<TComponent, TProperty>(
        Func<XLDxfValue, TComponent> getComponent,
        Func<TComponent, TProperty?> getProperty
    )
        where TProperty : struct
    {
        TComponent component = getComponent(this.Dxf);
        return getProperty(component);
    }

    internal TProperty? Resolve<TComponent, TProperty>(
        Func<XLDxfValue, TComponent> getComponent,
        Func<TComponent, TProperty?> getProperty
    )
        where TProperty : class
    {
        TComponent component = getComponent(this.Dxf);
        return getProperty(component);
    }

    internal void ModifyNumberFormat(XLNumberFormat numberFormat) =>
        this._container.FormatValue = this._styles.RegisterDxFormat(
            this.Dxf with
            {
                NumberFormat = numberFormat,
            }
        );

    internal void ModifyFont<T>(
        Func<XLDifferentialFontValue, T, XLDifferentialFontValue> modify,
        T value
    )
    {
        XLDxfValue modifiedDxf = this._styles.GetRegisteredDxFormat(
            this.Dxf,
            dxf =>
            {
                XLDifferentialFontValue modifiedFont = modify(dxf.Font, value);
                XLDxfValue modifiedDxf = dxf with { Font = modifiedFont };
                return modifiedDxf;
            }
        );
        this._container.FormatValue = modifiedDxf;
    }

    internal void ModifyFill<T>(
        Func<XLDifferentialFillValue, T, XLDifferentialFillValue> modify,
        T value
    )
    {
        XLDxfValue modifiedDxf = this._styles.GetRegisteredDxFormat(
            this.Dxf,
            dxf =>
            {
                XLDifferentialFillValue modifiedFill = modify(dxf.Fill, value);
                XLDxfValue modifiedDxf = dxf with { Fill = modifiedFill };
                return modifiedDxf;
            }
        );
        this._container.FormatValue = modifiedDxf;
    }

    internal void ModifyAlignment<T>(
        Func<XLDifferentialAlignmentValue, T, XLDifferentialAlignmentValue> modify,
        T value
    )
    {
        XLDxfValue modifiedDxf = this._styles.GetRegisteredDxFormat(
            this.Dxf,
            dxf =>
            {
                XLDifferentialAlignmentValue modifiedAlignment = modify(dxf.Alignment, value);
                XLDxfValue modifiedDxf = dxf with { Alignment = modifiedAlignment };
                return modifiedDxf;
            }
        );
        this._container.FormatValue = modifiedDxf;
    }

    internal void ModifyBorder<T>(
        Func<XLDifferentialBorderValue, T, XLDifferentialBorderValue> modify,
        T value
    )
    {
        XLDxfValue modifiedDxf = this._styles.GetRegisteredDxFormat(
            this.Dxf,
            dxf =>
            {
                XLDifferentialBorderValue modifiedBorder = modify(dxf.Border, value);
                XLDxfValue modifiedDxf = dxf with { Border = modifiedBorder };
                return modifiedDxf;
            }
        );
        this._container.FormatValue = modifiedDxf;
    }

    internal void ModifyProtection<T>(
        Func<XLDifferentialProtectionValue, T, XLDifferentialProtectionValue> modify,
        T value
    )
    {
        XLDxfValue modifiedDxf = this._styles.GetRegisteredDxFormat(
            this.Dxf,
            dxf =>
            {
                XLDifferentialProtectionValue modifiedProtection = modify(dxf.Protection, value);
                XLDxfValue modifiedDxf = dxf with { Protection = modifiedProtection };
                return modifiedDxf;
            }
        );
        this._container.FormatValue = modifiedDxf;
    }

    /// <summary>
    /// A helper method that is used when a style if copied from one object to another.
    /// For example, <c>conditionaFormat.Style = someOtherApi.Style</c>.
    /// </summary>
    internal void SetStyle(IXLStyle other)
    {
        if (other is not XLDxFormat otherFormat)
        {
            throw new NotSupportedException("Can only copy from dxf.");
        }

        this._container.FormatValue = this._styles.RegisterDxFormat(otherFormat.Dxf);
    }
}

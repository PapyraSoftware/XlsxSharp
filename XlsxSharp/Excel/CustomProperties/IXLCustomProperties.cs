#nullable disable

namespace XlsxSharp.Excel.CustomProperties;

public interface IXLCustomProperties : IEnumerable<IXLCustomProperty>
{
    public void Add(IXLCustomProperty customProperty);
    public void Add<T>(string name, T value);
    public void Delete(string name);
    public IXLCustomProperty CustomProperty(string name);
}

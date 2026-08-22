using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class FormulasWithEvaluation : Formulas
{
    public override void Create(string filePath)
    {
        base.Create(filePath);
        using (XLWorkbook wb = new(filePath))
        {
            wb.Save(true, true);
        }
    }
}

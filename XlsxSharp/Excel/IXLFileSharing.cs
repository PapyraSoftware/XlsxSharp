// Keep this file CodeMaid organised and cleaned

namespace XlsxSharp.Excel;

public interface IXLFileSharing
{
    //String AlgorithmName { get; set; }
    //Byte[] HashValue { get; set; }
    public bool ReadOnlyRecommended { get; set; }

    //Byte[] ReservationPassword { get; set; }
    //Byte[] SaltValue { get; set; }
    //Int32 SpinCount { get; set; }
    public string? UserName { get; set; }
}

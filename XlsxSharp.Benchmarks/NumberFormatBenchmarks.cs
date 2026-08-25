using System.Globalization;
using BenchmarkDotNet.Attributes;
using OursNumberFormat = XlsxSharp.ExcelNumberFormat.NumberFormat;
using UpstreamNumberFormat = ExcelNumberFormat.NumberFormat;

namespace XlsxSharp.Benchmarks;

// Compares XlsxSharp.ExcelNumberFormat (our fork, reworked so the tokenizer/parser slice
// ReadOnlyMemory<char> out of the original format string instead of allocating a new System.String
// per token) against ExcelNumberFormat 1.1.0, the real upstream NuGet package this project was
// originally vendored from - see XlsxSharp.ExcelNumberFormat/third-party.txt.
//
// Both contenders are exercised over the same corpus of real format strings (pulled verbatim from
// XlsxSharp.Tests/ExcelNumberFormat/ExcelNumberFormatTests.cs) covering every SectionType: number,
// date, duration, fraction, exponential, text, conditional, colored and locale-tagged.
//
//  - NumberFormatParseBenchmarks: constructs a NumberFormat from each string. This is the code path
//    the span rewrite actually touched, so [MemoryDiagnoser] should show the allocation delta
//    directly.
//  - NumberFormatFormatBenchmarks: formats a fixed value through each pre-parsed NumberFormat. The
//    formatting hot path itself was deliberately left untouched by the span rewrite (it just reads
//    tokens in a different representation now), so this is mainly a check that nothing regressed.
internal static class NumberFormatCorpus
{
    public static readonly string[] FormatStrings =
    [
        "General",
        "0",
        "0.00",
        "#,##0",
        "#,##0.00",
        "0%",
        "0.00%",
        "\"$\"#,##0.00",
        "\"€\"#,##0.00",
        "##0.0E+0",
        "0.00E+00",
        "0.00000000E+00",
        "# ?/?",
        "# ??/??",
        "#\\ ?/4",
        "#\\ ??/??",
        "#\\ ??/100",
        "d-mmm-yy",
        "dd/mm/yyyy",
        "yyyy-mm-dd",
        "m/d/yyyy;@",
        "h:mm AM/PM",
        "hh:mm:ss",
        "dddd, mmmm dd, yyyy",
        "[h]:mm:ss",
        "[hh]",
        "mm:ss.0;@",
        "\"Yes\";\"Yes\";\"No\";@",
        "\"True\";\"True\";\"False\";@",
        "[Red]#.##",
        "[Blue]General",
        "@",
        "0;[Red]0",
        "#,##0.00;[Red](#,##0.00)",
        "#,##0.00;(#,##0.00);0.00",
        "[$-409]d\\-mmm\\-yy;@",
        "[$-1010409]dddd, mmmm dd, yyyy",
        "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        "0.00_);[Red]\\(0.00\\)",
        "yyyy\\-mm\\-dd\\Thh:mm",
        "mmm-yy",
        "AM/PMh\"時\"mm\"分\"ss\"秒\";@",
        "0.000000000",
        "00000\\-0000",
        "[=0]?;#,##0.00",
        "[>999999]#,,\"M\";[>999]#,\"K\";#",
        "[$RD$-1C0A]#,##0.00;[Red]\\-[$RD$-1C0A]#,##0.00",
        "\\$0.00",
        "yy/mm/dd",
        "[$-40E]h\\ \"óra\"\\ m\\ \"perckor\"\\ AM/PM;@",
        "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
    ];
}

[MemoryDiagnoser]
public class NumberFormatParseBenchmarks
{
    [Benchmark(Description = "XlsxSharp.ExcelNumberFormat (spans)")]
    public int ParseXlsxSharp()
    {
        int valid = 0;
        foreach (string format in NumberFormatCorpus.FormatStrings)
        {
            if (new OursNumberFormat(format).IsValid)
            {
                valid++;
            }
        }

        return valid;
    }

    [Benchmark(Baseline = true, Description = "ExcelNumberFormat (upstream)")]
    public int ParseUpstream()
    {
        int valid = 0;
        foreach (string format in NumberFormatCorpus.FormatStrings)
        {
            if (new UpstreamNumberFormat(format).IsValid)
            {
                valid++;
            }
        }

        return valid;
    }
}

[MemoryDiagnoser]
public class NumberFormatFormatBenchmarks
{
    private const double Value = 1234.567;
    private static readonly CultureInfo Culture = CultureInfo.InvariantCulture;

    private OursNumberFormat[] _ours = [];
    private UpstreamNumberFormat[] _upstream = [];

    [GlobalSetup]
    public void GlobalSetup()
    {
        this._ours = NumberFormatCorpus
            .FormatStrings.Select(f => new OursNumberFormat(f))
            .Where(f => f.IsValid)
            .ToArray();
        this._upstream = NumberFormatCorpus
            .FormatStrings.Select(f => new UpstreamNumberFormat(f))
            .Where(f => f.IsValid)
            .ToArray();
    }

    [Benchmark(Description = "XlsxSharp.ExcelNumberFormat (spans)")]
    public int FormatXlsxSharp()
    {
        int length = 0;
        foreach (OursNumberFormat format in this._ours)
        {
            length += format.Format(Value, Culture).Length;
        }

        return length;
    }

    [Benchmark(Baseline = true, Description = "ExcelNumberFormat (upstream)")]
    public int FormatUpstream()
    {
        int length = 0;
        foreach (UpstreamNumberFormat format in this._upstream)
        {
            length += format.Format(Value, Culture).Length;
        }

        return length;
    }
}

using System.Diagnostics;
using Antlr4.Runtime;

using FileStream file = File.OpenRead(@"c:\Temp\formulas.txt");
using StreamReader reader = new(file);

int goodCount = 0;
int badCount = 0;
int total = 0;
Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
for (string? line = reader.ReadLine(); line != null; line = reader.ReadLine())
{
    total++;
    line = line.Substring(0, line.Length - 1);
    line = line.Substring(1);
    line = line.Replace("\"\"", "\"");
    AntlrInputStream inputStream = new(line.ToString());
    FormulaLexer speakLexer = new(inputStream);
    speakLexer.RemoveErrorListeners();
    CommonTokenStream commonTokenStream = new(speakLexer);
    FormulaParser speakParser = new(commonTokenStream, TextWriter.Null, TextWriter.Null)
    {
        Interpreter = { PredictionMode = Antlr4.Runtime.Atn.PredictionMode.SLL },
    };
    FormulaParser.FormulaContext res = speakParser.formula();
    if (res.exception is not null)
    {
        badCount++;
        Console.WriteLine("ERROR                {0}   {1}", line, res.exception.Message);
    }
    else
    {
        goodCount++;
    }
}

sw.Stop();
Console.WriteLine(
    $"Total: {total}\nGoodCount: {goodCount}\nBadCount: {badCount}\n\nElapsed {sw.ElapsedMilliseconds} ms"
);
Console.WriteLine("\nPress any key...\n");
Console.ReadKey();

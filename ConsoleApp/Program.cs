using System.Diagnostics;
using System.Text;
using ExporterCore;

var commandLineArgs = Environment.GetCommandLineArgs();
#if DEBUG
commandLineArgs = new[] {"", @"C:\Users\Administrator\Downloads\30594988.csv" };
#endif
if (commandLineArgs.Length < 2)
{
    Console.WriteLine(@"input file path invaild");
    return ;
}

string file = commandLineArgs[1];
if (!File.Exists(file))
{
    Console.WriteLine(@"input file path not found");
    return;
}

var parser = new CsvParser.CsvParser(delimeter: ',');
IEnumerable<string[]> lines = parser.ParseFile(file, Encoding.UTF8);
var data = ExportFactory.ExportDataCsv(lines, ExportToFormat.Excel);
if (data == null)
{
    Console.WriteLine(@"input file parse to xlsx error");
    return;
}

var target = file.Replace(".csv", ".xlsx");
File.WriteAllBytes(target, data);
if (!File.Exists(target))
{
    Console.WriteLine(@"input file parse to xlsx fail");
    return;
}

Process.Start("explorer.exe",new FileInfo(target).Directory.FullName);



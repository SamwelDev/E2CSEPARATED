using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.CompilerServices;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string projectRoot = Path.GetFullPath(Path.Combine(baseDirectory, @"..\..\.."));
        string excelDir = Path.Combine(projectRoot, "ExcelFiles");
        string csvDirectory = Path.Combine(projectRoot, "CsvFiles");

        Directory.CreateDirectory(excelDir);
        Directory.CreateDirectory(csvDirectory);

        Console.WriteLine("Excel folder" + excelDir);
        Console.WriteLine("Csv folder" + csvDirectory);

        var excelFiles = Directory.GetFiles(excelDir,"*.xlsx");
        if(excelFiles.Length == 0)
        {
            Console.WriteLine("no excel found in " + excelDir);
            return;
        }

        foreach(var excelPath in excelFiles)
        {
            try
            {
                using var workbook = new XLWorkbook(excelPath);
                var worksheet = workbook.Worksheet(1);
                string fileName = Path.GetFileNameWithoutExtension(excelPath);
                string csvPath = Path.Combine(csvDirectory, fileName + ".csv");
                using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);
                foreach(var row in worksheet.RowsUsed())
                {
                    var cells = row.CellsUsed();
                    string line = string.Join(",", cells.Select(c => EscapeCsv(c.GetValue<string>())));
                    writer.WriteLine(line);
                }
                Console.WriteLine($"Conveerted {fileName}.xlsx -> {fileName}.csv");
            }
            catch(Exception ex)
            {
                Console.WriteLine($"ERROR {excelPath}:{ex.Message}");
            }
            Console.WriteLine("All Excel files processed ...");
        }

        static string EscapeCsv(string value)
        {
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
            {
                value = "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }
    }



}
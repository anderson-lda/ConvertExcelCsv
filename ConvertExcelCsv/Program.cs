using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

namespace ConvertExcelCsv
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string pathInit = @"C:\users\ander\Desktop\curto_prazo\teste_convert\";
            string pathEnd = @"C:\users\ander\Desktop\curto_prazo\teste_convert\"; //neste caso são iguais mas não necessariamente seriam
            string planilhaNome = "Planilha1";
            string csvSeparator = ";";

            try
            {
                var files = Directory.EnumerateFiles(pathInit, "*.*", SearchOption.AllDirectories);
                foreach (string s in files)
                {
                    Console.WriteLine("Arquivo lido:");
                    Console.WriteLine(s);
                    var x = new XLWorkbook(pathInit + Path.GetFileName(s));
                    IXLWorksheet worksheet;
                    x.Worksheets.TryGetWorksheet(planilhaNome, out worksheet);
                    System.IO.File.WriteAllLines(pathEnd + Path.GetFileNameWithoutExtension(s) + ".csv",
                    worksheet.RowsUsed().Select(row =>
                                    string.Join(csvSeparator, row.Cells(1, row.LastCellUsed(false).Address.ColumnNumber)
                                    .Select(cell => cell.GetValue<string>()))
                    ));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("An error ocurred");
                Console.WriteLine(e.Message);
            }
        }
    }
}

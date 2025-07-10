using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetContextColumnFromExcelFile
{
    public class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\PC\OneDrive\Desktop\today task.xlsx";
            string outputPath = @"C:\Users\PC\OneDrive\Desktop\result.txt";

            using (var workbook = new XLWorkbook(filePath))
            using (var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8))
            {
                var worksheet = workbook.Worksheet(1); 
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); 

                foreach (var row in rows)
                {
                    string query = row.Cell("C").GetString();    
                    string trans1 = row.Cell("D").GetString();    
                    string trans2 = row.Cell("F").GetString();    
                    string trans3 = row.Cell("H").GetString();    
                    string trans4 = row.Cell("J").GetString();    
                    string trans5 = row.Cell("L").GetString();    

                    writer.WriteLine(query);
                    writer.WriteLine();
                    writer.WriteLine("BẢN DỊCH 1\n" + trans1 + "\n");
                    writer.WriteLine("BẢN DỊCH 2\n" + trans2 + "\n");
                    writer.WriteLine("BẢN DỊCH 3\n" + trans3 + "\n");
                    writer.WriteLine("BẢN DỊCH 4\n" + trans4 + "\n");
                    writer.WriteLine("BẢN DỊCH 5\n" + trans5 + "\n");
                    writer.WriteLine("--------------------------------------------------\n");
                }
            }

            Console.WriteLine("Xuất thành công ra file: " + outputPath);
            Console.ReadLine();
        }
    }
}

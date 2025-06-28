using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetInformFromExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\PC\OneDrive\Desktop\ABC.xlsx"; 
            string outputPath = @"C:\Users\PC\OneDrive\Desktop\result.txt";  


            var excludedAlways = new HashSet<string> { "语言", "类型", "小结" };

            using (var workbook = new XLWorkbook(filePath))
            using (StreamWriter writer = new StreamWriter(outputPath, false))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().ToList();

                var headerRow = rows[0];
                var dataRows = rows.Skip(1);

                var headers = headerRow.Cells().Select((cell, index) => new
                {
                    Name = cell.GetString().Trim(),
                    Index = index + 1
                }).ToList();

                int count = 1; 

                foreach (var row in dataRows)
                {
                    writer.WriteLine($"========== BÀI SỐ {count} ==========");
                    foreach (var header in headers)
                    {
                        string colName = header.Name;
                        string value = row.Cell(header.Index).GetString().Trim();

                        if (excludedAlways.Contains(colName))
                            continue;

                        if (colName == "错误标签" && string.IsNullOrWhiteSpace(value))
                            continue;

                        writer.WriteLine($"{colName}: {value}");
                    }

                    writer.WriteLine(new string('-', 80));
                    count++;
                }
            }

            Console.WriteLine("✅ Ghi xong file: " + outputPath);
            Console.ReadLine();
        }
    }
}

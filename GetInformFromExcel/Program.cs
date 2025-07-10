using ClosedXML.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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


            Console.Write("Nhập số dòng cần đọc: ");
            if (!int.TryParse(Console.ReadLine(), out int maxRows) || maxRows <= 0)
            {
                Console.WriteLine("⚠️ Số dòng không hợp lệ!");
                return;
            }

            var sb = new StringBuilder();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var sheet = workbook.Worksheet(1); // Lấy sheet đầu tiên

                    int rowCount = Math.Min(maxRows, sheet.LastRowUsed().RowNumber() - 1); // bỏ dòng header

                    for (int row = 2; row < 2 + rowCount; row++) // Bắt đầu từ dòng 2 (bỏ header)
                    {
                        string query = sheet.Cell(row, 1).GetString().Trim();
                        string[] translations = new string[5];
                        for (int i = 0; i < 5; i++)
                        {
                            translations[i] = sheet.Cell(row, i + 2).GetString().Trim();
                        }

                        sb.AppendLine($"🔎 Query: {query}");
                        for (int i = 0; i < 5; i++)
                        {
                            sb.AppendLine($"译文{i + 1}: {translations[i]}");
                        }
                        sb.AppendLine(new string('-', 80));
                    }

                    File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
                    Console.WriteLine($"\nĐã xuất thành công {rowCount} dòng vào: {outputPath}");
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Lỗi: " + ex.Message);
            }

            Console.WriteLine("\nNhấn Enter để thoát...");
            Console.ReadLine();

            Console.ReadLine();
        }
    }
}

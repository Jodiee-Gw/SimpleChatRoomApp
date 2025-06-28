using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ConnectAPIAndAutomaticallyUpdateExcel
{
    public class Program
    {
        static async Task Main()
        {
            string excelPath = @"C:\Users\PC\OneDrive\Desktop\ABC.xlsx";
            string outputPath = @"C:\Users\PC\OneDrive\Desktop\Updated_File.xlsx";
            string apiKey = "a78907bd9e3bf355e59248b5437197f4a63c932072967405dad1b019696b2345"; 
            var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed();

            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

            int rowIndex = 1;

            foreach (var row in rows.Skip(1)) 
            {
                string query = row.Cell(3).GetString().Trim(); 

                for (int i = 0; i < 5; i++)
                {
                    int baseCol = 4 + i * 4;

                    string translation = row.Cell(baseCol).GetString().Trim();
                    string score = row.Cell(baseCol + 1).GetString().Trim();
                    string errorTag = row.Cell(baseCol + 2).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(translation)) continue;

                    string prompt = $@"Dưới đây là một câu gốc và một bản dịch kèm thông tin đánh giá. Hãy viết một đoạn 小结 ngắn (1-2 câu) nhận xét về bản dịch này, có thể nêu điểm tốt hoặc cần cải thiện.

Câu gốc: {query}
Bản dịch: {translation}
Điểm số (评分): {(string.IsNullOrWhiteSpace(score) ? "Không có" : score)}
Lỗi (错误标签): {(string.IsNullOrWhiteSpace(errorTag) ? "Không có" : errorTag)}

=> 小结 tiếng việt:";

                    string summary = await GetSummaryFromAI(prompt, httpClient);

                    row.Cell(baseCol + 3).Value = summary;

                    Console.WriteLine($"Context {rowIndex} - Part{i + 1} completed");
                    await Task.Delay(1000); 
                }

                rowIndex++;
            }

            workbook.SaveAs(outputPath);
            Console.WriteLine("Saved to file: " + outputPath);
            Console.WriteLine("");
        }

        static async Task<string> GetSummaryFromAI(string prompt, HttpClient httpClient)
        {
            var requestBody = new
            {
                model = "mistralai/Mixtral-8x7B-Instruct-v0.1",
                messages = new[]
                {
            new { role = "user", content = prompt }
        }
            };

            var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");
            var response = await httpClient.PostAsync("https://api.together.xyz/v1/chat/completions", content);
            string responseText = await response.Content.ReadAsStringAsync();

            if (!responseText.TrimStart().StartsWith("{"))
            {
                Console.WriteLine("Invalid Resonse from API:");
                Console.WriteLine(responseText);  
                return "[API return an invalid context]";
            }

            using (var doc = JsonDocument.Parse(responseText))
            {
                var root = doc.RootElement;

                if (root.TryGetProperty("error", out var errorObj))
                {
                    string msg = errorObj.GetProperty("message").GetString();
                    Console.WriteLine("❌ API Error: " + msg);
                    return "[API Error: " + msg + "]";
                }

                return root.GetProperty("choices")[0]
                           .GetProperty("message")
                           .GetProperty("content")
                           .GetString()
                           .Trim();
            }
        }

       
    }
}

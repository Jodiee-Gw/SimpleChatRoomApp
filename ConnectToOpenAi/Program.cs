using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ConnectToOpenAi
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var apiKey = "a78907bd9e3bf355e59248b5437197f4a63c932072967405dad1b019696b2345";

            using (HttpClient httpClient = new HttpClient())
            {
                try
                {
                    while (true)
                    {
                        httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                        Console.Write("Enter something: ");
                        var requestBody = new
                        {
                            model = "mistralai/Mixtral-8x7B-Instruct-v0.1",
                            messages = new[]
                             {
                            new { role = "user", content = Console.ReadLine() }
                        }
                        };

                        var jsonRequest = JsonSerializer.Serialize(requestBody);
                        var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

                        var response = await httpClient.PostAsync("https://api.together.xyz/v1/chat/completions", content);
                        var responseString = await response.Content.ReadAsStringAsync();

                        //Console.WriteLine("Response:");
                        //Console.WriteLine(responseString);


                        var json = JsonNode.Parse(responseString);
                        var message = json["choices"][0]["message"]["content"].ToString();

                        Console.WriteLine();
                        Console.Write("AI says:");
                        Console.WriteLine(message);
                        Console.WriteLine();
                    }

                }
                finally
                {
                    httpClient.Dispose();
                }
            }

            Console.ReadLine();

        }
    }
}


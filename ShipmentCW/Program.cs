using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

class Program
{
    static async Task Main()
    {
        var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            Console.WriteLine("Missing ANTHROPIC_API_KEY environment variable.");
            return;
        }

        using var http = new HttpClient();

        // ✅ Correct Anthropic base URL
        var url = "https://api.anthropic.com/v1/messages";

        http.DefaultRequestHeaders.Clear();
        http.DefaultRequestHeaders.Add("x-api-key", apiKey);
        http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01"); // required :contentReference[oaicite:2]{index=2}
        http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var requestBody = new
        {
            // Use a model you have access to in the API console
            model = "claude-sonnet-4-5",
            max_tokens = 200,
            messages = new[]
            {
                new { role = "user", content = "Say hello in one sentence." }
            }
        };

        var json = JsonSerializer.Serialize(requestBody);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");

        using var resp = await http.PostAsync(url, content);
        var body = await resp.Content.ReadAsStringAsync();

        Console.WriteLine($"Request URL: {url}");
        Console.WriteLine($"Status: {(int)resp.StatusCode} {resp.ReasonPhrase}");
        Console.WriteLine("Response body:");
        Console.WriteLine(body);
    }
}

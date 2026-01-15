using ClosedXML.Excel;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

class Program
{
    private const string ApiUrl = "https://api.anthropic.com/v1/messages";

    static async Task Main(string[] args)
    {
        var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            Console.WriteLine("Missing ANTHROPIC_API_KEY. Set it in Project Properties > Debug > Environment variables.");
            return;
        }

        // Root folder defaults to current working directory, but you can set it with /solution
        var root = Directory.GetCurrentDirectory();

        using var http = new HttpClient();
        http.DefaultRequestHeaders.Add("x-api-key", apiKey);
        http.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
        http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var messages = new List<object>();

        Console.WriteLine("Claude Chat Tool (Console)");
        Console.WriteLine("Commands:");
        Console.WriteLine("  /exit                         quit");
        Console.WriteLine("  /reset                        clear conversation");
        Console.WriteLine("  /solution <path>              set repo/solution root folder");
        Console.WriteLine("  /tree [maxFiles]              show a quick file list from root (default 80)");
        Console.WriteLine("  /file <absolutePath>          send a file to Claude");
        Console.WriteLine("  /open <relativePathFromRoot>  send a file relative to root");
        Console.WriteLine("  /code                         paste multi-line code (end with a single line: END)");
        Console.WriteLine();
        Console.WriteLine($"Root folder: {root}");

        while (true)
        {
            Console.Write("\nYou: ");
            var input = Console.ReadLine();
            if (input is null) continue;

            if (input.Equals("/exit", StringComparison.OrdinalIgnoreCase))
                break;

            if (input.Equals("/reset", StringComparison.OrdinalIgnoreCase))
            {
                messages.Clear();
                Console.WriteLine("Conversation cleared.");
                continue;
            }

            if (input.StartsWith("/solution ", StringComparison.OrdinalIgnoreCase))
            {
                var path = input.Substring(10).Trim().Trim('"');
                if (!Directory.Exists(path))
                {
                    Console.WriteLine($"Folder not found: {path}");
                    continue;
                }
                root = path;
                Console.WriteLine($"Root folder set to: {root}");
                continue;
            }

            if (input.StartsWith("/tree", StringComparison.OrdinalIgnoreCase))
            {
                var max = 80;
                var parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2 && int.TryParse(parts[1], out var parsed)) max = parsed;

                var listing = BuildFileListing(root, max);
                Console.WriteLine(listing);

                // Also send listing to Claude (so it understands project structure)
                messages.Add(new { role = "user", content = $"Project file tree (root: {root}):\n\n```\n{listing}\n```" });
                var reply = await CallClaude(http, messages);
                Console.WriteLine("\nClaude: " + reply);
                messages.Add(new { role = "assistant", content = reply });
                continue;
            }

            if (input.StartsWith("/excel ", StringComparison.OrdinalIgnoreCase))
            {
                var path = input.Substring(7).Trim().Trim('"');
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    continue;
                }

                var summary = SummarizeExcelTemplate(path);
                var message = $"Here is a summary of my Excel template:\n\n```\n{summary}\n```\n\n" +
                              "Now generate C# code (a separate console app) that reads this Excel into a strongly typed object and validates required fields. " +
                              "Use ClosedXML. Output NuGet commands and the full code files.";

                messages.Add(new { role = "user", content = message });
                var reply = await CallClaude(http, messages);
                Console.WriteLine("\nClaude: " + reply);
                messages.Add(new { role = "assistant", content = reply });
                continue;
            }

            if (input.StartsWith("/file ", StringComparison.OrdinalIgnoreCase))
            {
                var path = input.Substring(6).Trim().Trim('"');
                input = await ReadFileAsMessage(path);
                if (input.StartsWith("ERROR:"))
                {
                    Console.WriteLine(input);
                    continue;
                }
            }
            else if (input.StartsWith("/open ", StringComparison.OrdinalIgnoreCase))
            {
                var rel = input.Substring(6).Trim().Trim('"');
                var path = Path.Combine(root, rel);
                input = await ReadFileAsMessage(path);
                if (input.StartsWith("ERROR:"))
                {
                    Console.WriteLine(input);
                    continue;
                }
            }
            else if (input.Equals("/code", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Paste your code now. Type END on its own line to finish:");
                var sb = new StringBuilder();
                while (true)
                {
                    var line = Console.ReadLine();
                    if (line is null) continue;
                    if (line.Trim().Equals("END", StringComparison.OrdinalIgnoreCase))
                        break;
                    sb.AppendLine(line);
                }

                input = $"Please review this code and help me improve/fix it:\n\n```csharp\n{sb}\n```";
            }

            // Add user message
            messages.Add(new { role = "user", content = input });

            // Ask Claude
            var responseText = await CallClaude(http, messages);

            Console.WriteLine("\nClaude: " + responseText);

            // Keep assistant reply in history
            messages.Add(new { role = "assistant", content = responseText });
        }
    }

    private static async Task<string> CallClaude(HttpClient http, List<object> messages)
    {
        // IMPORTANT: keep using the model name that worked for you (you got 200 OK).
        // If you change this and get errors, set it back.
        var requestBody = new
        {
            model = "claude-sonnet-4-5",
            max_tokens = 800,
            messages = messages.ToArray()
        };

        var json = JsonSerializer.Serialize(requestBody);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");

        using var resp = await http.PostAsync(ApiUrl, content);
        var body = await resp.Content.ReadAsStringAsync();

        if (!resp.IsSuccessStatusCode)
        {
            return $"[ERROR] Status {(int)resp.StatusCode} {resp.ReasonPhrase}\n{body}";
        }

        var node = JsonNode.Parse(body);
        return node?["content"]?[0]?["text"]?.ToString() ?? "(no text returned)";
    }

    private static async Task<string> ReadFileAsMessage(string path)
    {
        try
        {
            if (!File.Exists(path))
                return $"ERROR: File not found: {path}";

            // Avoid sending huge files accidentally
            var info = new FileInfo(path);
            if (info.Length > 300_000) // ~300 KB
                return $"ERROR: File too large to send ({info.Length} bytes). Consider sending a smaller snippet.";

            // Avoid binaries
            var ext = Path.GetExtension(path).ToLowerInvariant();
            var binaryExts = new HashSet<string> { ".dll", ".exe", ".png", ".jpg", ".jpeg", ".gif", ".zip", ".pdf" };
            if (binaryExts.Contains(ext))
                return $"ERROR: Refusing to send binary file type: {ext}";

            var text = await File.ReadAllTextAsync(path);

            var lang = ext switch
            {
                ".cs" => "csharp",
                ".json" => "json",
                ".xml" => "xml",
                ".csproj" => "xml",
                ".sln" => "",
                _ => ""
            };

            var fence = string.IsNullOrEmpty(lang) ? "```" : $"```{lang}";
            return $"Here is the file `{Path.GetFileName(path)}` from `{path}`:\n\n{fence}\n{text}\n```";
        }
        catch (Exception ex)
        {
            return $"ERROR: Failed reading file: {ex.Message}";
        }
    }

    private static string BuildFileListing(string root, int maxFiles)
    {
        // Keep the tree readable and avoid node_modules/bin/obj
        var ignore = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "bin", "obj", ".git", ".vs", "node_modules"
        };

        var sb = new StringBuilder();
        var count = 0;

        foreach (var file in Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories))
        {
            if (count >= maxFiles) break;

            var rel = Path.GetRelativePath(root, file);

            // Skip ignored folders
            var parts = rel.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (parts.Any(p => ignore.Contains(p))) continue;

            // Prefer typical code/config files
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext is not (".cs" or ".csproj" or ".sln" or ".json" or ".xml" or ".md" or ".yml" or ".yaml"))
                continue;

            sb.AppendLine(rel);
            count++;
        }

        if (count == 0) sb.AppendLine("(No files found or all filtered out)");
        if (count >= maxFiles) sb.AppendLine($"... (showing first {maxFiles})");
        return sb.ToString();
    }

    static string SummarizeExcelTemplate(string path, int maxRows = 60)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();

        var sb = new StringBuilder();
        sb.AppendLine($"Workbook: {Path.GetFileName(path)}");
        sb.AppendLine($"Worksheet: {ws.Name}");
        sb.AppendLine("Showing Column A (field) and Column B (value) rows:");
        sb.AppendLine();

        for (int r = 1; r <= maxRows; r++)
        {
            var a = ws.Cell(r, 1).GetString();
            var b = ws.Cell(r, 2).GetString();

            if (string.IsNullOrWhiteSpace(a)) break;
            sb.AppendLine($"{r:00}. A: {a} | B: {b}");
        }

        return sb.ToString();
    }
}

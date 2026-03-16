using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SlideTemplateFiller.Functions.Helpers
{
    public static class OpenAiHelper
    {
        // imageUrl = previewSas; shapesJson = JSON returned by ShapeExtractor; newTextBlob = content to distribute
        public static async Task<string?> CallOpenAiForMappingsAsync(string imageUrl, string shapesJson, string newTextBlob, string apiKey, ILogger logger)
        {
            if (string.IsNullOrEmpty(apiKey))
            {
                logger.LogWarning("OPENAI_API_KEY empty; skipping real call.");
                return null;
            }

            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

            var payload = new
            {
                model = Environment.GetEnvironmentVariable("OPENAI_MODEL") ?? "gpt-4o",
                input = new object[]
                {
                    new
                    {
                        role = "user",
                        content = new object[]
                        {
                            new { type = "input_text", text = @"You are an assistant that maps a single long content blob into the existing slide's placeholders. You are given:
1) A thumbnail image of the slide (visual context).
2) A JSON list of shapes. Each shape contains: id, font info, existing_text.
3) A ""new_content_blob"" — long text to distribute.

Goal: produce a mapping entry for EVERY shape in the JSON list. Do not skip any shape.
- Respects visual prominence: put headline text into the largest title placeholders.
- Keeps meaning: headline text should contain main claim; body text should keep list elements if present.
- If there is no relevant new content for a shape, keep the existing_text value unchanged.

Output: ONLY JSON with this schema:

{
  ""mappings"": [
    {
      ""shape_id"": ""<id>"",
      ""text"": ""<text to insert (plain text)>"",
      ""style_suggestion"": {""font_size"": <int>, ""bold"": true/false},
      ""reason"": ""<1-2 sentence justification>""
    }
  ],
  ""fallback"": {
    ""action"": ""none"" | ""create_slide"" | ""truncate"" | ""put_in_notes"",
    ""explanation"": ""<why>""
  }
}

Rules:
- You MUST include every shape id from the provided shapes JSON in the mappings array.
- If content must be compressed, prefer semantic compression (shorten sentences, preserve keywords).
- If parts of the blob are lists, keep them as bullets.
- Keep each mapping text concise and human-readable." },
                            new { type = "input_image", image_url = imageUrl },
                            new { type = "input_text", text = shapesJson },
                            new { type = "input_text", text = newTextBlob }
                        }
                    }
                }
            };

            var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
            var openAiEndpoint = Environment.GetEnvironmentVariable("OPENAI_ENDPOINT") ?? "https://api.openai.com/v1/responses";

            var resp = await http.PostAsync(openAiEndpoint, content);
            var raw = await resp.Content.ReadAsStringAsync();

            logger.LogWarning("===== OPENAI RAW RESPONSE START =====");
            logger.LogWarning(raw);
            logger.LogWarning("===== OPENAI RAW RESPONSE END =====");

            if (!resp.IsSuccessStatusCode)
            {
                logger.LogError("OpenAI error: {s}", raw);
                return null;
            }

            try
            {
                using var doc = JsonDocument.Parse(raw);
                // same extraction logic as your console app:
                var sb = new StringBuilder();
                if (doc.RootElement.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array)
                {
                    foreach (var outItem in output.EnumerateArray())
                    {
                        if (outItem.TryGetProperty("content", out var contentArr) && contentArr.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var c in contentArr.EnumerateArray())
                            {
                                if (c.ValueKind == JsonValueKind.Object)
                                {
                                    if (c.TryGetProperty("type", out var typeEl) && typeEl.GetString() == "output_text")
                                    {
                                        if (c.TryGetProperty("text", out var textEl))
                                            sb.AppendLine(textEl.GetString());
                                    }
                                    else if (c.TryGetProperty("text", out var textEl2))
                                    {
                                        sb.AppendLine(textEl2.GetString());
                                    }
                                }
                            }
                        }
                        else if (outItem.TryGetProperty("text", out var t))
                        {
                            sb.AppendLine(t.GetString());
                        }
                    }
                }

                var finalText = sb.ToString().Trim();
                // strip triple backtick fences if present
                if (finalText.StartsWith("```"))
                {
                    int firstNewline = finalText.IndexOf('\n');
                    if (firstNewline >= 0) finalText = finalText.Substring(firstNewline + 1);
                    if (finalText.EndsWith("```")) finalText = finalText.Substring(0, finalText.LastIndexOf("```")).Trim();
                }

                logger.LogInformation("Extracted mapping text length = {L}", finalText?.Length ?? 0);
                return string.IsNullOrWhiteSpace(finalText) ? null : finalText;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Failed to parse OpenAI response");
                return null;
            }
        }
    }
}
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using SpirePresentation = Spire.Presentation;
using System.Drawing;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace SlideTemplateFiller
{
  class Program
  {
    static async Task<int> Main(string[] args)
    {
      if (args.Length < 1)
      {
        Console.WriteLine("Usage: SlideTemplateFiller <input.pptx>");
        return 1;
      }

      string inputPath = args[0];

      if (!File.Exists(inputPath))
      {
        Console.WriteLine($"Input file not found: {inputPath}");
        return 2;
      }

      try
      {
        // --- EXTRACT ---
        var result = SlideProcessor.ProcessPresentation(inputPath);

        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string shapesJson = JsonSerializer.Serialize(result, jsonOptions);
        Console.WriteLine(shapesJson);

        string newTextBlob = "New Text - 1. Artificial Intelligence (AI) & Agentic AI AI is evolving from simple tools to autonomous systems that can plan, decide, and act with minimal human input. Agentic AI: AI systems that complete complex tasks independently (e.g., research, coding, scheduling). AI-native software development: AI helps write, debug, and optimize code. AI for science & medicine: Accelerates drug discovery and climate modeling. Why it matters: Automates knowledge work and boosts productivity. Drives innovation across industries such as healthcare, finance, and manufacturing. Experts note that AI adoption is accelerating across organizations and is becoming a foundational technology rather than optional infrastructure. 2. Quantum Computing Quantum computing uses quantum bits (qubits) to perform calculations that are impossible for classical computers. Key developments: Hybrid quantum–classical systems entering early industry use. Faster processing for AI training, cryptography, and molecular simulations. Governments and companies heavily investing in quantum research. Potential impact: Breakthroughs in drug discovery, material science, and optimization problems. Could eventually crack traditional encryption, prompting post-quantum cybersecurity efforts. Quantum technologies are moving from research labs toward real-world business applications in the mid-2020s. 3. Biotechnology & Genetic Engineering Biotech is rapidly advancing with tools that can modify and engineer living systems. Major areas: CRISPR gene editing and advanced DNA sequencing. Synthetic biology to create new medicines and bio-materials. AI-driven drug discovery and personalized medicine. Impact: Faster development of vaccines and treatments. Potential cures for genetic diseases. Personalized healthcare based on individual DNA. Biotech innovations in gene editing, AI-driven research, and personalized medicine are reshaping healthcare and life sciences.";

        string customInstruction = @"You are an assistant that maps a single long content blob into the existing slide's placeholders. You are given:
1) A thumbnail image of the slide (visual context).
2) A JSON list of shapes. Each shape contains: id, type, bbox, font info, existing_text, capacity_estimate_chars.
3) A ""new_content_blob"" — long text to distribute.

Goal: assign the blob's content into the shapes so the resulting slide:
- Respects visual prominence: put headline text into the largest title placeholders.
- Does not exceed capacity_estimate_chars per shape (treat this as a hard target).
- Keeps meaning: headline text should contain main claim; body text should keep list elements if present.

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

Guidelines:
- Try not to exceed capacity_estimate_chars. If content must be compressed, prefer semantic compression (shorten sentences, preserve keywords), not blind truncation.
- If parts of the blob are lists, keep them as bullets (preserve up to capacity limit).
- If content remains after filling all shapes, set fallback.action to ""create_slide"" and produce mappings for a second slide (repeat shapes with new IDs like sX_cont).
- Keep each mapping text concise and human-readable.";

        string pngPath = Path.ChangeExtension(inputPath, ".png");
        if (!File.Exists(pngPath))
        {
          Console.WriteLine("PNG preview not present — skipping OpenAI call.");
          return 0;
        }

        // --- CALL OPENAI ---
        string? mappingsJson = await OpenAiHelper.SendImageAndJsonToResponsesApiAsync(
          pngPath, shapesJson, customInstruction, newTextBlob);

        if (string.IsNullOrWhiteSpace(mappingsJson))
        {
          Console.WriteLine("No mappings received from OpenAI — aborting write step.");
          return 3;
        }

        // --- WRITE ---
        string outputPath = Path.Combine(
          Path.GetDirectoryName(inputPath) ?? ".",
          "result.pptx");

        SlideWriter.ApplyMappings(inputPath, outputPath, mappingsJson);

        return 0;
      }
      catch (Exception ex)
      {
        Console.WriteLine($"Fatal error: {ex.Message}");
        return 99;
      }
    }
  }

  // ─── EXTRACT ────────────────────────────────────────────────────────────────

  static class SlideProcessor
  {
    public static ExtractionResult ProcessPresentation(string inputPptxPath)
    {
      SaveFirstSlideAsPng(inputPptxPath);

      using var doc = PresentationDocument.Open(inputPptxPath, false);
      var presentationPart = doc.PresentationPart ?? throw new InvalidOperationException("No PresentationPart found.");

      var slidePart = presentationPart.SlideParts.FirstOrDefault()
                      ?? throw new InvalidOperationException("No slides found.");

      var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree
                      ?? throw new InvalidOperationException("No shape tree on slide.");

      var shapes = shapeTree.Descendants<P.Shape>().ToList();
      var metas = ExtractShapeMetasFromSlide(slidePart, shapes);

      return new ExtractionResult { ShapeCount = metas.Count, Shapes = metas };
    }

    public static List<ShapeMeta> ExtractShapeMetasFromSlide(SlidePart slidePart, List<P.Shape> shapes)
    {
      var shapeMetas = new List<ShapeMeta>();

      foreach (var sp in shapes)
      {
        if (sp.TextBody == null) continue;

        string fullText = sp.TextBody.InnerText ?? "";
        if (string.IsNullOrWhiteSpace(fullText)) continue;

        var nv = sp.NonVisualShapeProperties?.NonVisualDrawingProperties;
        string id = nv?.Id?.Value.ToString() ?? Guid.NewGuid().ToString();

        double? chosenFontSizePts = null;
        var firstRunPropsWithSize = sp.TextBody
                                      .Descendants<A.RunProperties>()
                                      .FirstOrDefault(rp => rp.FontSize != null);

        if (firstRunPropsWithSize?.FontSize != null)
        {
          if (double.TryParse(firstRunPropsWithSize.FontSize.Value.ToString(), out double raw))
            chosenFontSizePts = raw / 100.0;
        }

        bool startsWithBold = false;
        var firstRunWithText = sp.TextBody
                                .Descendants<A.Run>()
                                .FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.Text?.Text));

        var firstRunProps = firstRunWithText?.RunProperties;
        if (firstRunProps != null)
        {
          try { startsWithBold = firstRunProps.Bold?.Value ?? false; }
          catch { startsWithBold = false; }
        }

        shapeMetas.Add(new ShapeMeta
        {
          Id = id,
          FullText = fullText,
          ChosenFontSizePts = chosenFontSizePts,
          Bold = startsWithBold
        });
      }

      return shapeMetas;
    }

    public static void SaveFirstSlideAsPng(string pptxPath)
    {
      try
      {
        var ppt = new SpirePresentation.Presentation();
        ppt.LoadFromFile(pptxPath);
        Image img = ppt.Slides[0].SaveAsImage();
        string pngPath = Path.ChangeExtension(pptxPath, ".png");
        img.Save(pngPath);
        Console.WriteLine($"Slide preview saved: {pngPath}");
      }
      catch (Exception ex)
      {
        Console.WriteLine($"PNG export failed: {ex.Message}");
      }
    }
  }

  // ─── OPENAI ──────────────────────────────────────────────────────────────────

  public static class OpenAiHelper
  {
    public static async Task<string?> SendImageAndJsonToResponsesApiAsync(
      string imagePath, string shapesJson, string instruction, string newTextBlob)
    {
      var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
      if (string.IsNullOrEmpty(apiKey))
      {
        Console.WriteLine("OPENAI_API_KEY not set.");
        return null;
      }

      byte[] bytes = await File.ReadAllBytesAsync(imagePath);
      string base64 = Convert.ToBase64String(bytes);
      string dataUri = $"data:image/png;base64,{base64}";

      var payload = new
      {
        model = "gpt-5",
        input = new object[]
        {
          new
          {
            role = "user",
            content = new object[]
            {
              new { type = "input_text", text = instruction },
              new { type = "input_image", image_url = dataUri },
              new { type = "input_text", text = shapesJson },
              new { type = "input_text", text = newTextBlob }
            }
          }
        }
      };

      string json = JsonSerializer.Serialize(payload);

      using var client = new HttpClient();
      client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
      client.DefaultRequestHeaders.UserAgent.ParseAdd("SlideTemplateFiller/1.0");

      try
      {
        using var requestContent = new StringContent(json, Encoding.UTF8, "application/json");
        using var resp = await client.PostAsync("https://api.openai.com/v1/responses", requestContent);
        string respBody = await resp.Content.ReadAsStringAsync();

        if (!resp.IsSuccessStatusCode)
        {
          Console.WriteLine($"OpenAI API error {(int)resp.StatusCode}: {respBody}");
          return null;
        }

        try
        {
          using var doc = JsonDocument.Parse(respBody);
          var root = doc.RootElement;
          var sb = new StringBuilder();

          if (root.TryGetProperty("output", out var output) && output.GetArrayLength() > 0)
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

            string finalText = sb.ToString().Trim();
            if (!string.IsNullOrEmpty(finalText))
            {
              // Strip markdown code fences if present (e.g. ```json ... ```)
              if (finalText.StartsWith("```"))
              {
                int firstNewline = finalText.IndexOf('\n');
                if (firstNewline >= 0)
                  finalText = finalText.Substring(firstNewline + 1);
                if (finalText.EndsWith("```"))
                  finalText = finalText.Substring(0, finalText.LastIndexOf("```")).Trim();
              }

              Console.WriteLine("=== OpenAI Response (extracted) ===");
              Console.WriteLine(finalText);
              Console.WriteLine("=== End ===");
              return finalText;
            }
          }

          Console.WriteLine("Full response JSON (could not find extracted text):");
          Console.WriteLine(respBody);
          return null;
        }
        catch (Exception ex)
        {
          Console.WriteLine("Failed to parse response: " + ex.Message);
          Console.WriteLine("Raw response: " + respBody);
          return null;
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("HTTP error calling OpenAI: " + ex.Message);
        return null;
      }
    }
  }

  // ─── WRITE ───────────────────────────────────────────────────────────────────

  static class SlideWriter
  {
    public static void ApplyMappings(string inputPath, string outputPath, string mappingsJson)
    {
      File.Copy(inputPath, outputPath, true);

      var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
      var doc = JsonSerializer.Deserialize<SlideUpdateResponse>(mappingsJson, options);
      if (doc?.Mappings == null || doc.Mappings.Count == 0)
      {
        Console.WriteLine("No mappings found in JSON.");
        return;
      }

      using (var presentationDocument = PresentationDocument.Open(outputPath, true))
      {
        var presentationPart = presentationDocument.PresentationPart;
        if (presentationPart == null)
        {
          Console.WriteLine("PresentationPart is null.");
          return;
        }

        foreach (var mapping in doc.Mappings)
        {
          if (!int.TryParse(mapping.Shape_Id ?? string.Empty, out int targetId))
          {
            Console.WriteLine($"Skipping mapping; invalid shape_id: '{mapping.Shape_Id}'");
            continue;
          }

          bool applied = false;

          foreach (var slidePart in presentationPart.SlideParts)
          {
            foreach (var shape in slidePart.Slide.Descendants<Shape>())
            {
              var nv = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
              if (nv == null || !nv.Id.HasValue) continue;

              if (nv.Id.Value == targetId)
              {
                SetShapeText(shape, mapping.Text ?? string.Empty, mapping.Style_Suggestion);
                applied = true;
                break;
              }
            }

            if (applied)
            {
              slidePart.Slide.Save();
              break;
            }
          }

          if (!applied)
            Console.WriteLine($"Warning: shape with id={targetId} not found in any slide.");
        }
      }

      Console.WriteLine($"Done. Updated presentation saved as '{outputPath}'.");
    }

    private static void SetShapeText(Shape shape, string newText, StyleSuggestion? style)
    {
      var textBody = shape.TextBody;
      if (textBody == null)
      {
        textBody = new TextBody(
          new A.BodyProperties(),
          new A.ListStyle(),
          new A.Paragraph());
        shape.Append(textBody);
      }

      textBody.RemoveAllChildren<A.Paragraph>();

      string[] lines = newText.Replace("\r\n", "\n").Split('\n');

      foreach (var rawLine in lines)
      {
        string line = rawLine ?? string.Empty;
        var paragraph = new A.Paragraph();

        if (line.TrimStart().StartsWith("•"))
        {
          paragraph.ParagraphProperties = new A.ParagraphProperties() { Level = 0 };
          line = line.Trim();
        }

        var run = new A.Run();
        var runProperties = new A.RunProperties();

        if (style != null)
        {
          if (style.Font_Size.HasValue)
            runProperties.FontSize = style.Font_Size.Value * 100;

          if (style.Bold.HasValue && style.Bold.Value)
            runProperties.Bold = true;
        }

        if (runProperties.HasChildren || runProperties.FontSize != null || runProperties.Bold != null)
          run.Append(runProperties);

        run.Append(new A.Text(line));
        paragraph.Append(run);
        textBody.Append(paragraph);
      }
    }
  }

  // ─── MODELS ──────────────────────────────────────────────────────────────────

  class ExtractionResult
  {
    public int ShapeCount { get; set; }
    public List<ShapeMeta> Shapes { get; set; } = new();
  }

  class ShapeMeta
  {
    public string Id { get; set; } = "";
    public string FullText { get; set; } = "";
    public double? ChosenFontSizePts { get; set; }
    public bool Bold { get; set; } = false;
  }

  class SlideUpdateResponse
  {
    public List<ShapeMapping> Mappings { get; set; } = new();
    public FallbackInfo? Fallback { get; set; }
  }

  class ShapeMapping
  {
    [JsonPropertyName("shape_id")]
    public string? Shape_Id { get; set; }
    public string? Text { get; set; }

    [JsonPropertyName("style_suggestion")]
    public StyleSuggestion? Style_Suggestion { get; set; }
    public string? Reason { get; set; }
  }

  class StyleSuggestion
  {
    [JsonPropertyName("font_size")]
    public int? Font_Size { get; set; }
    public bool? Bold { get; set; }
  }

  class FallbackInfo
  {
    public string? Action { get; set; }
    public string? Explanation { get; set; }
  }
}

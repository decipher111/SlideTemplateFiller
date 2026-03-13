// Program.cs
// Dependencies (NuGet):
// - DocumentFormat.OpenXml
// - Spire.Presentation
// Place this file into a console project and restore packages. Usage:
//   dotnet run -- <input.pptx>

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text;
using System.Net.Http;
using System.Net.Http.Headers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Spire.Presentation;
using System.Drawing;

namespace SlideTemplateFillerMerged
{
  class Program
  {
    static async Task<int> Main(string[] args)
    {
      if (args.Length < 1)
      {
        Console.WriteLine("Usage: SlideTemplateFillerMerged <input.pptx>");
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
        // 1) Extract slide metadata and save PNG preview
        Console.WriteLine("Extracting slide metadata and saving preview PNG...");
        var extraction = SlideProcessor.ProcessPresentation(inputPath);
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string shapesJson = JsonSerializer.Serialize(extraction, jsonOptions);

        // 2) Call OpenAI Responses API with image + shapes JSON
        string pngPath = Path.ChangeExtension(inputPath, ".png");
        string prompt = """
You are an assistant that maps a single long content blob into the existing slide's placeholders. You are given:

1) A thumbnail image of the slide (visual context).
2) A JSON list of shapes. Each shape contains: id, type, bbox, font info, existing_text, capacity_estimate_chars.
3) A "new_content_blob" — long text to distribute.

Goal: assign the blob's content into the shapes so the resulting slide:
- Respects visual prominence: put headline text into the largest title placeholders.
- Does not exceed capacity_estimate_chars per shape (treat this as a hard target).
- Keeps meaning: headline text should contain main claim; body text should keep list elements if present.

Output: ONLY JSON with this schema:

{
  "mappings": [
    {
      "shape_id": "<id>",
      "text": "<text to insert (plain text)>",
      "style_suggestion": {"font_size": <int>, "bold": true/false},
      "reason": "<1-2 sentence justification>"
    }
  ],
  "fallback": {
    "action": "none" | "create_slide" | "truncate" | "put_in_notes",
    "explanation": "<why>"
  }
}

Guidelines:
- Try not to exceed capacity_estimate_chars.
- If content must be compressed, prefer semantic compression (shorten sentences, preserve keywords), not blind truncation.
- If parts of the blob are lists, keep them as bullets (preserve up to capacity limit).
- If content remains after filling all shapes, set fallback.action to "create_slide".
- Keep each mapping text concise and human-readable.
""";
        string newText = @"New Text to be inserted in the slide - 1. Artificial Intelligence (AI) & Agentic AI AI is evolving from simple tools to autonomous systems that can plan, decide, and act with minimal human input. Agentic AI: AI systems that complete complex tasks independently (e.g., research, coding, scheduling). AI-native software development: AI helps write, debug, and optimize code. AI for science & medicine: Accelerates drug discovery and climate modeling. Why it matters: Automates knowledge work and boosts productivity. Drives innovation across industries such as healthcare, finance, and manufacturing. Experts note that AI adoption is accelerating across organizations and is becoming a foundational technology rather than optional infrastructure. 2. Quantum Computing Quantum computing uses quantum bits (qubits) to perform calculations that are impossible for classical computers. Key developments: Hybrid quantum–classical systems entering early industry use. Faster processing for AI training, cryptography, and molecular simulations. Governments and companies heavily investing in quantum research. Potential impact: Breakthroughs in drug discovery, material science, and optimization problems. Could eventually crack traditional encryption, prompting post-quantum cybersecurity efforts. Quantum technologies are moving from research labs toward real-world business applications in the mid-2020s. 3. Biotechnology & Genetic Engineering Biotech is rapidly advancing with tools that can modify and engineer living systems. Major areas: CRISPR gene editing and advanced DNA sequencing. Synthetic biology to create new medicines and bio-materials. AI-driven drug discovery and personalized medicine. Impact: Faster development of vaccines and treatments. Potential cures for genetic diseases. Personalized healthcare based on individual DNA. Biotech innovations in gene editing, AI-driven research, and personalized medicine are reshaping healthcare and life sciences.";
        string assistantReply = await OpenAiHelper.SendImageAndJsonToResponsesApiAsync(pngPath, shapesJson, prompt, newText);
        if (string.IsNullOrWhiteSpace(assistantReply))
        {
          Console.WriteLine("No assistant response obtained; aborting.");
          return 10;
        }

        // 3) Parse assistant reply into mapping model
        Console.WriteLine("Parsing assistant response as mapping JSON...");
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        SlideUpdateResponse mappingDoc = null;

        try
        {
          mappingDoc = JsonSerializer.Deserialize<SlideUpdateResponse>(assistantReply, options);
        }
        catch (Exception)
        {
          // Try to extract a JSON object from assistantReply (first '{' .. last '}')
          int firstBrace = assistantReply.IndexOf('{');
          int lastBrace = assistantReply.LastIndexOf('}');
          if (firstBrace >= 0 && lastBrace > firstBrace)
          {
            string jsonSub = assistantReply.Substring(firstBrace, lastBrace - firstBrace + 1);
            try
            {
              mappingDoc = JsonSerializer.Deserialize<SlideUpdateResponse>(jsonSub, options);
            }
            catch (Exception ex)
            {
              Console.WriteLine("Failed to parse extracted JSON block: " + ex.Message);
            }
          }

          if (mappingDoc == null)
          {
            Console.WriteLine("Unable to parse assistant response into mapping JSON. Printing assistant text:");
            Console.WriteLine(assistantReply);
            return 11;
          }
        }

        if (mappingDoc?.Mappings == null || mappingDoc.Mappings.Count == 0)
        {
          Console.WriteLine("Assistant returned no mappings; aborting.");
          return 12;
        }

        // 4) Apply mappings to a copy of the PPTX (so original left intact)
        string outputPath = MakeOutputPath(inputPath);
        Console.WriteLine($"Copying input to output: {outputPath}");
        File.Copy(inputPath, outputPath, overwrite: true);

        PresentationUpdater.ApplyMappingsToPresentation(outputPath, mappingDoc, out var warnings);

        Console.WriteLine($"Done. Updated presentation saved as '{outputPath}'.");
        if (warnings != null && warnings.Count > 0)
        {
          Console.WriteLine("Warnings:");
          foreach (var w in warnings) Console.WriteLine(" - " + w);
        }

        return 0;
      }
      catch (Exception ex)
      {
        Console.WriteLine($"Fatal error: {ex.Message}");
        Console.WriteLine(ex.ToString());
        return 99;
      }
    }

    static string MakeOutputPath(string input)
    {
      var dir = Path.GetDirectoryName(input) ?? ".";
      var name = Path.GetFileNameWithoutExtension(input);
      var ext = Path.GetExtension(input);
      var outputName = $"{name}-filled{ext}";
      return Path.Combine(dir, outputName);
    }
  }

  // ---------------------------------------------------------
  // SlideProcessor: renders PNG and extracts shape metadata
  // ---------------------------------------------------------
  static class SlideProcessor
  {
    public static ExtractionResult ProcessPresentation(string inputPptxPath)
    {
      // Save preview PNG (non-fatal)
      SaveFirstSlideAsPng(inputPptxPath);

      using var doc = PresentationDocument.Open(inputPptxPath, false); // read-only
      var presentationPart = doc.PresentationPart ?? throw new InvalidOperationException("No PresentationPart found.");

      var slidePart = presentationPart.SlideParts.FirstOrDefault()
                      ?? throw new InvalidOperationException("No slides found.");

      var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree
                      ?? throw new InvalidOperationException("No shape tree on slide.");

      var shapes = shapeTree.Descendants<P.Shape>().ToList();

      var metas = ExtractShapeMetasFromSlide(slidePart, shapes);

      return new ExtractionResult
      {
        ShapeCount = metas.Count,
        Shapes = metas
      };
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

        // Font size from first RunProperties with font size set (OpenXML stores font size as hundredths of a point)
        double? chosenFontSizePts = null;
        var firstRunPropsWithSize = sp.TextBody
                                  .Descendants<A.RunProperties>()
                                  .FirstOrDefault(rp => rp.FontSize != null);

        if (firstRunPropsWithSize?.FontSize != null)
        {
          if (double.TryParse(firstRunPropsWithSize.FontSize.Value.ToString(), out double raw))
          {
            chosenFontSizePts = raw / 100.0;
          }
        }

        // Determine whether the shape's text *starts* with bold:
        bool startsWithBold = false;
        var firstRunWithText = sp.TextBody
                                .Descendants<A.Run>()
                                .FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.Text?.Text));

        var firstRunProps = GetFirstRunProperties(firstRunWithText);
        if (firstRunProps != null)
        {
          try { startsWithBold = firstRunProps.Bold?.Value ?? false; } catch { startsWithBold = false; }
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

    private static A.RunProperties GetFirstRunProperties(A.Run? run) => run?.RunProperties;

    public static void SaveFirstSlideAsPng(string pptxPath)
    {
      try
      {
        var ppt = new Spire.Presentation.Presentation();
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

  // ---------------------------------------------------------
  // OpenAiHelper: send image + JSON + instruction, return textual assistant reply
  // ---------------------------------------------------------
  static class OpenAiHelper
  {
    public static async Task<string> SendImageAndJsonToResponsesApiAsync(string imagePath, string shapesJson, string instruction, string newText)
    {
      var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
      if (string.IsNullOrEmpty(apiKey))
      {
        Console.WriteLine("OPENAI_API_KEY not set.");
        return null;
      }

      if (!File.Exists(imagePath))
      {
        Console.WriteLine("Image file not found: " + imagePath);
        return null;
      }

      byte[] bytes = await File.ReadAllBytesAsync(imagePath);
      string base64 = Convert.ToBase64String(bytes);
      string dataUri = $"data:image/png;base64,{base64}";

      var payload = new
      {
        model = "gpt-4.1-mini",
        input = new object[]
          {
                    new
                    {
                        role = "user",
                        content = new object[]
                        {
                            new { type = "input_text", text = instruction ?? "Analyze the slide using image and JSON." },
                            new { type = "input_image", image_url = dataUri },
                            new { type = "input_text", text = shapesJson },
                            new { type = "input_text", text = newText }
                        }
                    }
          }
      };

      string json = JsonSerializer.Serialize(payload);
      using var client = new HttpClient();
      client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
      client.DefaultRequestHeaders.UserAgent.ParseAdd("SlideTemplateFillerMerged/1.0");

      try
      {
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var resp = await client.PostAsync("https://api.openai.com/v1/responses", content);
        string respBody = await resp.Content.ReadAsStringAsync();

        if (!resp.IsSuccessStatusCode)
        {
          Console.WriteLine($"OpenAI API error {(int)resp.StatusCode}: {respBody}");
          return null;
        }

        // Try to extract textual output robustly:
        try
        {
          using var doc = JsonDocument.Parse(respBody);
          var root = doc.RootElement;

          var sb = new StringBuilder();

          if (root.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array)
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
                      {
                        sb.AppendLine(textEl.GetString());
                      }
                    }
                    else if (c.TryGetProperty("text", out var textEl2))
                    {
                      sb.AppendLine(textEl2.GetString());
                    }
                  }
                  else if (c.ValueKind == JsonValueKind.String)
                  {
                    sb.AppendLine(c.GetString());
                  }
                }
              }
              else if (outItem.TryGetProperty("text", out var t))
              {
                sb.AppendLine(t.GetString());
              }
            }
          }

          string finalText = sb.ToString().Trim();
          if (!string.IsNullOrEmpty(finalText))
          {
            Console.WriteLine("=== Assistant reply (trimmed) ===");
            Console.WriteLine(finalText.Length > 1000 ? finalText.Substring(0, 1000) + "..." : finalText);
            Console.WriteLine("=== End preview ===");
            return finalText;
          }

          // fallback: return whole response JSON as string if no extracted text
          Console.WriteLine("No 'output' textual content found; returning full response JSON as fallback.");
          return respBody;
        }
        catch (Exception ex)
        {
          Console.WriteLine("Failed to parse response: " + ex.Message);
          Console.WriteLine("Raw response:");
          Console.WriteLine(respBody);
          return respBody;
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine("HTTP error calling OpenAI: " + ex.Message);
        return null;
      }
    }
  }

  // ---------------------------------------------------------
  // PresentationUpdater: applies mappings to a PPTX copy
  // ---------------------------------------------------------
  static class PresentationUpdater
  {
    public static void ApplyMappingsToPresentation(string pptxPath, SlideUpdateResponse doc, out List<string> warnings)
    {
      warnings = new List<string>();

      using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxPath, true))
      {
        var presentationPart = presentationDocument.PresentationPart;
        if (presentationPart == null)
        {
          warnings.Add("PresentationPart is null.");
          return;
        }

        foreach (var mapping in doc.Mappings)
        {
          string shapeIdRaw = mapping.Shape_Id ?? mapping.ShapeId ?? string.Empty;
          if (!int.TryParse(shapeIdRaw, out int targetId))
          {
            warnings.Add($"Skipping mapping; invalid shape_id: '{shapeIdRaw}'");
            continue;
          }

          bool applied = false;

          foreach (var slidePart in presentationPart.SlideParts)
          {
            var shapes = slidePart.Slide.Descendants<P.Shape>();
            foreach (var shape in shapes)
            {
              var nv = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
              if (nv == null || !nv.Id.HasValue)
                continue;

              if ((int)nv.Id.Value == targetId)
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
          {
            warnings.Add($"Warning: shape with id={targetId} not found in any slide.");
          }
        }
      }
    }

    /// <summary>
    /// Replace the text of a Shape (text box) with new text and optionally apply style suggestions.
    /// This replaces all paragraphs with a single or multiline paragraph preserving bullet-like line breaks.
    /// </summary>
    private static void SetShapeText(P.Shape shape, string newText, StyleSuggestion style)
    {
      if (shape == null)
        return;

      var textBody = shape.TextBody;
      if (textBody == null)
      {
        textBody = new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph()
        );
        shape.Append(textBody);
      }

      textBody.RemoveAllChildren<A.Paragraph>();

      string[] lines = newText.Replace("\r\n", "\n").Split('\n');

      foreach (var rawLine in lines)
      {
        string line = rawLine ?? string.Empty;
        var paragraph = new A.Paragraph();

        if (line.TrimStart().StartsWith("•") || line.TrimStart().StartsWith("-"))
        {
          paragraph.ParagraphProperties = new A.ParagraphProperties() { Level = 0 };
          // keep the bullet glyph visually by not including the bullet char in the text
          line = line.TrimStart().TrimStart('•').TrimStart('-').Trim();
        }

        var run = new A.Run();
        var runProperties = new A.RunProperties();

        if (style != null)
        {
          if (style.Font_Size.HasValue)
          {
            // OpenXML FontSize is in hundredths of a point
            runProperties.FontSize = style.Font_Size.Value * 100;
          }

          if (style.Bold.HasValue && style.Bold.Value)
          {
            runProperties.Bold = true;
          }
        }

        if (runProperties.HasChildren || runProperties.FontSize != null || runProperties.Bold != null)
        {
          run.Append(runProperties);
        }

        run.Append(new A.Text(line ?? string.Empty));
        paragraph.Append(run);
        textBody.Append(paragraph);
      }
    }
  }

  // ---------------------------------------------------------
  // Models for extraction and mapping JSON
  // ---------------------------------------------------------
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

  // Mapping models (from assistant)
  class SlideUpdateResponse
  {
    public List<ShapeMapping> Mappings { get; set; }
    public FallbackInfo Fallback { get; set; }
  }

  class ShapeMapping
  {
    // Accept both "shape_id" and "shapeId"
    public string Shape_Id { get; set; }
    public string ShapeId { get => Shape_Id; set => Shape_Id = value; }

    public string Text { get; set; }
    public StyleSuggestion Style_Suggestion { get; set; }
    public string Reason { get; set; }
  }

  class StyleSuggestion
  {
    public int? Font_Size { get; set; }
    public bool? Bold { get; set; }

    // lenient property names
    public int? FontSize { get => Font_Size; set => Font_Size = value; }
  }

  class FallbackInfo
  {
    public string Action { get; set; }
    public string Explanation { get; set; }
  }
}
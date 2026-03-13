// using System;
// using System.IO;
// using System.Linq;
// using System.Collections.Generic;
// using System.Threading.Tasks;
// using System.Text.Json;
// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Presentation;
// using DocumentFormat.OpenXml;
// using A = DocumentFormat.OpenXml.Drawing;
// using P = DocumentFormat.OpenXml.Presentation;
// using SpirePresentation = Spire.Presentation;
// using System.Drawing;
// using System.Net.Http;
// using System.Net.Http.Headers;
// using System.Text;
// using System.Text.Json;


// namespace SlideTemplateFiller
// {
//   class Program
//   {
//     static async Task<int> Main(string[] args)
//     {
//       if (args.Length < 1)
//       {
//         Console.WriteLine("Usage: SlideTemplateFiller <input.pptx>");
//         return 1;
//       }

//       string inputPath = args[0];

//       if (!File.Exists(inputPath))
//       {
//         Console.WriteLine($"Input file not found: {inputPath}");
//         return 2;
//       }

//       try
//       {
//         // Process presentation in-place
//         var result = SlideProcessor.ProcessPresentation(inputPath);


//         // Serialize with indentation for console output
//         var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
//         string finalResponse = JsonSerializer.Serialize(result, jsonOptions);

//         Console.WriteLine(finalResponse);

//         string newTextBlob = "New Text - 1. Artificial Intelligence (AI) & Agentic AI AI is evolving from simple tools to autonomous systems that can plan, decide, and act with minimal human input. Agentic AI: AI systems that complete complex tasks independently (e.g., research, coding, scheduling). AI-native software development: AI helps write, debug, and optimize code. AI for science & medicine: Accelerates drug discovery and climate modeling. Why it matters: Automates knowledge work and boosts productivity. Drives innovation across industries such as healthcare, finance, and manufacturing. Experts note that AI adoption is accelerating across organizations and is becoming a foundational technology rather than optional infrastructure. 2. Quantum Computing Quantum computing uses quantum bits (qubits) to perform calculations that are impossible for classical computers. Key developments: Hybrid quantum–classical systems entering early industry use. Faster processing for AI training, cryptography, and molecular simulations. Governments and companies heavily investing in quantum research. Potential impact: Breakthroughs in drug discovery, material science, and optimization problems. Could eventually crack traditional encryption, prompting post-quantum cybersecurity efforts. Quantum technologies are moving from research labs toward real-world business applications in the mid-2020s. 3. Biotechnology & Genetic Engineering Biotech is rapidly advancing with tools that can modify and engineer living systems. Major areas: CRISPR gene editing and advanced DNA sequencing. Synthetic biology to create new medicines and bio-materials. AI-driven drug discovery and personalized medicine. Impact: Faster development of vaccines and treatments. Potential cures for genetic diseases. Personalized healthcare based on individual DNA. Biotech innovations in gene editing, AI-driven research, and personalized medicine are reshaping healthcare and life sciences.";
        
//         string customInstruction = @"You are an assistant that maps a single long content blob into the existing slide's placeholders. You are given: 
// 1) A thumbnail image of the slide (visual context). 
// 2) A JSON list of shapes. Each shape contains: id, type, bbox, font info, existing_text, capacity_estimate_chars. 
// 3) A ""new_content_blob"" — long text to distribute.

// Goal: assign the blob's content into the shapes so the resulting slide:
// - Respects visual prominence: put headline text into the largest title placeholders.
// - Does not exceed capacity_estimate_chars per shape (treat this as a hard target).
// - Keeps meaning: headline text should contain main claim; body text should keep list elements if present.

// Output: ONLY JSON with this schema:

// {
//   ""mappings"": [
//     {
//       ""shape_id"": ""<id>"",
//       ""text"": ""<text to insert (plain text)>"",
//       ""style_suggestion"": {""font_size"": <int>, ""bold"": true/false},
//       ""reason"": ""<1-2 sentence justification>""
//     }
//   ],
//   ""fallback"": {
//     ""action"": ""none"" | ""create_slide"" | ""truncate"" | ""put_in_notes"",
//     ""explanation"": ""<why>""
//   }
// }

// Guidelines:
// - Try not to exceed capacity_estimate_chars. If content must be compressed, prefer semantic compression (shorten sentences, preserve keywords), not blind truncation.
// - If parts of the blob are lists, keep them as bullets (preserve up to capacity limit).
// - If content remains after filling all shapes, set fallback.action to ""create_slide"" and produce mappings for a second slide (repeat shapes with new IDs like sX_cont).
// - Keep each mapping text concise and human-readable.";


//         string pngPath = Path.ChangeExtension(inputPath, ".png");
//         if (File.Exists(pngPath))
//         {

//           await OpenAiHelper.SendImageAndJsonToResponsesApiAsync(pngPath, finalResponse, customInstruction, newTextBlob);
//         }
//         else
//         { 
//           Console.WriteLine("PNG preview not present — skipping OpenAI call.");
//         }

//         // Optional: also save to file (uncomment if you want a sidecar JSON file)
//         // File.WriteAllText(Path.ChangeExtension(inputPath, ".shapes.json"), finalResponse);

//         return 0;
//       }
//       catch (Exception ex)
//       {
//         Console.WriteLine($"Fatal error: {ex.Message}");
//         return 99;
//       }
//     }


//   }

//   static class SlideProcessor
//   {
//     /// <summary>
//     /// Orchestrates opening the presentation, exporting preview PNG, extracting shapes from the first slide,
//     /// and returning an ExtractionResult. Keeps reading-only access to the package.
//     /// </summary>
//     public static ExtractionResult ProcessPresentation(string inputPptxPath)
//     {
//       // Save preview PNG (non-fatal)
//       SaveFirstSlideAsPng(inputPptxPath);

//       using var doc = PresentationDocument.Open(inputPptxPath, false); // read-only
//       var presentationPart = doc.PresentationPart ?? throw new InvalidOperationException("No PresentationPart found.");

//       var slidePart = presentationPart.SlideParts.FirstOrDefault()
//                       ?? throw new InvalidOperationException("No slides found.");

//       var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree
//                       ?? throw new InvalidOperationException("No shape tree on slide.");

//       var shapes = shapeTree.Descendants<P.Shape>().ToList();

//       var metas = ExtractShapeMetasFromSlide(slidePart, shapes);

//       Console.WriteLine(metas);

//       return new ExtractionResult
//       {
//         ShapeCount = metas.Count,
//         Shapes = metas
//       };
//     }

//     /// <summary>
//     /// Extracts ShapeMeta objects from a slide. This is the extracted logic previously inside Main.
//     /// Accepts the slide part and the list of shape elements to inspect.
//     /// </summary>
//     public static List<ShapeMeta> ExtractShapeMetasFromSlide(SlidePart slidePart, List<P.Shape> shapes)
//     {
//       var shapeMetas = new List<ShapeMeta>();

//       foreach (var sp in shapes)
//       {
//         if (sp.TextBody == null) continue;

//         string fullText = sp.TextBody.InnerText ?? "";
//         if (string.IsNullOrWhiteSpace(fullText)) continue;

//         var nv = sp.NonVisualShapeProperties?.NonVisualDrawingProperties;
//         string id = nv?.Id?.Value.ToString() ?? Guid.NewGuid().ToString();

//         // Font size from first RunProperties with font size set (OpenXML stores font size as hundredths of a point)
//         double? chosenFontSizePts = null;
//         var firstRunPropsWithSize = sp.TextBody
//                                       .Descendants<A.RunProperties>()
//                                       .FirstOrDefault(rp => rp.FontSize != null);

//         if (firstRunPropsWithSize?.FontSize != null)
//         {
//           if (double.TryParse(firstRunPropsWithSize.FontSize.Value.ToString(), out double raw))
//           {
//             chosenFontSizePts = raw / 100.0;
//           }
//         }

//         // Determine whether the shape's text *starts* with bold:
//         // The first visible run of text's RunProperties.Bold.Value (if present)
//         bool startsWithBold = false;
//         var firstRunWithText = sp.TextBody
//                                 .Descendants<A.Run>()
//                                 .FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.Text?.Text));

//         var firstRunProps = GetFirstRunProperties(firstRunWithText);

//         if (firstRunProps != null)
//         {
//           try
//           {
//             startsWithBold = firstRunProps.Bold?.Value ?? false;
//           }
//           catch
//           {
//             startsWithBold = false;
//           }
//         }

//         shapeMetas.Add(new ShapeMeta
//         {
//           Id = id,
//           FullText = fullText,
//           ChosenFontSizePts = chosenFontSizePts,
//           Bold = startsWithBold
//         });
//       }

//       return shapeMetas;
//     }

//     /// <summary>
//     /// Returns the RunProperties for a Run, or null if none present.
//     /// Kept as a helper to make intent clearer and to isolate any future parsing logic.
//     /// </summary>
//     private static A.RunProperties GetFirstRunProperties(A.Run? run)
//     {
//       return run?.RunProperties;
//     }

//     /// <summary>
//     /// Save first slide as PNG (keeps your original Spire usage).
//     /// </summary>
//     public static void SaveFirstSlideAsPng(string pptxPath)
//     {
//       try
//       {
//         var ppt = new SpirePresentation.Presentation();
//         ppt.LoadFromFile(pptxPath);

//         // SaveAsImage returns a System.Drawing.Image for the first slide
//         Image img = ppt.Slides[0].SaveAsImage();

//         string pngPath = Path.ChangeExtension(pptxPath, ".png");
//         img.Save(pngPath);

//         Console.WriteLine($"Slide preview saved: {pngPath}");
//       }
//       catch (Exception ex)
//       {
//         Console.WriteLine($"PNG export failed: {ex.Message}");
//       }
//     }

//   }

// public static class OpenAiHelper
// {
//     public static async Task SendImageAndJsonToResponsesApiAsync(string imagePath, string shapesJson, string instruction, string newTextBlob)
//     {
//         var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
//         if (string.IsNullOrEmpty(apiKey))
//         {
//             Console.WriteLine("OPENAI_API_KEY not set.");
//             return;
//         }

//         if (!File.Exists(imagePath))
//         {
//             Console.WriteLine("Image file not found: " + imagePath);
//             return;
//         }

//         byte[] bytes = await File.ReadAllBytesAsync(imagePath);
//         string base64 = Convert.ToBase64String(bytes);
//         string dataUri = $"data:image/png;base64,{base64}";

//         // Build the Responses API payload:
//         var payload = new
//         {
//             model = "gpt-5",
//             // `input` is an array of input items. We'll send a single user input that contains multiple content blocks:
//             input = new object[]
//             {
//                 new
//                 {
//                     role = "user",
//                     content = new object[]
//                     {
//                         new { type = "input_text", text = instruction ?? "Instruction missing" },
//                         new { type = "input_image", image_url = dataUri },
//                         new { type = "input_text", text = shapesJson },
//                         new { type = "input_text", text = newTextBlob ?? "Content blob missing" }
//                     }
//                 }
//             },
//             // optional: max_output_tokens = 800
//         };

//         string json = JsonSerializer.Serialize(payload);

//         using var client = new HttpClient();
//         client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
//         client.DefaultRequestHeaders.UserAgent.ParseAdd("SlideTemplateFiller/1.0");

//         try
//         {
//             using var content = new StringContent(json, Encoding.UTF8, "application/json");

//             using var resp = await client.PostAsync("https://api.openai.com/v1/responses", content);
//             string respBody = await resp.Content.ReadAsStringAsync();

//             if (!resp.IsSuccessStatusCode)
//             {
//                 Console.WriteLine($"OpenAI API error {(int)resp.StatusCode}: {respBody}");
//                 return;
//             }

//             // Try to extract the textual output robustly:
//             try
//             {
//                 using var doc = JsonDocument.Parse(respBody);
//                 var root = doc.RootElement;
//                 Console.WriteLine("Im here");

//                 // Many Responses API outputs place text in `output` → array → each item has `content` array of objects;
//                 // frequently an object of type "output_text" or with property "text" holds the string.
//                 if (root.TryGetProperty("output", out var output) && output.GetArrayLength() > 0)
//                 {
//                     var sb = new StringBuilder();
//                     foreach (var outItem in output.EnumerateArray())
//                     {
//                         if (outItem.TryGetProperty("content", out var contentArr) && contentArr.ValueKind == JsonValueKind.Array)
//                         {
//                             foreach (var c in contentArr.EnumerateArray())
//                             {
//                               Console.WriteLine(c);
//                                 if (c.ValueKind == JsonValueKind.Object)
//                                 {
//                                     if (c.TryGetProperty("type", out var typeEl) && typeEl.GetString() == "output_text")
//                                     {
//                                         if (c.TryGetProperty("text", out var textEl))
//                                         {
//                                             sb.AppendLine(textEl.GetString());
//                                         }
//                                     }
//                                     // fallback: some schema place text under "content"->{"text": "..."} or "text"
//                                     else if (c.TryGetProperty("text", out var textEl2))
//                                     {
//                                         sb.AppendLine(textEl2.GetString());
//                                     }
//                                 }
//                             }
//                         }
//                         // older/alternate field:
//                         else if (outItem.TryGetProperty("text", out var t))
//                         {
//                             sb.AppendLine(t.GetString());
//                         }
//                     }

//                     string finalText = sb.ToString().Trim();
//                     if (!string.IsNullOrEmpty(finalText))
//                     {
//                         Console.WriteLine("=== OpenAI Response (extracted) ===");
//                         Console.WriteLine(finalText);
//                         Console.WriteLine("=== End ===");
//                         return;
//                     }
//                 }

//                 // fallback: print whole response
//                 Console.WriteLine("Full response JSON (could not find extracted text):");
//                 Console.WriteLine(respBody);
//             }
//             catch (Exception ex)
//             {
//                 Console.WriteLine("Failed to parse response: " + ex.Message);
//                 Console.WriteLine("Raw response:");
//                 Console.WriteLine(respBody);
//             }
//         }
//         catch (Exception ex)
//         {
//             Console.WriteLine("HTTP error calling OpenAI: " + ex.Message);
//         }
//     }
// }

//   class ExtractionResult
//   {
//     public int ShapeCount { get; set; }
//     public List<ShapeMeta> Shapes { get; set; } = new();
//   }

//   class ShapeMeta
//   {
//     public string Id { get; set; } = "";
//     public string FullText { get; set; } = "";
//     public double? ChosenFontSizePts { get; set; }

//     // true when the first visible run of text in the shape is bold.
//     public bool Bold { get; set; } = false;
//   }
// }
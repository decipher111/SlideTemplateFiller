// Requires these NuGet packages:
// - DocumentFormat.OpenXml (Open XML SDK)
// This is a complete, copy-pasteable Program.cs. Put "Template Test.pptx" next to the exe
// and run. It will create/overwrite "result.pptx" in the same folder.

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideUpdater
{
    class Program
    {  
        const string BODY_TEXT = @"1. Artificial Intelligence (AI) & Agentic AI AI is evolving from simple tools to autonomous systems that can plan, decide, and act with minimal human input. Agentic AI: AI systems that complete complex tasks independently (e.g., research, coding, scheduling). AI-native software development: AI helps write, debug, and optimize code. AI for science & medicine: Accelerates drug discovery and climate modeling. Why it matters: Automates knowledge work and boosts productivity. Drives innovation across industries such as healthcare, finance, and manufacturing. Experts note that AI adoption is accelerating across organizations and is becoming a foundational technology rather than optional infrastructure. 2. Quantum Computing Quantum computing uses quantum bits (qubits) to perform calculations that are impossible for classical computers. Key developments: Hybrid quantum–classical systems entering early industry use. Faster processing for AI training, cryptography, and molecular simulations. Governments and companies heavily investing in quantum research. Potential impact: Breakthroughs in drug discovery, material science, and optimization problems. Could eventually crack traditional encryption, prompting post-quantum cybersecurity efforts. Quantum technologies are moving from research labs toward real-world business applications in the mid-2020s. 3. Biotechnology & Genetic Engineering Biotech is rapidly advancing with tools that can modify and engineer living systems. Major areas: CRISPR gene editing and advanced DNA sequencing. Synthetic biology to create new medicines and bio-materials. AI-driven drug discovery and personalized medicine. Impact: Faster development of vaccines and treatments. Potential cures for genetic diseases. Personalized healthcare based on individual DNA. Biotech innovations in gene editing, AI-driven research, and personalized medicine are reshaping healthcare and life sciences.";
        private static readonly string HardcodedJson = """
{
  "mappings": [
    {
      "shape_id": "2",
      "text": "Emerging Technology Trends: AI, Quantum, and Biotech",
      "style_suggestion": {"font_size": 28, "bold": true},
      "reason": "Use the main title area to state the slide’s core claim covering all three trends."
    },
    {
      "shape_id": "16",
      "text": "Tech Landscape 2026",
      "style_suggestion": {"font_size": 14, "bold": false},
      "reason": "Secondary banner sets quick context; short to fit the small ribbon."
    },
    {
      "shape_id": "101",
      "text": "Artificial Intelligence & Agentic AI",
      "style_suggestion": {"font_size": 14, "bold": true},
      "reason": "Column 1 header should carry the first key trend."
    },
    {
      "shape_id": "102",
      "text": "• Agentic AI completes complex tasks (research, coding, scheduling)\n• AI‑native dev: write, debug, optimize code\n• AI for science/medicine: drug discovery, climate models",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Concise bullets introduce definitions/examples for the AI trend."
    },
    {
      "shape_id": "8",
      "text": "AI is evolving from simple tools to autonomous systems that can plan, decide, and act with minimal human input. Why it matters: automates knowledge work, boosts productivity, and drives innovation across industries.",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Replaces the ‘what is this pattern’ paragraph with a clear description and impact."
    },
    {
      "shape_id": "10",
      "text": "Key signals\n• Enterprise adoption accelerating; becoming foundational tech\n• Expanding use in healthcare, finance, and manufacturing",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Keeps a short bulleted list of developments to fit the capacity."
    },
    {
      "shape_id": "123",
      "text": "Status: Mainstreaming fast",
      "style_suggestion": {"font_size": 11, "bold": false},
      "reason": "This small footer box succinctly states adoption status."
    },
    {
      "shape_id": "179",
      "text": "Quantum Computing",
      "style_suggestion": {"font_size": 14, "bold": true},
      "reason": "Column 2 header names the second trend."
    },
    {
      "shape_id": "180",
      "text": "• Hybrid quantum–classical systems entering early use\n• Faster processing for AI training, cryptography, molecular sims\n• Heavy government and corporate investment",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Bulleted key developments fit this preface area."
    },
    {
      "shape_id": "182",
      "text": "What it is: Quantum computers use qubits to tackle problems intractable for classical machines.",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Short definition aligns with the ‘what is’ paragraph slot."
    },
    {
      "shape_id": "183",
      "text": "Potential impact\n• Breakthroughs in drug discovery, materials, and optimization\n• Could crack current encryption; drives post‑quantum security\n• Moving from labs toward real‑world apps in the mid‑2020s",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Concise list preserves the main implications."
    },
    {
      "shape_id": "184",
      "text": "Status: Early/partial enterprise use",
      "style_suggestion": {"font_size": 11, "bold": false},
      "reason": "Summarizes maturity in the small footer slot."
    },
    {
      "shape_id": "190",
      "text": "Biotechnology & Genetic Engineering",
      "style_suggestion": {"font_size": 14, "bold": true},
      "reason": "Column 3 header introduces the third trend."
    },
    {
      "shape_id": "191",
      "text": "• CRISPR editing and advanced DNA sequencing\n• Synthetic biology for new medicines and biomaterials\n• AI‑driven discovery and personalized medicine",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Bulleted ‘major areas’ fits the existing list area."
    },
    {
      "shape_id": "193",
      "text": "Biotech is rapidly advancing with tools to modify and engineer living systems, enabling targeted therapies and new biological products.",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Brief explanatory paragraph mirrors the prior ‘what is’ section."
    },
    {
      "shape_id": "194",
      "text": "Impact\n• Faster vaccines and treatments\n• Potential cures for genetic diseases\n• Personalized care based on individual DNA",
      "style_suggestion": {"font_size": 12, "bold": false},
      "reason": "Keeps the key impacts as a short, readable list."
    },
    {
      "shape_id": "195",
      "text": "Status: Rapid progress; strong clinical/regulatory focus",
      "style_suggestion": {"font_size": 11, "bold": false},
      "reason": "Uses the small footer to note maturity and constraints."
    }
  ],
  "fallback": {
    "action": "none",
    "explanation": "All key content was semantically compressed to fit the existing three-column layout without exceeding typical text box capacities."
  }
}
""";
        static void Main(string[] args)
        {
            try
            {
                string inputFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template Test.pptx");
                string outputFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "result.pptx");

                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"ERROR: '{inputFile}' not found. Place 'Template Test.pptx' next to the executable and try again.");
                    return;
                }



                // Make a copy so we modify the copy
                File.Copy(inputFile, outputFile, true);

                // Parse JSON
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var doc = JsonSerializer.Deserialize<SlideUpdateResponse>(HardcodedJson, options);
                if (doc?.Mappings == null || doc.Mappings.Count == 0)
                {
                    Console.WriteLine("No mappings found in JSON.");
                    return;
                }

                // Open result presentation and apply mappings
                using (PresentationDocument presentationDocument = PresentationDocument.Open(outputFile, true))
                {
                    var presentationPart = presentationDocument.PresentationPart;
                    if (presentationPart == null)
                    {
                        Console.WriteLine("PresentationPart is null.");
                        return;
                    }

                    // Build a quick lookup of shape id -> (SlidePart, Shape)
                    // Note: shape_id in mappings corresponds to the shape's NonVisualDrawingProperties.Id
                    foreach (var mapping in doc.Mappings)
                    {
                        if (!int.TryParse(mapping.Shape_Id ?? mapping.ShapeId ?? string.Empty, out int targetId))
                        {
                            // fallback: try to parse numeric from text
                            Console.WriteLine($"Skipping mapping; invalid shape_id: '{mapping.Shape_Id ?? mapping.ShapeId}'");
                            continue;
                        }

                        bool applied = false;

                        foreach (var slidePart in presentationPart.SlideParts)
                        {
                            var shapes = slidePart.Slide.Descendants<Shape>();
                            foreach (var shape in shapes)
                            {
                                var nv = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                                if (nv == null || !nv.Id.HasValue)
                                    continue;

                                if (nv.Id.Value == targetId)
                                {
                                    // Apply replacement
                                    SetShapeText(shape, mapping.Text ?? string.Empty, mapping.Style_Suggestion);
                                    applied = true;
                                    break;
                                }
                            }

                            if (applied)
                            {
                                // Save the slide
                                slidePart.Slide.Save();
                                break;
                            }
                        }

                        if (!applied)
                        {
                            Console.WriteLine($"Warning: shape with id={targetId} not found in any slide.");
                        }
                    }
                }

                Console.WriteLine($"Done. Updated presentation saved as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.ToString());
            }
        }

        /// Replace the text of a Shape (text box) with new text and optionally apply style suggestions.Replaces all paragraphs with a single or multiline paragraph preserving bullet-like line breaks.
        private static void SetShapeText(Shape shape, string newText, StyleSuggestion style)
        {
            if (shape == null)
                return;

            // Ensure TextBody exists
            var textBody = shape.TextBody;
            if (textBody == null)
            {
                // create a basic text body if missing
                textBody = new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph()
                );
                shape.Append(textBody);
            }

            // Clear existing paragraphs
            textBody.RemoveAllChildren<A.Paragraph>();

            // If the newText contains bullet lines (• or leading "-"), keep them on separate paragraphs.
            string[] lines = newText.Replace("\r\n", "\n").Split('\n');

            foreach (var rawLine in lines)
            {
                string line = rawLine ?? string.Empty;

                var paragraph = new A.Paragraph();

                // Basic paragraph properties: if line starts with "•" keep it as bullet.
                if (line.TrimStart().StartsWith("•"))
                {
                    // set a basic bullet by setting a paragraph level property
                    paragraph.ParagraphProperties = new A.ParagraphProperties() { Level = 0 };
                    // remove leading bullet symbol for text content (we keep the bullet glyph visually)
                    line = line.Trim();
                }

                var run = new A.Run();
                var runProperties = new A.RunProperties();

                // Apply style if provided
                if (style != null)
                {
                    if (style.Font_Size.HasValue)
                    {
                        // OpenXML FontSize is specified in 100ths of a point
                        runProperties.FontSize = style.Font_Size.Value * 100;
                    }

                    if (style.Bold.HasValue && style.Bold.Value)
                    {
                        runProperties.Bold = true;
                    }
                }

                // Attach run properties only if any property set; otherwise RunProperties may be empty but it's ok.
                if (runProperties.HasChildren || runProperties.FontSize != null || runProperties.Bold != null)
                {
                    run.Append(runProperties);
                }

                run.Append(new A.Text(line ?? string.Empty));
                paragraph.Append(run);
                textBody.Append(paragraph);
            }

            // If the shape had a placeholder-type drawing properties, preserve them.
        }

        private class SlideUpdateResponse
        {
            public List<ShapeMapping> Mappings { get; set; }
            public FallbackInfo Fallback { get; set; }
        }

        private class ShapeMapping
        {
            // Accept both "shape_id" and "shapeId"
            public string Shape_Id { get; set; }
            // helper property for lenient parsing:
            public string ShapeId { get => Shape_Id; set => Shape_Id = value; }
            public string Text { get; set; }
            public StyleSuggestion Style_Suggestion { get; set; }
            public string Reason { get; set; }
        }

        private class StyleSuggestion
        {
            public int? Font_Size { get; set; }
            public bool? Bold { get; set; }

            // lenient property names
            public int? FontSize { get => Font_Size; set => Font_Size = value; }
        }

        private class FallbackInfo
        {
            public string Action { get; set; }
            public string Explanation { get; set; }
        }
    }
}

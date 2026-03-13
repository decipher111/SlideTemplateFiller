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
    class Program2
    {
        // JSON hardcoded from the previous assistant response (case-insensitive parsing enabled).
        // Use verbatim string with @ and escape double quotes for valid JSON.

                const string BODY_TEXT =
@"1. Artificial Intelligence (AI) & Agentic AI AI is evolving from simple tools to autonomous systems that can plan, decide, and act with minimal human input. Agentic AI: AI systems that complete complex tasks independently (e.g., research, coding, scheduling). AI-native software development: AI helps write, debug, and optimize code. AI for science & medicine: Accelerates drug discovery and climate modeling. Why it matters: Automates knowledge work and boosts productivity. Drives innovation across industries such as healthcare, finance, and manufacturing. Experts note that AI adoption is accelerating across organizations and is becoming a foundational technology rather than optional infrastructure. 2. Quantum Computing Quantum computing uses quantum bits (qubits) to perform calculations that are impossible for classical computers. Key developments: Hybrid quantum–classical systems entering early industry use. Faster processing for AI training, cryptography, and molecular simulations. Governments and companies heavily investing in quantum research. Potential impact: Breakthroughs in drug discovery, material science, and optimization problems. Could eventually crack traditional encryption, prompting post-quantum cybersecurity efforts. Quantum technologies are moving from research labs toward real-world business applications in the mid-2020s. 3. Biotechnology & Genetic Engineering Biotech is rapidly advancing with tools that can modify and engineer living systems. Major areas: CRISPR gene editing and advanced DNA sequencing. Synthetic biology to create new medicines and bio-materials. AI-driven drug discovery and personalized medicine. Impact: Faster development of vaccines and treatments. Potential cures for genetic diseases. Personalized healthcare based on individual DNA. Biotech innovations in gene editing, AI-driven research, and personalized medicine are reshaping healthcare and life sciences.";



        private static readonly string HardcodedJson =
@"{
  ""mappings"": [
    {
      ""shape_id"": ""2"",
      ""text"": ""Emerging Technology Scenario Patterns"",
      ""style_suggestion"": {""font_size"": 14, ""bold"": true},
      ""reason"": ""This is the main slide title and should summarize the overall theme of the new content about major technology domains.""
    },
    {
      ""shape_id"": ""16"",
      ""text"": ""Technology Landscape Overview"",
      ""style_suggestion"": {""font_size"": 18, ""bold"": true, ""color"": ""blue""},
      ""reason"": ""This top banner acts as a category label; it now reflects the broader technology context of the slide.""
    },
    {
      ""shape_id"": ""103"",
      ""text"": ""1"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Number indicator for the first technology pattern column.""
    },
    {
      ""shape_id"": ""101"",
      ""text"": ""Artificial Intelligence & Agentic AI"",
      ""style_suggestion"": {""font_size"": 14, ""bold"": true},
      ""reason"": ""This placeholder is the headline for the first column and should contain the primary topic.""
    },
    {
      ""shape_id"": ""102"",
      ""text"": ""Key developments:\n• Agentic AI systems performing complex tasks autonomously\n• AI-native software development (coding, debugging, optimization)\n• AI accelerating scientific discovery and medical research"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""This box fits the engagement/example section and now lists the main developments in AI from the provided content.""
    },
    {
      ""shape_id"": ""8"",
      ""text"": ""What is this technology trend?\nAI is evolving from simple tools to autonomous systems capable of planning, decision-making, and action with minimal human intervention."",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The section expects a description of the scenario; this summarizes the core definition of modern AI and Agentic AI.""
    },
    {
      ""shape_id"": ""10"",
      ""text"": ""Why it matters:\n• Automates knowledge work and increases productivity\n• Drives innovation across healthcare, finance, and manufacturing\n• Becoming a foundational technology for organizations"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""This area lists key ingredients or implications, which align with the impact and importance of AI adoption.""
    },
    {
      ""shape_id"": ""123"",
      ""text"": ""Technology maturity\n🟢 Rapid enterprise adoption"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The original capability indicator is repurposed to show the maturity/adoption status of AI technologies.""
    },
    {
      ""shape_id"": ""181"",
      ""text"": ""2"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Number indicator for the second technology column.""
    },
    {
      ""shape_id"": ""179"",
      ""text"": ""Quantum Computing"",
      ""style_suggestion"": {""font_size"": 14, ""bold"": true},
      ""reason"": ""This is the headline placeholder for the second column and should display the main topic.""
    },
    {
      ""shape_id"": ""180"",
      ""text"": ""Key developments:\n• Hybrid quantum–classical systems emerging\n• Faster processing for AI training and molecular simulations\n• Strong investment from governments and technology companies"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The engagement list area is repurposed to summarize the main developments described for quantum computing.""
    },
    {
      ""shape_id"": ""182"",
      ""text"": ""What is this technology trend?\nQuantum computing uses qubits to perform calculations far beyond the capability of classical computers."",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""This section is designed for conceptual explanation and now provides the core definition of quantum computing.""
    },
    {
      ""shape_id"": ""183"",
      ""text"": ""Potential impact:\n• Breakthroughs in drug discovery and material science\n• Solving complex optimization problems\n• Driving the need for post-quantum cybersecurity"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The ingredients section now lists the major impacts and applications highlighted in the content.""
    },
    {
      ""shape_id"": ""184"",
      ""text"": ""Technology maturity\n🟡 Emerging / early industry adoption"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The capability indicator is adapted to show maturity stage for quantum technologies.""
    },
    {
      ""shape_id"": ""192"",
      ""text"": ""3"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Number indicator for the third technology column.""
    },
    {
      ""shape_id"": ""190"",
      ""text"": ""Biotechnology & Genetic Engineering"",
      ""style_suggestion"": {""font_size"": 14, ""bold"": true},
      ""reason"": ""This placeholder represents the headline for the third column and should contain the core topic.""
    },
    {
      ""shape_id"": ""191"",
      ""text"": ""Major areas:\n• CRISPR gene editing and advanced DNA sequencing\n• Synthetic biology creating new medicines and bio-materials\n• AI-driven drug discovery and research"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The engagement/example area is repurposed to highlight the primary biotechnology innovation areas.""
    },
    {
      ""shape_id"": ""193"",
      ""text"": ""What is this technology trend?\nBiotechnology is advancing rapidly through tools that can modify and engineer living systems for medical and industrial applications."",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""This text box is intended for conceptual explanation and now summarizes the biotechnology trend.""
    },
    {
      ""shape_id"": ""194"",
      ""text"": ""Impact:\n• Faster vaccine and treatment development\n• Potential cures for genetic diseases\n• Personalized healthcare based on DNA"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The ingredients section now captures the key outcomes and benefits described in the biotech content.""
    },
    {
      ""shape_id"": ""195"",
      ""text"": ""Technology maturity\n🟠 Rapid innovation phase"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""This indicator conveys the evolving but high-innovation stage of biotechnology.""
    },
    {
      ""shape_id"": ""197"",
      ""text"": ""Technology Complexity"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Axis label adjusted to reflect comparison of technological sophistication across the three domains.""
    },
    {
      ""shape_id"": ""199"",
      ""text"": ""Societal & Industry Impact"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""The lower axis describes the value delivered dimension, reframed for technology impact.""
    },
    {
      ""shape_id"": ""203"",
      ""text"": ""Lower"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Start of the complexity scale.""
    },
    {
      ""shape_id"": ""204"",
      ""text"": ""Foundational"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Reworded scale label to fit technology progression context.""
    },
    {
      ""shape_id"": ""205"",
      ""text"": ""Advanced"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Mid-to-high complexity stage descriptor.""
    },
    {
      ""shape_id"": ""206"",
      ""text"": ""Transformative"",
      ""style_suggestion"": {""font_size"": 11},
      ""reason"": ""Highest impact category aligned with breakthrough technologies.""
    }
  ],
  ""fallback"": {
    ""action"": ""none"",
    ""explanation"": ""The provided content fits the existing three-column structure of the slide, with each technology topic mapped to one column.""
  }
}";
        static void Main2(string[] args)
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

        /// <summary>
        /// Replace the text of a Shape (text box) with new text and optionally apply style suggestions.
        /// This replaces all paragraphs with a single or multiline paragraph preserving bullet-like line breaks.
        /// </summary>
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

        // --- JSON mapping models ---
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
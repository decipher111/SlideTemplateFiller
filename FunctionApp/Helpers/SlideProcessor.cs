using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideTemplateFiller.Functions.Helpers
{
    public static class SlideProcessor
    {
        /// <summary>
        /// Applies OpenAI mappings to the template PPTX using OpenXml SDK.
        /// Matches shapes by their NonVisualDrawingProperties.Id — the same ID used by ShapeExtractor.
        /// </summary>
        public static async Task ProcessPresentationAsync(string localTemplatePath, string localOutputPath, string? openAiMappingJson, ILogger logger)
        {
            if (!File.Exists(localTemplatePath))
                throw new FileNotFoundException("Template not found", localTemplatePath);

            // Parse mappings
            var mappings = new Dictionary<string, string>(); // shape id -> new text
            if (!string.IsNullOrWhiteSpace(openAiMappingJson))
            {
                try
                {
                    using var doc = JsonDocument.Parse(openAiMappingJson);
                    if (doc.RootElement.TryGetProperty("mappings", out var arr) && arr.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var entry in arr.EnumerateArray())
                        {
                            var shapeId = entry.TryGetProperty("shape_id", out var s) && s.ValueKind == JsonValueKind.String ? s.GetString() : null;
                            var text = entry.TryGetProperty("text", out var t) && t.ValueKind == JsonValueKind.String ? t.GetString() : null;
                            if (!string.IsNullOrEmpty(shapeId) && text != null)
                                mappings[shapeId] = text;
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "Could not parse mapping JSON; no mappings will be applied.");
                }
            }

            logger.LogInformation("Applying {count} mappings.", mappings.Count);

            // Copy template to output first, then modify in place
            File.Copy(localTemplatePath, localOutputPath, overwrite: true);

            using var pptx = PresentationDocument.Open(localOutputPath, isEditable: true);
            var presPart = pptx.PresentationPart ?? throw new InvalidOperationException("No PresentationPart");

            foreach (var slidePart in presPart.SlideParts)
            {
                var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
                if (shapeTree == null) continue;

                foreach (var sp in shapeTree.Descendants<Shape>())
                {
                    if (sp.TextBody == null) continue;

                    var nv = sp.NonVisualShapeProperties?.NonVisualDrawingProperties;
                    var shapeId = nv?.Id?.Value.ToString();
                    if (shapeId == null || !mappings.TryGetValue(shapeId, out var newText))
                        continue;

                    // Replace all text while preserving the first paragraph's run properties
                    var firstPara = sp.TextBody.Descendants<A.Paragraph>().FirstOrDefault();
                    var firstRunProps = firstPara?.Descendants<A.RunProperties>().FirstOrDefault()?.CloneNode(true);
                    var firstParaProps = firstPara?.ParagraphProperties?.CloneNode(true);

                    // Remove all paragraphs
                    sp.TextBody.RemoveAllChildren<A.Paragraph>();

                    // Add one paragraph with the new text, preserving run properties
                    var para = new A.Paragraph();
                    if (firstParaProps != null)
                        para.AppendChild(firstParaProps.CloneNode(true));

                    var run = new A.Run();
                    if (firstRunProps != null)
                        run.AppendChild(firstRunProps.CloneNode(true));
                    run.AppendChild(new A.Text(newText));
                    para.AppendChild(run);
                    sp.TextBody.AppendChild(para);

                    logger.LogInformation("Applied mapping to shape id={id}.", shapeId);
                }

                slidePart.Slide.Save();
            }

            await Task.CompletedTask;
        }
    }
}

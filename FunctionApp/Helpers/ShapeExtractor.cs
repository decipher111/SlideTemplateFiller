using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideTemplateFiller.Functions.Helpers
{
    public class ShapeInfo
    {
        public string id { get; set; } = "";
        public string type { get; set; } = "";               // guessed: title/body/other
        public BBox bbox { get; set; } = new();
        public double? font_size_pts { get; set; }
        public bool bold { get; set; }
        public string existing_text { get; set; } = "";
        public int capacity_estimate_chars { get; set; }
    }

    public class BBox { public double x { get; set; } public double y { get; set; } public double w { get; set; } public double h { get; set; } }

    public static class ShapeExtractor
    {
        public static string ExtractShapesJson(string localPptxPath)
        {
            var metas = Extract(localPptxPath);
            var opts = new JsonSerializerOptions { WriteIndented = true };
            return JsonSerializer.Serialize(new { shapes = metas }, opts);
        }

        public class ShapeInfoSimple
        {
            public string id { get; set; } = "";
            public double? font_size_pts { get; set; }
            public bool bold { get; set; }
            public string existing_text { get; set; } = "";
        }

        public static List<ShapeInfoSimple> Extract(string localPptxPath)
        {
            var list = new List<ShapeInfoSimple>();

            using var doc = PresentationDocument.Open(localPptxPath, false);
            var presPart = doc.PresentationPart ?? throw new InvalidOperationException("No PresentationPart");
            var firstSlide = presPart.SlideParts.FirstOrDefault() ?? throw new InvalidOperationException("No slide");
            var shapeTree = firstSlide.Slide.CommonSlideData?.ShapeTree;
            if (shapeTree == null) return list;

            var shapes = shapeTree.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();

            foreach (var sp in shapes)
            {
                if (sp.TextBody == null) continue;
                var fullText = sp.TextBody.InnerText ?? "";
                if (string.IsNullOrWhiteSpace(fullText)) continue;

                var nv = sp.NonVisualShapeProperties?.NonVisualDrawingProperties;
                string id = nv?.Id?.Value.ToString() ?? Guid.NewGuid().ToString();

                double? fontPts = null;
                var firstRunPropsWithSize = sp.TextBody
                    .Descendants<A.RunProperties>()
                    .FirstOrDefault(rp => rp.FontSize != null);
                if (firstRunPropsWithSize?.FontSize != null)
                {
                    if (double.TryParse(firstRunPropsWithSize.FontSize.Value.ToString(), out double raw))
                        fontPts = raw / 100.0;
                }

                bool bold = false;
                var firstRun = sp.TextBody.Descendants<A.Run>().FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.Text?.Text));
                var firstRunProps = firstRun?.RunProperties;
                if (firstRunProps != null)
                {
                    try { bold = firstRunProps.Bold?.Value ?? false; }
                    catch { bold = false; }
                }

                list.Add(new ShapeInfoSimple
                {
                    id = id,
                    font_size_pts = fontPts,
                    bold = bold,
                    existing_text = fullText
                });
            }

            return list;
        }
    }

}
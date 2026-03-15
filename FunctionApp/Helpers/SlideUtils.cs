using System;
using System.Drawing.Imaging;
using Spire.Presentation;
using Microsoft.Extensions.Logging;

namespace SlideTemplateFiller.Functions.Helpers
{
    public static class SlideUtils
    {
        public static void SaveFirstSlideAsPng(string pptxPath, string pngPath, ILogger logger)
        {
            try
            {
                var presentation = new Presentation();
                presentation.LoadFromFile(pptxPath);

                var slide = presentation.Slides[0];
                var img = slide.SaveAsImage();

                img.Save(pngPath, ImageFormat.Png);

                presentation.Dispose();
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error creating PNG from PPTX");
                throw;
            }
        }
    }
}
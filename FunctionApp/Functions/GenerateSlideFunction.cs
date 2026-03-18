using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Azure;
using Azure.Storage.Blobs;
using SlideTemplateFiller.Functions.Helpers; // adjust namespace as you choose

namespace SlideTemplateFiller.Functions
{
    public class GenerateSlideFunction
    {
        private readonly ILogger _logger;
        private readonly BlobHelper _blobHelper;
        private readonly string _templateContainer;
        private readonly string _outputContainer;
        private readonly string _previewContainer;
        private readonly string _openAiKey;

        public GenerateSlideFunction(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<GenerateSlideFunction>();

            // read env vars (configured locally in local.settings.json and in Azure portal)
            var connectionString = Environment.GetEnvironmentVariable("BLOB_CONNECTION_STRING");
            _templateContainer = Environment.GetEnvironmentVariable("TEMPLATE_CONTAINER") ?? "templates";
            _outputContainer = Environment.GetEnvironmentVariable("OUTPUT_CONTAINER") ?? "presentations";
            _previewContainer = Environment.GetEnvironmentVariable("PREVIEW_CONTAINER") ?? "previews";
            _openAiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY") ?? "";

            if (string.IsNullOrEmpty(connectionString))
                throw new InvalidOperationException("BLOB_CONNECTION_STRING is not set.");

            _blobHelper = new BlobHelper(connectionString);
        }

        [Function("GenerateSlide")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req,
            FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("GenerateSlide");
            logger.LogInformation("GenerateSlide function triggered.");

            // read request JSON
            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            JsonElement? requestJson = null;
            if (!string.IsNullOrWhiteSpace(requestBody))
            {
                requestJson = JsonSerializer.Deserialize<JsonElement>(requestBody);
            }

            // optional: caller can specify template name (defaults to "template.pptx")
            string templateBlobName = "template.pptx";
            try
            {
                if (requestJson.HasValue &&
                    requestJson.Value.ValueKind == JsonValueKind.Object &&
                    requestJson.Value.TryGetProperty("templateName", out var tnameEl) &&
                    tnameEl.ValueKind == JsonValueKind.String &&
                    !string.IsNullOrWhiteSpace(tnameEl.GetString()))
                {
                    templateBlobName = tnameEl.GetString();
                }
            }
            catch { /* ignore parsing issues */ }

            // create temp working paths
            var tempFolder = Path.Combine(Path.GetTempPath(), "slidetemplate_" + Guid.NewGuid().ToString("n"));
            Directory.CreateDirectory(tempFolder);

            var localTemplatePath = Path.Combine(tempFolder, templateBlobName);
            var pngName = $"preview-{Guid.NewGuid():N}.png";
            var localPngPath = Path.Combine(tempFolder, pngName);
            var outputName = $"result-{Guid.NewGuid():N}.pptx";
            var localOutputPath = Path.Combine(tempFolder, outputName);

            try
            {
                logger.LogInformation("Downloading template from blob: {container}/{blob}", _templateContainer, templateBlobName);
                // Download template from templates container into localTemplatePath
                await _blobHelper.DownloadBlobToFileAsync(_templateContainer, templateBlobName, localTemplatePath);

                // --- Extract shape metadata from the template using your existing logic (or ShapeExtractor) ---
                logger.LogInformation("Extracting shape metadata from template...");
                string shapesJson = ShapeExtractor.ExtractShapesJson(localTemplatePath);
                logger.LogWarning("Shapes JSON (first 1k chars): {s}", shapesJson.Length > 1000 ? shapesJson.Substring(0,1000) + "..." : shapesJson);

                // --- Generate PNG preview of first slide using Spire (or your existing logic) ---
                logger.LogInformation("Creating PNG preview from template (local): {path}", localPngPath);
                // IMPORTANT: Replace SaveFirstSlideAsPng below with your exact current method if needed.
                SlideUtils.SaveFirstSlideAsPng(localTemplatePath, localPngPath, logger);

                // --- Upload PNG preview and create a SAS URL to pass to OpenAI ---
                logger.LogInformation("Uploading preview PNG to blob container: {container}", _previewContainer);
                var previewBlobName = pngName;
                await _blobHelper.UploadFileAsync(_previewContainer, previewBlobName, localPngPath);
                var previewSas = _blobHelper.GenerateReadSasUri(_previewContainer, previewBlobName, minutes: 30);

                logger.LogInformation("Preview uploaded. SAS: {sas}", previewSas);

                // --- Call OpenAI (or your existing OpenAI wrapper) using the previewSas URL ---
                var newTextBlob = requestJson.HasValue && requestJson.Value.TryGetProperty("newTextBlob", out var nt) ? (nt.ValueKind == JsonValueKind.String ? nt.GetString() : null) : null;
                newTextBlob ??= Environment.GetEnvironmentVariable("DEFAULT_NEW_TEXT") ?? "Put your default content here";

                logger.LogInformation("Calling OpenAI with preview URL for mapping...");
                var openAiResult = await OpenAiHelper.CallOpenAiForMappingsAsync(previewSas.ToString(), shapesJson, newTextBlob, _openAiKey, logger);
                logger.LogWarning("OpenAI extracted text (first 2k chars): {t}", openAiResult?.Length > 2000 ? openAiResult.Substring(0,2000) + "..." : openAiResult);
                // (This helper is a small wrapper - replace with your exact OpenAI logic if you have one.)

                // --- Apply mappings to template to generate final PPTX ---
                logger.LogInformation("Applying mappings to generate final PPT: {localOutput}", localOutputPath);
                // TODO: Replace this with your actual slide-mapping implementation.
                // For now we copy the template to output (so functi    on produces an output PPT).
                // SlideProcessor.ApplyMappings(localTemplatePath, localOutputPath, openAiResult);
                await SlideProcessor.ProcessPresentationAsync(localTemplatePath, localOutputPath, openAiResult, logger);
                // File.Copy(localTemplatePath, localOutputPath, overwrite: true);

                // --- Upload the final PPT to output container ---
                logger.LogInformation("Uploading final PPT to blob container: {container}", _outputContainer);
                await _blobHelper.UploadFileAsync(_outputContainer, outputName, localOutputPath);

                // --- Generate SAS for download ---
                var resultSas = _blobHelper.GenerateReadSasUri(_outputContainer, outputName, minutes: 30);

                // Build response payload
                var payload = new
                {
                    status = "success",
                    fileName = outputName,
                    downloadUrl = resultSas.ToString()
                };

                var response = req.CreateResponse(System.Net.HttpStatusCode.OK);
                response.Headers.Add("Content-Type", "application/json");
                await response.WriteStringAsync(JsonSerializer.Serialize(payload));
                return response;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "FULL ERROR");

                var resp = req.CreateResponse(System.Net.HttpStatusCode.InternalServerError);
                await resp.WriteStringAsync(ex.ToString());   // <-- IMPORTANT change
                return resp;
            }
            finally
            {
                try
                {
                    // cleanup temp folder
                    if (Directory.Exists(tempFolder))
                        Directory.Delete(tempFolder, recursive: true);
                }
                catch { /* ignore cleanup errors */ }
            }
        }
    }
}
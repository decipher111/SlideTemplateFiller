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
        // Example: this is a placeholder that demonstrates passing an image URL to LLM.
        // Replace with your exact OpenAI API call / SDK usage and post-processing.
        public static async Task<string> CallOpenAiForMappingsAsync(string imageUrl, string apiKey, ILogger logger)
        {
            if (string.IsNullOrEmpty(apiKey)) 
            {
                logger.LogWarning("OPENAI_API_KEY is empty — skipping real OpenAI call and returning empty mapping.");
                return "{}";
            }

            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

            // Example payload — adapt to model / endpoint you use
            var payload = new
            {
                prompt = $"Analyze this slide and return JSON mapping to fill placeholders. image_url: {imageUrl}",
                // model and other args depend on your OpenAI usage; replace accordingly
            };

            var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
            // TODO: replace URL with your actual OpenAI endpoint (Azure OpenAI or openai.com)
            var openAiEndpoint = Environment.GetEnvironmentVariable("OPENAI_ENDPOINT") ?? "https://api.openai.com/v1/responses";

            var resp = await http.PostAsync(openAiEndpoint, content);
            var text = await resp.Content.ReadAsStringAsync();
            logger.LogInformation("OpenAI response status: {status}", resp.StatusCode);

            // Return the textual JSON / mapping — caller should deserialize as needed
            return text;
        }
    }
}
using System;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OutlookAiAddIn.Services
{
    internal sealed class OpenAIService : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly OpenAIOptions _options;

        public OpenAIService()
        {
            _options = OpenAIOptions.Load();

            if (string.IsNullOrWhiteSpace(_options.ApiKey))
            {
                throw new InvalidOperationException("Chiave API OpenAI non configurata. Imposta OPENAI_API_KEY o aggiorna app.config.");
            }

            _httpClient = new HttpClient
            {
                BaseAddress = new Uri(_options.BaseUrl, UriKind.Absolute),
                Timeout = TimeSpan.FromSeconds(60)
            };

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _options.ApiKey);
        }

        public async Task<string> GenerateAsync(AiInteractionMode mode, string emailContext, string additionalNotes, CancellationToken cancellationToken)
        {
            var prompt = PromptFactory.BuildPrompt(mode, emailContext, additionalNotes);

            var requestObj = new
            {
                model = _options.Model,
                messages = new object[]
                {
                    new { role = "system", content = PromptFactory.SystemPrompt },
                    new { role = "user", content = prompt }
                },
                max_tokens = _options.MaxTokens,
                temperature = _options.Temperature
            };

            var serializer = new JavaScriptSerializer();
            var payload = serializer.Serialize(requestObj);
            using (var content = new StringContent(payload, Encoding.UTF8, "application/json"))
            using (var response = await _httpClient.PostAsync("chat/completions", content, cancellationToken).ConfigureAwait(false))
            {
                var rawBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException($"OpenAI API error: {response.StatusCode} - {rawBody}");
                }

                // Parse minimal JSON to extract choices[0].message.content
                var resp = serializer.DeserializeObject(rawBody) as System.Collections.Generic.Dictionary<string, object>;
                if (resp == null || !resp.ContainsKey("choices"))
                {
                    throw new InvalidOperationException("Risposta OpenAI non valida o vuota.");
                }

                var choices = resp["choices"] as object[];
                if (choices == null || choices.Length == 0)
                {
                    throw new InvalidOperationException("Risposta OpenAI non valida o vuota.");
                }

                string firstContent = null;
                foreach (var c in choices)
                {
                    var choiceDict = c as System.Collections.Generic.Dictionary<string, object>;
                    if (choiceDict != null && choiceDict.ContainsKey("message"))
                    {
                        var msg = choiceDict["message"] as System.Collections.Generic.Dictionary<string, object>;
                        if (msg != null && msg.ContainsKey("content"))
                        {
                            firstContent = msg["content"] as string;
                            if (!string.IsNullOrWhiteSpace(firstContent)) break;
                        }
                    }
                }

                return firstContent == null ? string.Empty : firstContent.Trim();
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }

        private sealed class OpenAIOptions
        {
            public string ApiKey { get; private set; }
            public string Model { get; private set; } = "gpt-4o-mini";
            public string BaseUrl { get; private set; } = "https://api.openai.com/v1/";
            public int MaxTokens { get; private set; } = 600;
            public double Temperature { get; private set; } = 0.7;

            public static OpenAIOptions Load()
            {
                var options = new OpenAIOptions
                {
                    ApiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY") ?? string.Empty
                };

                var appSettings = ConfigurationManager.AppSettings;
                if (appSettings.HasKeys())
                {
                    options.ApiKey = string.IsNullOrWhiteSpace(options.ApiKey)
                        ? appSettings["OpenAI__ApiKey"] ?? string.Empty
                        : options.ApiKey;

                    options.Model = appSettings["OpenAI__Model"] ?? options.Model;
                    options.BaseUrl = appSettings["OpenAI__BaseUrl"] ?? options.BaseUrl;

                    if (int.TryParse(appSettings["OpenAI__MaxTokens"], out var maxTokens))
                    {
                        options.MaxTokens = maxTokens;
                    }

                    if (double.TryParse(appSettings["OpenAI__Temperature"], out var temperature))
                    {
                        options.Temperature = temperature;
                    }
                }

                if (!options.BaseUrl.EndsWith("/", StringComparison.Ordinal))
                {
                    options.BaseUrl += "/";
                }

                return options;
            }
        }

        private static class PromptFactory
        {
            public const string SystemPrompt = "Sei un assistente esperto nella redazione di email professionali. Rispetta il tono richiesto e rispondi sempre in italiano a meno che non sia specificato diversamente.";

            public static string BuildPrompt(AiInteractionMode mode, string emailContext, string additionalNotes)
            {
                var instructions = mode switch
                {
                    AiInteractionMode.SuggestedReply => "Prepara una risposta professionale e concisa alla seguente email.",
                    AiInteractionMode.ImproveDraft => "Riscrivi la bozza seguente rendendola più chiara e professionale mantenendo il significato.",
                    AiInteractionMode.Proofread => "Correggi ortografia, grammatica e punteggiatura del testo seguente senza alterare il significato.",
                    _ => "Aiuta con questa email."
                };

                if (!string.IsNullOrWhiteSpace(additionalNotes))
                {
                    instructions += $" Indicazioni aggiuntive: {additionalNotes}";
                }

                var builder = new StringBuilder();
                builder.AppendLine(instructions);
                builder.AppendLine();
                builder.AppendLine("=== TESTO EMAIL ===");
                builder.AppendLine(emailContext?.Trim() ?? string.Empty);
                builder.AppendLine("===================");

                return builder.ToString();
            }
        }
    }
}

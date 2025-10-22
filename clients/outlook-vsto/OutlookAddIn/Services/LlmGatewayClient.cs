using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using RedInk.OutlookAddIn.Configuration;
using RedInk.OutlookAddIn.Models;

namespace RedInk.OutlookAddIn.Services
{
    public sealed class LlmGatewayClient : IDisposable
    {
        private const string DefaultRoute = "mail-assistant";
        private readonly HttpClient _httpClient;
        private readonly LlmGatewayOptions _options;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private bool _disposed;

        public LlmGatewayClient(LlmGatewayOptions options, HttpClient httpClient = null)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _httpClient = httpClient ?? new HttpClient();

            if (!string.IsNullOrWhiteSpace(options.BaseUrl))
            {
                _httpClient.BaseAddress = new Uri(options.BaseUrl, UriKind.Absolute);
            }

            if (!string.IsNullOrWhiteSpace(options.ApiKey))
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", options.ApiKey);
            }

            if (!string.IsNullOrWhiteSpace(options.Deployment))
            {
                _httpClient.DefaultRequestHeaders.Remove("X-Deployment");
                _httpClient.DefaultRequestHeaders.Add("X-Deployment", options.Deployment);
            }
        }

        public async Task<LlmResponsePayload> ExecuteIntentAsync(LlmIntent intent, MailItemContext context, CancellationToken cancellationToken = default)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            var payload = new LlmRequestPayload
            {
                Intent = intent.ToString(),
                Mail = context
            };

            if (!string.IsNullOrWhiteSpace(_options.DefaultCulture))
            {
                payload.Metadata["culture"] = _options.DefaultCulture;
            }

            if (!string.IsNullOrWhiteSpace(context.ConversationId))
            {
                payload.Metadata["conversationId"] = context.ConversationId;
            }

            payload.Metadata["hasAttachments"] = context.HasAttachments.ToString();

            if (!string.IsNullOrWhiteSpace(context.Language))
            {
                payload.Metadata["language"] = context.Language;
            }

            var json = _serializer.Serialize(payload);
            using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
            using (var response = await _httpClient.PostAsync(DefaultRoute, content, cancellationToken).ConfigureAwait(false))
            {
                response.EnsureSuccessStatusCode();
                var responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(responseBody))
                {
                    return new LlmResponsePayload { Content = "Der KI-Dienst hat keine Antwort geliefert." };
                }

                return _serializer.Deserialize<LlmResponsePayload>(responseBody) ?? new LlmResponsePayload { Content = responseBody };
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _httpClient.Dispose();
            _disposed = true;
        }
    }
}

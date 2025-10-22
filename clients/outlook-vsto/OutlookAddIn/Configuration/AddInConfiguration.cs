using System;
using System.Configuration;
using System.Globalization;

namespace RedInk.OutlookAddIn.Configuration
{
    public static class AddInConfiguration
    {
        private const string BaseUrlKey = "LlmGateway:BaseUrl";
        private const string ApiKeyKey = "LlmGateway:ApiKey";
        private const string DeploymentKey = "LlmGateway:Deployment";
        private const string CultureKey = "LlmGateway:DefaultCulture";

        public static LlmGatewayOptions Load()
        {
            var settings = ConfigurationManager.AppSettings;

            return new LlmGatewayOptions
            {
                BaseUrl = GetSetting(settings, BaseUrlKey, Environment.GetEnvironmentVariable("LLM_GATEWAY_BASEURL") ?? "https://llm-gateway.local/"),
                ApiKey = GetSetting(settings, ApiKeyKey, Environment.GetEnvironmentVariable("LLM_GATEWAY_APIKEY")),
                Deployment = GetSetting(settings, DeploymentKey, "production"),
                DefaultCulture = GetSetting(settings, CultureKey, CultureInfo.CurrentUICulture.Name)
            };
        }

        private static string GetSetting(System.Collections.Specialized.NameValueCollection settings, string key, string fallback)
        {
            var value = settings[key];
            return string.IsNullOrWhiteSpace(value) ? fallback : value;
        }
    }
}

namespace RedInk.OutlookAddIn.Configuration
{
    public sealed class LlmGatewayOptions
    {
        public string BaseUrl { get; set; }

        public string ApiKey { get; set; }

        public string Deployment { get; set; }

        public string DefaultCulture { get; set; }
    }
}

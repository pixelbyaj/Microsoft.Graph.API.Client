namespace MSGraph.Mail.Client.Models
{
    public class Settings
    {
        public string ClientId { get; set; } = string.Empty;
        public string TenantId { get; set; } = string.Empty;
        public string SecretId { get; set; } = string.Empty;
        public string? UserEmail { get; set; }
        public string[]? GraphUserScopes { get; set; }

    }
}

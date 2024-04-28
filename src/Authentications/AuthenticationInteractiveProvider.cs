using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Mail.Client.Models;

namespace Microsoft.Graph.Mail.Client.Authentications
{
    public class AuthenticationInteractiveProvider : AuthenticationProvider
    {
        /// <summary>
        /// Initiate AuthenticationInteractiveProvider class
        /// </summary>
        /// <param name="settings"></param>
        public AuthenticationInteractiveProvider(Settings settings): base (settings)
        {
            ArgumentNullException.ThrowIfNull(settings.ClientId, nameof(settings.ClientId));
            ArgumentNullException.ThrowIfNull(settings.TenantId, nameof(settings.TenantId));
            UserClient = CreateGraphServiceClient();
        }

        /// <summary>
        ///  Create micorosoft graph client
        /// </summary>
        /// <returns></returns>
        protected override GraphServiceClient CreateGraphServiceClient()
        {
            var options = new InteractiveBrowserCredentialOptions
            {
                ClientId = Settings?.ClientId,
                TenantId = Settings?.TenantId,
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                RedirectUri = new Uri("http://localhost"),
            };
            InteractiveBrowserCredential _interactiveBrowserCredential = new(options);
            return new(_interactiveBrowserCredential, Settings?.GraphUserScopes);
        }
    }
}

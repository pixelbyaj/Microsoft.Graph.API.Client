using Azure.Identity;
using Microsoft.Graph;
using MSGraph.Mail.Client.Models;

namespace MSGraph.Mail.Client.Authentications
{
    public class AuthenticationClientSecretProvider : AuthenticationProvider
    {
        /// <summary>
        /// Initiate AuthenticationClientSecretProvider class
        /// </summary>
        /// <param name="settings"></param>
        public AuthenticationClientSecretProvider(Settings settings): base (settings)
        {
            ArgumentNullException.ThrowIfNull(settings.ClientId, nameof(settings.ClientId));
            ArgumentNullException.ThrowIfNull(settings.TenantId, nameof(settings.TenantId));
            ArgumentNullException.ThrowIfNull(settings.SecretId, nameof(settings.SecretId));
            UserClient = CreateGraphServiceClient();
        }

        /// <summary>
        /// Create micorosoft graph client
        /// </summary>
        /// <returns></returns>
        protected override GraphServiceClient CreateGraphServiceClient()
        {
            try
            {
                // using Azure.Identity;
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                };

                string[] scopes = new[] { "https://graph.microsoft.com/.default" };
                ClientSecretCredential clientSecretCredential = new(
                        Settings?.TenantId, Settings?.ClientId, Settings?.SecretId, options);

                return new GraphServiceClient(clientSecretCredential, scopes);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex.InnerException);
            }
        }
    }
}

using Microsoft.Graph;

namespace Microsoft.Graph.Mail.Client.Authentications
{
    public interface IAuthenticationProvider : IDisposable
    {
        internal GraphServiceClient? Client { get; }
        Task<Models.User?> GetUserProfile(string? userMail = null);     
    }
}

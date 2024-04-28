using Microsoft.Graph;

namespace Microsoft.Graph.API.Client.Authentications
{
    public interface IAuthenticationProvider : IDisposable
    {
        internal GraphServiceClient? Client { get; }
        Task<Models.User?> GetUserProfile(string? userMail = null);     
    }
}

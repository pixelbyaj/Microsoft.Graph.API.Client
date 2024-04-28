using Microsoft.Graph;

namespace MSGraph.Mail.Client.Authentications
{
    public interface IAuthenticationProvider : IDisposable
    {
        internal GraphServiceClient? Client { get; }
        Task<Models.User?> GetUserProfile(string? userMail = null);     
    }
}

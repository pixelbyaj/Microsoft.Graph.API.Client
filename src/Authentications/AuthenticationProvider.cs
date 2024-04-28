using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Mail.Client.Models;

namespace Microsoft.Graph.Mail.Client.Authentications
{
    public abstract class AuthenticationProvider : IAuthenticationProvider, IDisposable
    {
        private bool disposed = false;

        protected GraphServiceClient? UserClient { get; set; }

        protected Settings? Settings { get; private set; }

        internal Models.User? User { get; private set; }

        GraphServiceClient? IAuthenticationProvider.Client => UserClient;

        public AuthenticationProvider(Settings? settings)
        {
            ArgumentNullException.ThrowIfNull(settings, nameof(settings));
            Settings = settings;
        }
        /// <summary>
        /// Get User Profile
        /// </summary>
        public async Task<Models.User?> GetUserProfile(string? userMail = null)
        {
            try
            {
                if (User != null)
                {
                    return User;
                }

                if (UserClient == null) return null;

                Microsoft.Graph.Models.User? user = null;
                user = string.IsNullOrEmpty(userMail) ? await UserClient.Me.GetAsync() : await UserClient.Users[userMail].GetAsync();
               
                if (user != null)
                {
                    User = new Models.User
                    {
                        Id = user.Id,
                        DisplayName = user.DisplayName,
                        UserPrincipalName = user.UserPrincipalName,
                        GivenName = user.GivenName,
                        Surname = user.Surname,
                        Mail = user.Mail,
                        JobTitle = user.JobTitle,
                        MobilePhone = user.MobilePhone,
                        BusinessPhones = user.BusinessPhones,
                        OfficeLocation = user.OfficeLocation,
                        PreferredLanguage = user.PreferredLanguage,
                    };
                }
                return User;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex.InnerException);
            }
        }

        protected abstract GraphServiceClient CreateGraphServiceClient();

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    if (UserClient != null)
                    {
                        UserClient.Dispose();
                        UserClient = null;
                        User = null;
                        Settings = null;
                    }
                }


                disposed = true;
            }
        }
    }
}
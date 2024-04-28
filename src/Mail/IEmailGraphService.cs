using MSGraph.Mail.Client.Models;

namespace MSGraph.Mail.Client
{
    public interface IEmailGraphService
    {
        Task<IList<EmailMessage>> GetEmailsAsync(int top = 10, int limit = 10, EmailRequestParameterInformation? requestInformation = null, bool markRead = false);

        public Task<IList<EmailFileAttachment>> GetEmailAttachments(string? messageId);

        Task SendEmail(EmailMessage emailMessage);
    }
}

﻿using Microsoft.Graph.Mail.Client.Models;

namespace Microsoft.Graph.Mail.Client
{
    public interface IEmailGraphService
    {
        Task<IList<EmailMessage>> GetEmailsAsync(int top = 10, int limit = 10, EmailRequestParameterInformation? requestInformation = null, bool markRead = false);

        public Task<IList<EmailFileAttachment>> GetEmailAttachments(string? messageId);

        Task SendEmail(EmailMessage emailMessage);
    }
}

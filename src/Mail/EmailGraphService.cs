using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using MSGraph.Mail.Client.Authentications;
using MSGraph.Mail.Client.Models;

namespace MSGraph.Mail.Client
{
    public class EmailGraphService : IEmailGraphService, IDisposable
    {
        #region private memeber
        private readonly IAuthenticationProvider _authenticationProvider;
        private bool disposed = false;
        #endregion

        #region ctor
        public EmailGraphService(IAuthenticationProvider authenticationProvider)
        {
            ArgumentNullException.ThrowIfNull(nameof(authenticationProvider));
            ArgumentNullException.ThrowIfNull(nameof(authenticationProvider));
            _authenticationProvider = authenticationProvider;
        }
        #endregion

        #region public methods
        /// <summary>
        /// Get Emails. set limit to -1 if would like to fetch all the emails
        /// </summary>
        /// <param name="top"></param>
        /// <param name="limit"></param>
        /// <param name="requestInformation"></param>
        /// <param name="markRead"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public async Task<IList<EmailMessage>> GetEmailsAsync(int top = 10, int limit = 10, EmailRequestParameterInformation? requestInformation = null, bool markRead = false)
        {
            IList<EmailMessage>? messages = new List<EmailMessage>();
            if (_authenticationProvider.Client == null)
            {
                return messages;
            }
            var user = await _authenticationProvider.GetUserProfile();
            try
            {
                MessageCollectionResponse? emailMessages = await _authenticationProvider.Client
                    .Users[user?.Mail]
                    .Messages
                    .GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Top = top;
                        string? emailFilter = GetEmailFilter(requestInformation);
                        if (!string.IsNullOrEmpty(emailFilter))
                        {
                            requestConfiguration.QueryParameters.Filter = emailFilter;
                        }
                        if (requestInformation != null && !string.IsNullOrEmpty(requestInformation?.Search))
                        {
                            requestConfiguration.QueryParameters.Search = requestInformation?.Search;
                        }

                        if (requestInformation != null && requestInformation?.EmailOrderby != null)
                        {
                            IList<string> emailOrderby = new List<string>();
                            foreach (EmailOrderby orderby in requestInformation.EmailOrderby)
                            {
                                if (orderby != null)
                                {
                                    switch (orderby.OrderbyField)
                                    {
                                        case EmailOrderbyField.ToEmail:
                                            string emailToOrderby = "to/emailAddress/name";
                                            if (orderby.OrderbyType != null)
                                            {
                                                string orderbyType = orderby.OrderbyType == EmailOrderbyType.Asc ? "asc" : "desc";
                                                emailToOrderby = $"{emailToOrderby} {orderbyType}";
                                            }
                                            emailOrderby.Add(emailToOrderby);
                                            break;
                                        case EmailOrderbyField.FromEmail:
                                            string emailFromOrderby = "from/emailAddress/name";
                                            if (orderby.OrderbyType != null)
                                            {
                                                string orderbyType = orderby.OrderbyType == EmailOrderbyType.Asc ? "asc" : "desc";
                                                emailFromOrderby = $"{emailFromOrderby} {orderbyType}";
                                            }
                                            emailOrderby.Add(emailFromOrderby);
                                            break;
                                        case EmailOrderbyField.Subject:
                                            string emailSubjectOrderby = "subject";
                                            if (orderby.OrderbyType != null)
                                            {
                                                string orderbyType = orderby.OrderbyType == EmailOrderbyType.Asc ? "asc" : "desc";
                                                emailSubjectOrderby = $"{emailSubjectOrderby} {orderbyType}";
                                            }
                                            emailOrderby.Add(emailSubjectOrderby);
                                            break;
                                        case EmailOrderbyField.receivedDateTime:
                                            string emailReceivedDateTimeOrderby = "receivedDateTime";
                                            if (orderby.OrderbyType != null)
                                            {
                                                string orderbyType = orderby.OrderbyType == EmailOrderbyType.Asc ? "asc" : "desc";
                                                emailReceivedDateTimeOrderby = $"{emailReceivedDateTimeOrderby} {orderbyType}";
                                            }
                                            emailOrderby.Add(emailReceivedDateTimeOrderby);
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }
                            if (emailOrderby != null && emailOrderby.Any())
                            {
                                requestConfiguration.QueryParameters.Orderby = emailOrderby.ToArray();
                            }
                        }
                    });

                if (emailMessages != null && emailMessages.Value != null)
                {
                    messages = await GetEmailMessages(emailMessages, limit, requestInformation?.IncludeAttachments, markRead);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex.InnerException);
            }

            return messages;
        }

        /// <summary>
        /// Get Email Attachments
        /// </summary>
        /// <param name="messageId"></param>
        /// <returns></returns>
        public async Task<IList<EmailFileAttachment>> GetEmailAttachments(string? messageId)
        {
            ArgumentNullException.ThrowIfNull(messageId, nameof(messageId));

            IList<EmailFileAttachment> attachments = new List<EmailFileAttachment>();
            if (_authenticationProvider.Client == null)
            {
                return attachments;
            }
            var user = await _authenticationProvider.GetUserProfile();

            AttachmentCollectionResponse? attachmentCollectionResponse = await _authenticationProvider
                 .Client
                .Users[user?.Mail]
                .Messages[messageId]
                .Attachments
                .GetAsync();

            if (attachmentCollectionResponse != null && attachmentCollectionResponse.Value != null)
            {
                foreach (FileAttachment attachment in attachmentCollectionResponse.Value)
                {
                    var emailAttachment = new EmailFileAttachment
                    {
                        Id = attachment.Id,
                        ContentId = attachment.Id,
                        ContentType = attachment.ContentType,
                        Size = attachment.Size,
                        IsInline = attachment.IsInline,
                        ContentBytes = attachment.ContentBytes
                    };
                    attachments.Add(emailAttachment);
                }
            }

            return attachments;
        }

        public async Task SendEmail(EmailMessage emailMessage)
        {
            ArgumentNullException.ThrowIfNull(emailMessage, nameof(emailMessage));
            ArgumentNullException.ThrowIfNull(emailMessage.Subject, nameof(emailMessage.Subject));
            ArgumentNullException.ThrowIfNull(emailMessage.ToRecipients, nameof(emailMessage.ToRecipients));

            try
            {
                Models.User? user = await _authenticationProvider.GetUserProfile();
                Message message = new()
                {
                    ToRecipients = emailMessage.ToRecipients?.Select(address => new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = address
                        }
                    }
                }).FirstOrDefault(),

                    Subject = emailMessage.Subject,
                    Body = new ItemBody
                    {
                        ContentType = emailMessage.BodyType == EmailBodyType.Html ? BodyType.Html : BodyType.Text,
                        Content = emailMessage.BodyContent
                    },
                };

                if (emailMessage.CcRecipients != null && emailMessage.CcRecipients.Any())
                {
                    message.CcRecipients = emailMessage.CcRecipients?.Select(address => new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = address
                        }
                    }
                }).FirstOrDefault();
                }

                if (emailMessage.BccRecipients != null && emailMessage.BccRecipients.Any())
                {
                    message.BccRecipients = emailMessage.BccRecipients?.Select(address => new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = address
                        }
                    }
                }).FirstOrDefault();
                }

                if (emailMessage.HasAttachments == true && emailMessage.FileAttachments != null)
                {
                    message.Attachments = emailMessage.FileAttachments?.Select(attachment => new List<Attachment>
                {
                    new FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        Name = attachment.Name,
                        ContentType = attachment.ContentType,
                        ContentBytes = attachment.ContentBytes
                    }
                }).FirstOrDefault();
                }

                if (_authenticationProvider.Client != null && user != null)
                {
                    await _authenticationProvider.Client.Users[user.Mail].SendMail.PostAsync(new SendMailPostRequestBody
                    {
                        Message = message,
                        SaveToSentItems = true,
                    });
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex.InnerException);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region private method
        private async Task<IList<EmailMessage>> GetEmailMessages(MessageCollectionResponse? messages, int limit, bool? includeAttachments, bool markRead)
        {
            IList<EmailMessage>? emailMessages = new List<EmailMessage>();
            if (messages == null)
            {
                return emailMessages;
            }
            var user = await _authenticationProvider.GetUserProfile();
            var pageIterator = PageIterator<Message, MessageCollectionResponse>
            .CreatePageIterator(
                _authenticationProvider.Client,
                messages,
                // Callback executed for each item in
                // the collection
                async (message) =>
                {
                    var messageItem = new EmailMessage
                    {
                        Id = message.Id,
                        InternetMessageId = message.InternetMessageId,
                        From = message.From?.EmailAddress?.Address,
                        ToRecipients = message.ToRecipients?.Select(t => t.EmailAddress?.Address).ToList(),
                        CcRecipients = message.CcRecipients?.Select(c => c.EmailAddress?.Address).ToList(),
                        BccRecipients = message.BccRecipients?.Select(b => b.EmailAddress?.Address).ToList(),
                        CreatedDateTime = message.CreatedDateTime,
                        ReceivedDateTime = message.ReceivedDateTime,
                        LastModifiedDateTime = message.LastModifiedDateTime,
                        HasAttachments = message.HasAttachments,
                        FileAttachments = includeAttachments == true && message.HasAttachments == true ? await GetEmailAttachments(message.Id) : null,
                    };
                    if (markRead && _authenticationProvider.Client != null)
                    {
                        message.IsRead = true;
                        await _authenticationProvider
                         .Client
                         .Users[user?.Mail]
                         .Messages[message.Id]
                         .PatchAsync(new Message { IsRead = true });
                    }
                    emailMessages.Add(messageItem);
                    if (limit > 0 && emailMessages.Count == limit)
                    {
                        return false;
                    }

                    return true;
                });

            await pageIterator.IterateAsync();


            return emailMessages;
        }

        private static string? GetEmailFilter(EmailRequestParameterInformation? requestInformation)
        {
            string? emailFilter = string.Empty;

            if (requestInformation == null)
                return emailFilter;

            if (requestInformation != null && requestInformation?.IsRead != null)
            {
                emailFilter = $"isRead eq {requestInformation?.IsRead.Value.ToString().ToLowerInvariant()}";
            }

            if (requestInformation != null && !string.IsNullOrEmpty(requestInformation?.Filter))
            {
                if (string.IsNullOrEmpty(emailFilter))
                {
                    emailFilter = requestInformation?.Filter;
                }
                else
                {
                    emailFilter = $"{emailFilter} and {requestInformation?.Filter}";
                }
            }


            return emailFilter;
        }
        private void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    _authenticationProvider?.Dispose();
                }


                disposed = true;
            }
        }
        #endregion

    }
}
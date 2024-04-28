namespace Microsoft.Graph.API.Client.Models
{
    public class EmailMessage
    {
        public string? Id { get; set; }
        public string? InternetMessageId { get; set; }
        public string? Subject { get; set; }
        public string? BodyPreview { get; set; }
        public string? BodyContent { get; set; }
        public EmailBodyType? BodyType { get; set; }
        public string? Sender { get; set; }
        public string? From { get; set; }
        public IList<string?>? ToRecipients { get; set; }
        public IList<string?>? CcRecipients { get; set; }
        public IList<string?>? BccRecipients { get; set; }
        public bool? HasAttachments { get; set; }
        public IList<EmailFileAttachment>? FileAttachments { get; set; }
        public DateTimeOffset? CreatedDateTime { get; set; }
        public DateTimeOffset? ReceivedDateTime { get; set; }
        public DateTimeOffset? LastModifiedDateTime { get; set; }
    }

    public enum EmailBodyType
    {
        Text,
        Html
    }
}
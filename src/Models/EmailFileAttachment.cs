namespace Microsoft.Graph.API.Client.Models
{
    public class EmailFileAttachment
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? ContentId { get; set; }
        public string? ContentType { get; set; }
        public Int32? Size { get; set; }
        public bool? IsInline { get; set; }
        public byte[]? ContentBytes { get; set; }
        public DateTimeOffset? LastModifiedDateTime { get; set; }
    }
}

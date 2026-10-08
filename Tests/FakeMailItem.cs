using DuplicateEMailRemover.Core;

namespace DuplicateEMailRemover.Core.Tests
{
    // A simple in-memory IMailItem for tests (no Outlook needed).
    internal sealed class FakeMailItem : IMailItem
    {
        public string MessageId { get; set; } = "";
        public string SentOn { get; set; } = "";
        public string ReceivedTime { get; set; } = "";
        public string LastModificationTime { get; set; } = "";
        public string SenderEmailAddress { get; set; } = "";
        public string To { get; set; } = "";
        public string CC { get; set; } = "";
        public string BCC { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Body { get; set; } = "";
        public string HtmlBody { get; set; } = "";
        public List<string> Attachments { get; set; } = [];

        public IEnumerable<string> AttachmentFileNames => Attachments;
    }
}

// Released under the MIT License (see LICENSE.txt).
//
// Which fields are concatenated to decide whether two emails are duplicates.

namespace DuplicateEMailRemover.Core
{
    public sealed class MatchOptions
    {
        public bool MessageId { get; set; }
        public bool SentOn { get; set; }
        public bool ReceivedTime { get; set; }
        public bool LastModTime { get; set; }
        public bool SenderEmail { get; set; }
        public bool To { get; set; }
        public bool CC { get; set; }
        public bool BCC { get; set; }
        public bool Subject { get; set; }
        public bool Body { get; set; }
        public bool HtmlBody { get; set; }
        public bool Attachment { get; set; }

        // Number of fields selected. The app requires at least two.
        public int SelectedCount =>
            (MessageId ? 1 : 0) + (SentOn ? 1 : 0) + (ReceivedTime ? 1 : 0) + (LastModTime ? 1 : 0)
            + (SenderEmail ? 1 : 0) + (To ? 1 : 0) + (CC ? 1 : 0) + (BCC ? 1 : 0)
            + (Subject ? 1 : 0) + (Body ? 1 : 0) + (HtmlBody ? 1 : 0) + (Attachment ? 1 : 0);
    }
}

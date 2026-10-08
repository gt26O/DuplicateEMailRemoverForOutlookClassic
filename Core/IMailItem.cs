// Released under the MIT License (see LICENSE.txt).
//
// A minimal, Outlook-free view of an email, so the duplicate-detection logic
// can be unit tested without Outlook / COM / Windows. The WinForms app provides
// an adapter over Microsoft.Office.Interop.Outlook.MailItem.
//
// String-typed date fields mirror exactly what the original code hashed
// (DateTime.ToString()), so behaviour is preserved.

namespace DuplicateEMailRemover.Core
{
    public interface IMailItem
    {
        // The Internet "Message-ID" header (PR_INTERNET_MESSAGE_ID). Globally
        // unique per email when present; empty for items that have none.
        string MessageId { get; }

        string SentOn { get; }
        string ReceivedTime { get; }
        string LastModificationTime { get; }
        string SenderEmailAddress { get; }
        string To { get; }
        string CC { get; }
        string BCC { get; }
        string Subject { get; }
        string Body { get; }
        string HtmlBody { get; }
        IEnumerable<string> AttachmentFileNames { get; }
    }
}

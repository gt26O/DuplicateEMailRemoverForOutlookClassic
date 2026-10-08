// Released under the MIT License (see LICENSE.txt).
//
// Adapts an Outlook MailItem to the Outlook-free IMailItem interface used by
// DuplicateEMailRemover.Core, so the duplicate-detection logic can run against
// real emails here and against fakes in unit tests.

using System.Runtime.InteropServices;
using DuplicateEMailRemover.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DuplicateEMailRemover
{
    internal sealed class OutlookMailItemAdapter : IMailItem
    {
        private readonly Outlook.MailItem mail;

        public OutlookMailItemAdapter(Outlook.MailItem mailItem)
        {
            mail = mailItem;
        }

        public string SentOn => mail.SentOn.ToString();
        public string ReceivedTime => mail.ReceivedTime.ToString();
        public string LastModificationTime => mail.LastModificationTime.ToString();
        public string SenderEmailAddress => mail.SenderEmailAddress ?? "";
        public string To => mail.To ?? "";
        public string CC => mail.CC ?? "";
        public string BCC => mail.BCC ?? "";
        public string Subject => mail.Subject ?? "";
        public string Body => mail.Body ?? "";
        public string HtmlBody => mail.HTMLBody ?? "";

        public IEnumerable<string> AttachmentFileNames
        {
            get
            {
                List<string> names = [];
                Outlook.Attachments attachments = mail.Attachments;
                foreach (Outlook.Attachment attachment in attachments)
                {
                    names.Add(attachment.FileName);
                    _ = Marshal.ReleaseComObject(attachment);
                }

                _ = Marshal.ReleaseComObject(attachments);
                return names;
            }
        }
    }
}

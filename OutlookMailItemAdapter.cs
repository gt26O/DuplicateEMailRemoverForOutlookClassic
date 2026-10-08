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

        // MAPI property tag for the Internet "Message-ID" header (PR_INTERNET_MESSAGE_ID, PT_UNICODE).
        private const string PrInternetMessageId = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";

        public OutlookMailItemAdapter(Outlook.MailItem mailItem)
        {
            mail = mailItem;
        }

        public string MessageId
        {
            get
            {
                Outlook.PropertyAccessor? accessor = null;
                try
                {
                    accessor = mail.PropertyAccessor;
                    object value = accessor.GetProperty(PrInternetMessageId);
                    return value?.ToString() ?? "";
                }
                catch
                {
                    // Item has no Internet Message-ID (drafts, some non-SMTP items).
                    return "";
                }
                finally
                {
                    if (accessor != null)
                    {
                        _ = Marshal.ReleaseComObject(accessor);
                    }
                }
            }
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

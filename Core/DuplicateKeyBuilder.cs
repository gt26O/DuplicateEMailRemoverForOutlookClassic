// Released under the MIT License (see LICENSE.txt).
//
// Builds the duplicate-matching key for an email and hashes it. Extracted from
// MainForm so it can be unit tested without Outlook. The field order and the
// hex formatting match the original inline implementation exactly.

using System.Security.Cryptography;
using System.Text;

namespace DuplicateEMailRemover.Core
{
    public static class DuplicateKeyBuilder
    {
        // Concatenates the selected fields, in a fixed order, into the raw key.
        // Getters for unselected fields are never touched (important: reading an
        // email Body/HTMLBody over COM is expensive).
        public static string BuildKey(IMailItem item, MatchOptions options)
        {
            StringBuilder sb = new(1024);

            if (options.MessageId) { _ = sb.Append(item.MessageId); }
            if (options.SentOn) { _ = sb.Append(item.SentOn); }
            if (options.ReceivedTime) { _ = sb.Append(item.ReceivedTime); }
            if (options.LastModTime) { _ = sb.Append(item.LastModificationTime); }
            if (options.SenderEmail) { _ = sb.Append(item.SenderEmailAddress); }
            if (options.To) { _ = sb.Append(item.To); }
            if (options.CC) { _ = sb.Append(item.CC); }
            if (options.BCC) { _ = sb.Append(item.BCC); }
            if (options.Subject) { _ = sb.Append(item.Subject); }
            if (options.Body) { _ = sb.Append(item.Body); }
            if (options.HtmlBody) { _ = sb.Append(item.HtmlBody); }

            if (options.Attachment)
            {
                foreach (string name in item.AttachmentFileNames)
                {
                    _ = sb.Append(name);
                }
            }

            return sb.ToString();
        }

        // Upper-case hex MD5 of the key (matches the original behaviour).
        public static string Hash(string key)
        {
            byte[] digest = MD5.HashData(Encoding.UTF8.GetBytes(key));
            StringBuilder sb = new(digest.Length * 2);
            foreach (byte b in digest)
            {
                _ = sb.Append(b.ToString("X2"));
            }

            return sb.ToString();
        }

        // Convenience: build the key and hash it in one call.
        public static string ComputeHash(IMailItem item, MatchOptions options)
        {
            return Hash(BuildKey(item, options));
        }
    }
}

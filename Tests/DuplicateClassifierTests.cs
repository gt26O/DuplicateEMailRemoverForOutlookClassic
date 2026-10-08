using DuplicateEMailRemover.Core;
using Xunit;

namespace DuplicateEMailRemover.Core.Tests
{
    public class DuplicateClassifierTests
    {
        private static MatchOptions SubjectSender() => new() { Subject = true, SenderEmail = true };

        [Fact]
        public void FirstOccurrence_IsOriginal_LaterOnesAreDuplicates()
        {
            DuplicateClassifier c = new(SubjectSender());
            FakeMailItem m1 = new() { Subject = "S", SenderEmailAddress = "a@x.com" };
            FakeMailItem m2 = new() { Subject = "S", SenderEmailAddress = "a@x.com" };
            FakeMailItem m3 = new() { Subject = "S", SenderEmailAddress = "a@x.com" };

            ClassificationResult r1 = c.Classify(m1, @"\\Mailbox\Inbox");
            ClassificationResult r2 = c.Classify(m2, @"\\Mailbox\Archive");
            ClassificationResult r3 = c.Classify(m3, @"\\Mailbox\Deleted Items");

            Assert.Equal(MailDisposition.Original, r1.Disposition);
            Assert.Equal(MailDisposition.Duplicate, r2.Disposition);
            Assert.Equal(MailDisposition.Duplicate, r3.Disposition);
        }

        [Fact]
        public void Duplicate_PointsBackToTheOriginalsFolder()
        {
            DuplicateClassifier c = new(SubjectSender());
            FakeMailItem m1 = new() { Subject = "S", SenderEmailAddress = "a@x.com" };
            FakeMailItem m2 = new() { Subject = "S", SenderEmailAddress = "a@x.com" };

            _ = c.Classify(m1, @"\\Mailbox\Inbox");
            ClassificationResult dup = c.Classify(m2, @"\\Mailbox\Archive");

            Assert.Equal(@"\\Mailbox\Inbox", dup.OriginalFolderPath);
        }

        [Fact]
        public void DistinctEmails_AreAllOriginals()
        {
            DuplicateClassifier c = new(SubjectSender());
            ClassificationResult r1 = c.Classify(new FakeMailItem { Subject = "A", SenderEmailAddress = "a@x.com" }, "f1");
            ClassificationResult r2 = c.Classify(new FakeMailItem { Subject = "B", SenderEmailAddress = "a@x.com" }, "f1");

            Assert.Equal(MailDisposition.Original, r1.Disposition);
            Assert.Equal(MailDisposition.Original, r2.Disposition);
            Assert.Equal(2, c.UniqueCount);
        }

        [Fact]
        public void MatchingFields_ChangeWhatCountsAsDuplicate()
        {
            FakeMailItem a = new() { Subject = "Hi", Body = "one", SenderEmailAddress = "a@x.com" };
            FakeMailItem b = new() { Subject = "Hi", Body = "two", SenderEmailAddress = "a@x.com" };

            // Body OFF -> b is a duplicate of a.
            DuplicateClassifier bodyOff = new(new MatchOptions { Subject = true, SenderEmail = true });
            _ = bodyOff.Classify(a, "f");
            Assert.Equal(MailDisposition.Duplicate, bodyOff.Classify(b, "f").Disposition);

            // Body ON -> b is distinct.
            DuplicateClassifier bodyOn = new(new MatchOptions { Subject = true, SenderEmail = true, Body = true });
            _ = bodyOn.Classify(a, "f");
            Assert.Equal(MailDisposition.Original, bodyOn.Classify(b, "f").Disposition);
        }
    }
}

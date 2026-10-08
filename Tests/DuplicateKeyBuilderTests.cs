using DuplicateEMailRemover.Core;
using Xunit;

namespace DuplicateEMailRemover.Core.Tests
{
    public class DuplicateKeyBuilderTests
    {
        private static MatchOptions SubjectAndSender() => new() { Subject = true, SenderEmail = true };

        [Fact]
        public void IdenticalItems_ProduceSameHash()
        {
            MatchOptions o = SubjectAndSender();
            FakeMailItem a = new() { Subject = "Hello", SenderEmailAddress = "a@x.com", Body = "body-A" };
            FakeMailItem b = new() { Subject = "Hello", SenderEmailAddress = "a@x.com", Body = "body-B" };

            // Body is not selected, so the differing body must not matter.
            Assert.Equal(DuplicateKeyBuilder.ComputeHash(a, o), DuplicateKeyBuilder.ComputeHash(b, o));
        }

        [Fact]
        public void DifferenceInSelectedField_ProducesDifferentHash()
        {
            MatchOptions o = SubjectAndSender();
            FakeMailItem a = new() { Subject = "Hello", SenderEmailAddress = "a@x.com" };
            FakeMailItem b = new() { Subject = "Goodbye", SenderEmailAddress = "a@x.com" };

            Assert.NotEqual(DuplicateKeyBuilder.ComputeHash(a, o), DuplicateKeyBuilder.ComputeHash(b, o));
        }

        [Fact]
        public void BodySelected_MakesDifferentBodiesDistinct()
        {
            MatchOptions o = new() { Subject = true, Body = true };
            FakeMailItem a = new() { Subject = "Hi", Body = "one" };
            FakeMailItem b = new() { Subject = "Hi", Body = "two" };

            Assert.NotEqual(DuplicateKeyBuilder.ComputeHash(a, o), DuplicateKeyBuilder.ComputeHash(b, o));
        }

        [Fact]
        public void AttachmentNames_CountedOnlyWhenSelected()
        {
            FakeMailItem a = new() { Subject = "Hi", Attachments = ["file1.pdf"] };
            FakeMailItem b = new() { Subject = "Hi", Attachments = ["file2.pdf"] };

            MatchOptions without = new() { Subject = true };
            Assert.Equal(DuplicateKeyBuilder.ComputeHash(a, without), DuplicateKeyBuilder.ComputeHash(b, without));

            MatchOptions with = new() { Subject = true, Attachment = true };
            Assert.NotEqual(DuplicateKeyBuilder.ComputeHash(a, with), DuplicateKeyBuilder.ComputeHash(b, with));
        }

        [Fact]
        public void Hash_IsUpperHexMd5OfKey()
        {
            // Known MD5 of the empty string.
            Assert.Equal("D41D8CD98F00B204E9800998ECF8427E", DuplicateKeyBuilder.Hash(""));
        }

        [Fact]
        public void BuildKey_ConcatenatesSelectedFieldsInFixedOrder()
        {
            MatchOptions o = new() { Subject = true, SenderEmail = true };
            FakeMailItem a = new() { Subject = "SUBJ", SenderEmailAddress = "SND" };

            // Order is: ... SenderEmail ... Subject ... (sender before subject)
            Assert.Equal("SNDSUBJ", DuplicateKeyBuilder.BuildKey(a, o));
        }
    }
}

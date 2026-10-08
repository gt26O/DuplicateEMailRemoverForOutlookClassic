using DuplicateEMailRemover.Core;
using Xunit;

namespace DuplicateEMailRemover.Core.Tests
{
    public class DeletedEmailManifestTests
    {
        [Fact]
        public void CsvLine_EscapesQuotesCommasAndNewlines()
        {
            DeletedEmailRecord r = new()
            {
                Id = "abc",
                Subject = "Hello, \"world\"\nsecond line",
                SenderName = "Jane"
            };

            string line = r.ToCsvLine();

            // The comma and quotes must be contained inside a quoted, escaped field.
            Assert.Contains("\"Hello, \"\"world\"\" second line\"", line);
            // A record with an embedded newline must still be a single CSV line.
            Assert.DoesNotContain("\n", line);
            Assert.DoesNotContain("\r", line);
        }

        [Fact]
        public void CsvHeader_ListsExpectedColumns()
        {
            string header = DeletedEmailRecord.CsvHeader();
            Assert.StartsWith("Id,DeletedUtc,FolderPath,Subject", header);
            Assert.Contains("MsgFile", header);
            Assert.EndsWith("RestoredUtc", header);
        }

        [Fact]
        public void SaveThenLoad_RoundTripsRecords()
        {
            string dir = Path.Combine(Path.GetTempPath(), "derm_test_" + Guid.NewGuid().ToString("N"));
            _ = Directory.CreateDirectory(dir);
            try
            {
                DeletedEmailManifest m = new()
                {
                    CreatedUtc = DateTime.UtcNow.ToString("u"),
                    BackupFolder = dir,
                    Records =
                    [
                        new DeletedEmailRecord { Id = "1", Subject = "First", MsgFile = Path.Combine("msg", "1.msg") },
                        new DeletedEmailRecord { Id = "2", Subject = "Second", MsgFile = Path.Combine("msg", "2.msg"), Restored = true }
                    ]
                };
                m.SaveJson();

                DeletedEmailManifest loaded = DeletedEmailManifest.LoadJson(m.JsonPath);

                Assert.Equal(2, loaded.Records.Count);
                Assert.Equal("First", loaded.Records[0].Subject);
                Assert.True(loaded.Records[1].Restored);
                // BackupFolder is resolved from the file's own location.
                Assert.Equal(dir, loaded.BackupFolder);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void ResolveMsgPath_HandlesRelativeAndAbsolute()
        {
            DeletedEmailManifest m = new() { BackupFolder = Path.Combine(Path.GetTempPath(), "backup") };

            DeletedEmailRecord relative = new() { MsgFile = Path.Combine("msg", "x.msg") };
            Assert.Equal(Path.Combine(m.BackupFolder, "msg", "x.msg"), m.ResolveMsgPath(relative));

            string abs = Path.Combine(Path.GetTempPath(), "elsewhere", "y.msg");
            DeletedEmailRecord absolute = new() { MsgFile = abs };
            Assert.Equal(abs, m.ResolveMsgPath(absolute));
        }
    }
}

// Released under the MIT License (see LICENSE.txt).
//
// Data model + reader/writer for the "deleted emails" manifest.
// Every deleted email is given a unique Id and backed up as an individual
// .msg file so it can be listed, audited, and restored later.

using System.Text;
using System.Text.Json;

namespace DuplicateEMailRemover
{
    // One deleted email. Holds everything needed to list it and to restore it
    // from its backup .msg file, without relying on Outlook EntryIDs (which can
    // change on Exchange when items move between stores).
    public sealed class DeletedEmailRecord
    {
        public string Id { get; set; } = "";
        public string DeletedUtc { get; set; } = "";
        public string FolderPath { get; set; } = "";       // original folder (restore target)
        public string StoreId { get; set; } = "";
        public string OriginalEntryId { get; set; } = "";
        public string Subject { get; set; } = "";
        public string SenderName { get; set; } = "";
        public string SenderEmail { get; set; } = "";
        public string SentOn { get; set; } = "";
        public string ReceivedTime { get; set; } = "";
        public long Size { get; set; }
        public string Hash { get; set; } = "";
        public string MsgFile { get; set; } = "";          // path, relative to the backup folder
        public bool Restored { get; set; }
        public string RestoredUtc { get; set; } = "";

        private static string Csv(string? value)
        {
            value ??= "";
            return "\"" + value.Replace("\"", "\"\"").Replace("\r", " ").Replace("\n", " ") + "\"";
        }

        public static string CsvHeader()
        {
            return "Id,DeletedUtc,FolderPath,Subject,SenderName,SenderEmail,SentOn,ReceivedTime,Size,Hash,MsgFile,Restored,RestoredUtc";
        }

        public string ToCsvLine()
        {
            return string.Join(",",
                Csv(Id), Csv(DeletedUtc), Csv(FolderPath), Csv(Subject), Csv(SenderName),
                Csv(SenderEmail), Csv(SentOn), Csv(ReceivedTime), Size.ToString(), Csv(Hash),
                Csv(MsgFile), Restored ? "Yes" : "No", Csv(RestoredUtc));
        }
    }

    // The whole run: metadata + every deleted record. Serialized to JSON for
    // restore, and mirrored to a human-readable CSV.
    public sealed class DeletedEmailManifest
    {
        public string Version { get; set; } = "1";
        public string CreatedUtc { get; set; } = "";
        public string BackupFolder { get; set; } = "";
        public List<DeletedEmailRecord> Records { get; set; } = [];

        public const string JsonFileName = "DeletedEmails.json";
        public const string CsvFileName = "DeletedEmails.csv";
        public const string MsgSubFolder = "msg";

        private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

        public string JsonPath => Path.Combine(BackupFolder, JsonFileName);
        public string CsvPath => Path.Combine(BackupFolder, CsvFileName);
        public string MsgFolder => Path.Combine(BackupFolder, MsgSubFolder);

        public void SaveJson()
        {
            string json = JsonSerializer.Serialize(this, JsonOptions);
            File.WriteAllText(JsonPath, json, Encoding.UTF8);
        }

        public static DeletedEmailManifest LoadJson(string jsonPath)
        {
            string json = File.ReadAllText(jsonPath, Encoding.UTF8);
            DeletedEmailManifest? m = JsonSerializer.Deserialize<DeletedEmailManifest>(json);
            if (m == null)
            {
                throw new InvalidDataException("The manifest file could not be read.");
            }

            // The manifest may have been moved together with its folder; always
            // resolve the backup folder from the file we actually opened.
            m.BackupFolder = Path.GetDirectoryName(Path.GetFullPath(jsonPath)) ?? m.BackupFolder;
            return m;
        }

        // Absolute path of a record's backup .msg, tolerating absolute or relative MsgFile.
        public string ResolveMsgPath(DeletedEmailRecord record)
        {
            return Path.IsPathRooted(record.MsgFile)
                ? record.MsgFile
                : Path.Combine(BackupFolder, record.MsgFile);
        }
    }
}

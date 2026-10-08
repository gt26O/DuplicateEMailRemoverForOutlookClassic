// Released under the MIT License (see LICENSE.txt).
//
// Decides, as emails are streamed in folder-processing order, whether each one
// is the first (kept) copy or a later duplicate. The first email seen for a
// given hash is the "original"; every later email with the same hash is a
// duplicate pointing back at the original's folder.
//
// This is the heart of the tool, now isolated and unit-testable.

namespace DuplicateEMailRemover.Core
{
    public enum MailDisposition
    {
        Original,
        Duplicate
    }

    public sealed class ClassificationResult
    {
        public MailDisposition Disposition { get; init; }
        public string Hash { get; init; } = "";

        // Folder of the first (kept) copy. For an original this is its own folder.
        public string OriginalFolderPath { get; init; } = "";
    }

    public sealed class DuplicateClassifier
    {
        private readonly MatchOptions options;
        private readonly Dictionary<string, string> firstSeenFolderByHash = [];

        public DuplicateClassifier(MatchOptions options)
        {
            this.options = options;
        }

        public int UniqueCount => firstSeenFolderByHash.Count;

        public ClassificationResult Classify(IMailItem item, string folderPath)
        {
            string hash = DuplicateKeyBuilder.ComputeHash(item, options);

            if (firstSeenFolderByHash.TryGetValue(hash, out string? originalFolder))
            {
                return new ClassificationResult
                {
                    Disposition = MailDisposition.Duplicate,
                    Hash = hash,
                    OriginalFolderPath = originalFolder
                };
            }

            firstSeenFolderByHash[hash] = folderPath;
            return new ClassificationResult
            {
                Disposition = MailDisposition.Original,
                Hash = hash,
                OriginalFolderPath = folderPath
            };
        }
    }
}

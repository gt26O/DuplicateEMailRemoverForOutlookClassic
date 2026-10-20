namespace ContactsDoctor;

/// <summary>
/// One contact as read from a source (phone .vcf, Outlook/Exchange, PST, CSV).
/// Only the fields needed to compare contacts are kept; for vCard sources the
/// original lines are kept too so a corrected file loses nothing.
/// </summary>
public sealed class Contact
{
    /// <summary>Label of the source, e.g. "pst", "exchange", "telephone".</summary>
    public string Source { get; set; } = "";

    /// <summary>Unique id inside the source (Outlook EntryID, or index in a file).</summary>
    public string Id { get; set; } = "";

    /// <summary>Outlook StoreID (empty for files).</summary>
    public string StoreId { get; set; } = "";

    /// <summary>Folder path inside the source (Outlook), empty for files.</summary>
    public string Folder { get; set; } = "";

    public string GivenName { get; set; } = "";
    public string MiddleName { get; set; } = "";
    public string FamilyName { get; set; } = "";
    public string FullName { get; set; } = "";
    public string FileAs { get; set; } = "";
    public string Company { get; set; } = "";
    public string JobTitle { get; set; } = "";
    public List<string> Emails { get; set; } = [];
    public List<string> Phones { get; set; } = [];
    public DateTime? LastModified { get; set; }

    /// <summary>Unfolded vCard lines (only for contacts read from a .vcf).</summary>
    public List<string>? RawVCard { get; set; }

    /// <summary>Name as shown to a person: "Given Middle Family", else FullName, FileAs, company, email.</summary>
    public string DisplayName
    {
        get
        {
            string composed = string.Join(" ", new[] { GivenName, MiddleName, FamilyName }.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()));
            if (composed.Length > 0) return composed;
            if (!string.IsNullOrWhiteSpace(FullName)) return FullName.Trim();
            if (!string.IsNullOrWhiteSpace(FileAs)) return FileAs.Trim();
            if (!string.IsNullOrWhiteSpace(Company)) return Company.Trim();
            return Emails.FirstOrDefault() ?? Phones.FirstOrDefault() ?? "(sans nom)";
        }
    }

    /// <summary>Rough "how much information" score used to pick which duplicate to keep.</summary>
    public int Completeness =>
        Emails.Count * 3 + Phones.Count * 3
        + (string.IsNullOrWhiteSpace(GivenName) ? 0 : 2) + (string.IsNullOrWhiteSpace(FamilyName) ? 0 : 2)
        + (string.IsNullOrWhiteSpace(Company) ? 0 : 1) + (string.IsNullOrWhiteSpace(JobTitle) ? 0 : 1)
        + (RawVCard?.Count ?? 0) / 4;

    public override string ToString() => $"[{Source}] {DisplayName}";
}

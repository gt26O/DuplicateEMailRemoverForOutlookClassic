namespace ContactsDoctor;

public enum PlanAction
{
    /// <summary>Swap first and last name back (they are inverted compared with the reference).</summary>
    SwapName,
    /// <summary>Replace the name with the reference one (proposed, not applied by default).</summary>
    RenameFromReference,
    /// <summary>Remove an extra copy (Outlook: moved to a "Doublons" sub-folder, never deleted).</summary>
    MoveDuplicate,
    /// <summary>Re-create a contact that exists in the reference only.</summary>
    AddFromReference,
}

/// <summary>One line of the correction plan. The plan is a CSV the technician reviews in Excel
/// (column "Appliquer" OUI/NON) before running "apply".</summary>
public sealed class PlanRow
{
    public bool Apply { get; set; }
    public PlanAction Action { get; set; }
    public string Source { get; set; } = "";
    public int GroupId { get; set; }
    public string Contact { get; set; } = "";
    public string Folder { get; set; } = "";
    public string GivenBefore { get; set; } = "";
    public string MiddleBefore { get; set; } = "";
    public string FamilyBefore { get; set; } = "";
    public string GivenAfter { get; set; } = "";
    public string MiddleAfter { get; set; } = "";
    public string FamilyAfter { get; set; } = "";
    public string Reason { get; set; } = "";
    public string Id { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string KeepId { get; set; } = "";
    public string RefId { get; set; } = "";
    public string RefStoreId { get; set; } = "";
}

public static class Plan
{
    private static readonly string[] Header =
    [
        "Appliquer", "Action", "Source", "Groupe", "Contact", "Dossier",
        "Prenom_avant", "DeuxiemePrenom_avant", "Nom_avant", "Prenom_apres", "DeuxiemePrenom_apres", "Nom_apres",
        "Raison", "Id", "StoreId", "Conserver_Id", "Ref_Id", "Ref_StoreId",
    ];

    /// <summary>Builds the correction plan for one source from the analysis.</summary>
    public static List<PlanRow> Build(AnalysisResult analysis, string source)
    {
        var rows = new List<PlanRow>();
        var moved = new HashSet<Contact>(ReferenceEqualityComparer.Instance);

        foreach (Issue dup in analysis.For(source, IssueKind.Duplicate))
        {
            Contact keep = dup.Record!;
            var keepKeys = Keys(keep);
            foreach (Contact copy in dup.Others)
            {
                moved.Add(copy);
                var extra = copy.Emails.Where(e => Normalizer.Email(e) is { Length: > 0 } k && !keepKeys.Contains("e:" + k))
                    .Concat(copy.Phones.Where(p => Normalizer.Phone(p) is { Length: > 0 } k && !keepKeys.Contains("p:" + k)))
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                PlanRow row = FromRecord(PlanAction.MoveDuplicate, copy, dup.GroupId);
                row.KeepId = keep.Id;
                row.Apply = true;
                row.Reason = extra.Count == 0
                    ? $"doublon de « {keep.DisplayName} » (conservé), aucune donnée en plus"
                    : $"doublon de « {keep.DisplayName} » (conservé) ; ajouté au contact conservé : {string.Join(", ", extra)}";
                rows.Add(row);
            }
        }

        foreach (Issue i in analysis.Issues.Where(i => i.Source == source && i.Kind is IssueKind.NameSwapped or IssueKind.NameOrderDiffers or IssueKind.NameDiffers))
        {
            Contact r = i.Record!, m = i.Reference!;
            if (moved.Contains(r)) continue;
            PlanRow row;
            if (i.Kind == IssueKind.NameSwapped)
            {
                row = FromRecord(PlanAction.SwapName, r, i.GroupId);
                row.GivenAfter = r.FamilyName;
                row.MiddleAfter = r.MiddleName;
                row.FamilyAfter = r.GivenName;
                row.Apply = true;
                row.Reason = "prénom et nom inversés par rapport à la référence";
            }
            else
            {
                row = FromRecord(PlanAction.RenameFromReference, r, i.GroupId);
                row.GivenAfter = m.GivenName;
                row.MiddleAfter = m.MiddleName;
                row.FamilyAfter = m.FamilyName;
                row.Apply = false;
                row.Reason = (i.Kind == IssueKind.NameOrderDiffers ? "mêmes mots, autre ordre/champ" : "nom différent") + $" ; référence : {m.DisplayName}";
            }
            row.RefId = m.Id;
            row.RefStoreId = m.StoreId;
            rows.Add(row);
        }

        foreach (Issue i in analysis.For(source, IssueKind.Missing))
        {
            Contact m = i.Reference!;
            rows.Add(new PlanRow
            {
                Apply = true, Action = PlanAction.AddFromReference, Source = source, GroupId = i.GroupId,
                Contact = m.DisplayName, Folder = m.Folder,
                GivenAfter = m.GivenName, MiddleAfter = m.MiddleName, FamilyAfter = m.FamilyName,
                Reason = $"présent dans « {m.Source} » mais absent ici",
                RefId = m.Id, RefStoreId = m.StoreId,
            });
        }
        return rows;
    }

    /// <summary>
    /// New Outlook "Classer sous" (FileAs) after a rename: it follows the names when it was built from them
    /// ("Nom, Prénom" or "Prénom Nom"); anything else (company, custom text) is kept.
    /// </summary>
    public static string RecomputeFileAs(string fileAs, string oldFirst, string oldLast, string newFirst, string newLast)
    {
        string f = Normalizer.Text(fileAs);
        if (f.Length == 0) return fileAs;
        if (fileAs.Contains(',') && f == Normalizer.Text(oldLast + " " + oldFirst)) return $"{newLast}, {newFirst}";
        if (f == Normalizer.Text(oldFirst + " " + oldLast)) return $"{newFirst} {newLast}";
        return fileAs;
    }

    private static HashSet<string> Keys(Contact c) =>
        Normalizer.EmailKeys(c).Select(e => "e:" + e).Concat(Normalizer.PhoneKeys(c).Select(p => "p:" + p)).ToHashSet();

    private static PlanRow FromRecord(PlanAction action, Contact c, int groupId) => new()
    {
        Action = action, Source = c.Source, GroupId = groupId, Contact = c.DisplayName, Folder = c.Folder,
        GivenBefore = c.GivenName, MiddleBefore = c.MiddleName, FamilyBefore = c.FamilyName,
        Id = c.Id, StoreId = c.StoreId,
    };

    public static void Save(string path, IEnumerable<PlanRow> rows)
    {
        var lines = new List<string[]> { Header };
        lines.AddRange(rows.Select(r => new[]
        {
            r.Apply ? "OUI" : "NON", r.Action.ToString(), r.Source, r.GroupId.ToString(), r.Contact, r.Folder,
            r.GivenBefore, r.MiddleBefore, r.FamilyBefore, r.GivenAfter, r.MiddleAfter, r.FamilyAfter,
            r.Reason, r.Id, r.StoreId, r.KeepId, r.RefId, r.RefStoreId,
        }));
        Csv.WriteFile(path, lines);
    }

    public static List<PlanRow> Load(string path)
    {
        List<string[]> rows = Csv.Parse(Csv.ReadAllTextSmart(path));
        if (rows.Count == 0) return [];
        var col = rows[0].Select((h, i) => (h.Trim(), i)).ToDictionary(x => x.Item1, x => x.i, StringComparer.OrdinalIgnoreCase);
        string get(string[] r, string name) => col.TryGetValue(name, out int i) && i < r.Length ? r[i] : "";
        return rows.Skip(1).Select(r => new PlanRow
        {
            Apply = get(r, "Appliquer").Trim().ToUpperInvariant() is "OUI" or "O" or "YES" or "Y" or "1" or "X",
            Action = Enum.Parse<PlanAction>(get(r, "Action").Trim(), ignoreCase: true),
            Source = get(r, "Source"),
            GroupId = int.TryParse(get(r, "Groupe"), out int g) ? g : 0,
            Contact = get(r, "Contact"), Folder = get(r, "Dossier"),
            GivenBefore = get(r, "Prenom_avant"), MiddleBefore = get(r, "DeuxiemePrenom_avant"), FamilyBefore = get(r, "Nom_avant"),
            GivenAfter = get(r, "Prenom_apres"), MiddleAfter = get(r, "DeuxiemePrenom_apres"), FamilyAfter = get(r, "Nom_apres"),
            Reason = get(r, "Raison"), Id = get(r, "Id"), StoreId = get(r, "StoreId"), KeepId = get(r, "Conserver_Id"),
            RefId = get(r, "Ref_Id"), RefStoreId = get(r, "Ref_StoreId"),
        }).ToList();
    }

    /// <summary>
    /// Applies a plan to an in-memory list (used to produce a corrected .vcf for the phone).
    /// Duplicates are merged into the contact kept (emails/phones are not lost).
    /// </summary>
    public static List<Contact> ApplyToList(IReadOnlyList<Contact> contacts, IEnumerable<PlanRow> plan, IReadOnlyList<Contact> referenceContacts)
    {
        var result = contacts.Select(Clone).ToList();
        var byId = result.ToDictionary(c => c.Id);
        var refById = referenceContacts.GroupBy(c => c.Id).ToDictionary(g => g.Key, g => g.First());
        foreach (PlanRow row in plan.Where(r => r.Apply))
        {
            switch (row.Action)
            {
                case PlanAction.SwapName:
                case PlanAction.RenameFromReference:
                    if (byId.TryGetValue(row.Id, out Contact? c))
                    {
                        c.GivenName = row.GivenAfter;
                        c.MiddleName = row.MiddleAfter;
                        c.FamilyName = row.FamilyAfter;
                        c.FullName = "";
                    }
                    break;
                case PlanAction.MoveDuplicate:
                    if (byId.TryGetValue(row.Id, out Contact? copy))
                    {
                        if (byId.TryGetValue(row.KeepId, out Contact? keep))
                        {
                            keep.Emails.AddRange(copy.Emails.Where(e => !keep.Emails.Any(k => Normalizer.Email(k) == Normalizer.Email(e))));
                            keep.Phones.AddRange(copy.Phones.Where(p => !keep.Phones.Any(k => Normalizer.Phone(k) == Normalizer.Phone(p))));
                        }
                        result.Remove(copy);
                        byId.Remove(row.Id);
                    }
                    break;
                case PlanAction.AddFromReference:
                    if (refById.TryGetValue(row.RefId, out Contact? r))
                    {
                        Contact added = Clone(r);
                        added.Source = result.FirstOrDefault()?.Source ?? row.Source;
                        added.Id = "ajout-" + r.Id;
                        // Only keep raw vCard lines when the reference itself was a vCard.
                        if (added.RawVCard == null) added.FullName = "";
                        result.Add(added);
                    }
                    break;
            }
        }
        return result;
    }

    private static Contact Clone(Contact c) => new()
    {
        Source = c.Source, Id = c.Id, StoreId = c.StoreId, Folder = c.Folder,
        GivenName = c.GivenName, MiddleName = c.MiddleName, FamilyName = c.FamilyName, FullName = c.FullName, FileAs = c.FileAs,
        Company = c.Company, JobTitle = c.JobTitle, Emails = [.. c.Emails], Phones = [.. c.Phones], LastModified = c.LastModified,
        RawVCard = c.RawVCard == null ? null : [.. c.RawVCard],
    };
}

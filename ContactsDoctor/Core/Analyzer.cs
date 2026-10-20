namespace ContactsDoctor;

public enum IssueKind
{
    /// <summary>First and last name are swapped compared with the reference.</summary>
    NameSwapped,
    /// <summary>Same words but in another order / other fields (e.g. everything in "first name").</summary>
    NameOrderDiffers,
    /// <summary>The name is different (shorter, spelling...).</summary>
    NameDiffers,
    /// <summary>A reference contact has no counterpart in this source.</summary>
    Missing,
    /// <summary>The contact exists in this source but not in the reference (often simply created later).</summary>
    Extra,
    /// <summary>Several contacts of the same source are the same person.</summary>
    Duplicate,
    /// <summary>Emails/phones of the reference contact are missing in this source.</summary>
    DataMissing,
    /// <summary>The contact carries emails/phones that belong to another person of the reference (bad merge?).</summary>
    SuspectedMerge,
}

public sealed class Issue
{
    public IssueKind Kind { get; init; }
    /// <summary>Source the issue is about.</summary>
    public string Source { get; init; } = "";
    public int GroupId { get; init; }
    /// <summary>The record of <see cref="Source"/> concerned (null for <see cref="IssueKind.Missing"/>).</summary>
    public Contact? Record { get; init; }
    /// <summary>The matching reference record, when there is one.</summary>
    public Contact? Reference { get; init; }
    /// <summary>Other records involved (the extra copies of a duplicate, the other person of a suspected merge).</summary>
    public List<Contact> Others { get; init; } = [];
    public string Detail { get; init; } = "";
}

public sealed class SourceSummary
{
    public string Source { get; init; } = "";
    public int Total { get; set; }
    public int People { get; set; }
    public Dictionary<IssueKind, int> Counts { get; } = Enum.GetValues<IssueKind>().ToDictionary(k => k, _ => 0);
    /// <summary>Number of extra copies (a person present 3 times counts 2).</summary>
    public int DuplicateCopies { get; set; }
    public SortedDictionary<string, int> Folders { get; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class ContactGroup
{
    public int Id { get; init; }
    public List<Contact> Members { get; } = [];
    public IEnumerable<Contact> In(string source) => Members.Where(m => m.Source == source);
}

public sealed class AnalysisResult
{
    public string Reference { get; init; } = "";
    public List<string> Sources { get; init; } = [];
    public List<ContactGroup> Groups { get; init; } = [];
    public List<Issue> Issues { get; init; } = [];
    public Dictionary<string, SourceSummary> Summaries { get; init; } = [];
    /// <summary>Emails/phones shared by many different people (switchboard...), not used to match.</summary>
    public List<string> IgnoredSharedKeys { get; init; } = [];

    public IEnumerable<Issue> For(string source, IssueKind kind) => Issues.Where(i => i.Source == source && i.Kind == kind);
}

/// <summary>
/// Groups the contacts of all sources by person, then compares every source with the reference
/// (typically the PST from before the migration).
/// </summary>
public static class Analyzer
{
    /// <summary>A phone/email found on this many different names is treated as shared (switchboard, family line).</summary>
    public const int SharedKeyThreshold = 3;

    public static AnalysisResult Analyze(string reference, IReadOnlyList<Contact> contacts)
    {
        var sources = contacts.Select(c => c.Source).Distinct().ToList();
        if (!sources.Contains(reference)) sources.Insert(0, reference);
        else { sources.Remove(reference); sources.Insert(0, reference); }

        int n = contacts.Count;
        string[][] tokens = contacts.Select(c => Normalizer.NameTokens(c)).ToArray();
        string[] nameKeys = contacts.Select(Normalizer.NameKey).ToArray();
        List<string>[] keys = contacts.Select(ContactKeys).ToArray();

        // key -> records carrying it
        var index = new Dictionary<string, List<int>>();
        for (int i = 0; i < n; i++)
            foreach (string k in keys[i])
                (index.TryGetValue(k, out var l) ? l : index[k] = []).Add(i);

        var ignored = new List<string>();
        foreach ((string k, List<int> recs) in index)
        {
            int distinctNames = recs.Select(r => nameKeys[r]).Where(s => s.Length > 0).Distinct().Count();
            if (distinctNames >= SharedKeyThreshold) ignored.Add(k);
        }
        var ignoredSet = ignored.ToHashSet();

        // Union records that share an email/phone and have compatible names, or have the same full name.
        var uf = new UnionFind(n);
        foreach ((string k, List<int> recs) in index)
        {
            if (ignoredSet.Contains(k)) continue;
            bool isName = k.StartsWith("n:");
            for (int a = 0; a < recs.Count; a++)
                for (int b = a + 1; b < recs.Count; b++)
                    if (isName || Compatible(tokens[recs[a]], tokens[recs[b]]))
                        uf.Union(recs[a], recs[b]);
        }

        var groups = Enumerable.Range(0, n).GroupBy(uf.Find)
            .Select((g, gi) =>
            {
                var group = new ContactGroup { Id = gi + 1 };
                group.Members.AddRange(g.Select(i => contacts[i]));
                return group;
            }).ToList();

        var keysOf = new Dictionary<Contact, HashSet<string>>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < n; i++) keysOf[contacts[i]] = keys[i].Where(k => !k.StartsWith("n:")).ToHashSet();
        var groupOf = new Dictionary<Contact, ContactGroup>(ReferenceEqualityComparer.Instance);
        foreach (ContactGroup g in groups) foreach (Contact c in g.Members) groupOf[c] = g;

        var issues = new List<Issue>();
        foreach (ContactGroup g in groups)
        {
            List<Contact> refs = g.In(reference).ToList();
            foreach (string source in sources)
            {
                List<Contact> recs = g.In(source).ToList();

                if (recs.Count > 1)
                {
                    Contact keep = recs.OrderByDescending(r => r.Completeness).ThenByDescending(r => r.LastModified ?? DateTime.MinValue).First();
                    issues.Add(new Issue
                    {
                        Kind = IssueKind.Duplicate, Source = source, GroupId = g.Id, Record = keep,
                        Reference = refs.FirstOrDefault(),
                        Others = recs.Where(r => !ReferenceEquals(r, keep)).ToList(),
                        Detail = $"{recs.Count} exemplaires",
                    });
                }
                if (source == reference) continue;

                if (refs.Count == 0)
                {
                    foreach (Contact r in recs)
                        issues.Add(new Issue { Kind = IssueKind.Extra, Source = source, GroupId = g.Id, Record = r, Detail = "absent de la référence" });
                    continue;
                }
                if (recs.Count == 0)
                {
                    Contact best = refs.OrderByDescending(r => r.Completeness).First();
                    issues.Add(new Issue { Kind = IssueKind.Missing, Source = source, GroupId = g.Id, Reference = best, Detail = "absent de cette source" });
                    continue;
                }

                foreach (Contact r in recs)
                {
                    Contact m = refs.OrderByDescending(x => MatchScore(r, x, keysOf)).First();
                    IssueKind? nameIssue = CompareNames(r, m);
                    if (nameIssue != null)
                        issues.Add(new Issue
                        {
                            Kind = nameIssue.Value, Source = source, GroupId = g.Id, Record = r, Reference = m,
                            Detail = $"« {Describe(r)} » au lieu de « {Describe(m)} »",
                        });

                    // Data of another reference person on this record: typical of a wrong merge on the phone.
                    var foreign = new List<Contact>();
                    foreach (string k in keysOf[r].Where(k => !keysOf[m].Contains(k) && !ignoredSet.Contains(k)))
                        foreach (int o in index[k])
                        {
                            Contact other = contacts[o];
                            if (other.Source == reference && !ReferenceEquals(groupOf[other], g) && !foreign.Contains(other))
                                foreign.Add(other);
                        }
                    if (foreign.Count > 0)
                        issues.Add(new Issue
                        {
                            Kind = IssueKind.SuspectedMerge, Source = source, GroupId = g.Id, Record = r, Reference = m, Others = foreign,
                            Detail = "contient aussi des coordonnées de : " + string.Join(", ", foreign.Select(f => f.DisplayName)),
                        });
                }

                var present = recs.SelectMany(r => keysOf[r]).ToHashSet();
                var lost = new List<string>();
                foreach (Contact m in refs)
                {
                    lost.AddRange(m.Emails.Where(e => Normalizer.Email(e) is { Length: > 0 } k && !present.Contains("e:" + k)));
                    lost.AddRange(m.Phones.Where(p => Normalizer.Phone(p) is { Length: > 0 } k && !present.Contains("p:" + k)));
                }
                lost = lost.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if (lost.Count > 0)
                    issues.Add(new Issue
                    {
                        Kind = IssueKind.DataMissing, Source = source, GroupId = g.Id,
                        Record = recs.OrderByDescending(r => r.Completeness).First(), Reference = refs[0],
                        Detail = "manque : " + string.Join(", ", lost),
                    });
            }
        }

        var summaries = sources.ToDictionary(s => s, s => new SourceSummary { Source = s });
        foreach (Contact c in contacts)
        {
            SourceSummary sum = summaries[c.Source];
            sum.Total++;
            string folder = c.Folder.Length > 0 ? c.Folder : "(fichier)";
            sum.Folders[folder] = sum.Folders.GetValueOrDefault(folder) + 1;
        }
        foreach (string s in sources) summaries[s].People = groups.Count(g => g.In(s).Any());
        foreach (Issue i in issues)
        {
            summaries[i.Source].Counts[i.Kind]++;
            if (i.Kind == IssueKind.Duplicate) summaries[i.Source].DuplicateCopies += i.Others.Count;
        }

        return new AnalysisResult
        {
            Reference = reference, Sources = sources, Groups = groups, Issues = issues, Summaries = summaries,
            IgnoredSharedKeys = ignored.Select(k => k[2..]).OrderBy(k => k).ToList(),
        };
    }

    private static List<string> ContactKeys(Contact c)
    {
        var keys = new List<string>();
        keys.AddRange(Normalizer.EmailKeys(c).Select(e => "e:" + e));
        keys.AddRange(Normalizer.PhoneKeys(c).Select(p => "p:" + p));
        string[] t = Normalizer.NameTokens(c);
        if (t.Length >= 2) keys.Add("n:" + Normalizer.NameKey(c));
        return keys;
    }

    /// <summary>Names that can be the same person: one is empty, or one's words are all in the other.</summary>
    public static bool Compatible(string[] a, string[] b)
    {
        if (a.Length == 0 || b.Length == 0) return true;
        return a.All(b.Contains) || b.All(a.Contains);
    }

    private static int MatchScore(Contact r, Contact m, Dictionary<Contact, HashSet<string>> keysOf)
    {
        int score = keysOf[r].Count(keysOf[m].Contains) * 3;
        if (Normalizer.NameKey(r) == Normalizer.NameKey(m)) score += 4;
        if (Normalizer.OrderedName(r) == Normalizer.OrderedName(m)) score += 1;
        return score;
    }

    /// <summary>Null when the names are equivalent.</summary>
    public static IssueKind? CompareNames(Contact r, Contact m)
    {
        if (Normalizer.OrderedName(r) == Normalizer.OrderedName(m)) return null;
        string rg = Normalizer.Text(r.GivenName), rf = Normalizer.Text(r.FamilyName);
        string mg = Normalizer.Text(m.GivenName), mf = Normalizer.Text(m.FamilyName);
        if (rg.Length > 0 && rf.Length > 0 && rg != rf && rg == mf && rf == mg) return IssueKind.NameSwapped;
        if (Normalizer.NameKey(r) == Normalizer.NameKey(m)) return IssueKind.NameOrderDiffers;
        return IssueKind.NameDiffers;
    }

    /// <summary>"Prénom: X / Nom: Y", or the full name when the structured fields are empty.</summary>
    public static string Describe(Contact c)
    {
        if (c.GivenName.Length == 0 && c.FamilyName.Length == 0) return c.DisplayName;
        return $"Prénom: {c.GivenName} / Nom: {c.FamilyName}";
    }

    private sealed class UnionFind(int n)
    {
        private readonly int[] parent = Enumerable.Range(0, n).ToArray();
        public int Find(int x)
        {
            while (parent[x] != x) x = parent[x] = parent[parent[x]];
            return x;
        }
        public void Union(int a, int b) => parent[Find(a)] = Find(b);
    }
}

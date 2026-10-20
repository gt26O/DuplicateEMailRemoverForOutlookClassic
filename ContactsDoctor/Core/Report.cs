using System.Net;
using System.Text;

namespace ContactsDoctor;

/// <summary>Human readable output of an analysis: conclusions, HTML report, CSV of all issues.</summary>
public static class Report
{
    public static string Label(IssueKind k) => k switch
    {
        IssueKind.NameSwapped => "Prénom / nom inversés",
        IssueKind.NameOrderDiffers => "Nom dans un autre ordre ou un autre champ",
        IssueKind.NameDiffers => "Nom différent",
        IssueKind.Missing => "Manquant (présent dans la référence)",
        IssueKind.Extra => "En plus (absent de la référence)",
        IssueKind.Duplicate => "Doublons",
        IssueKind.DataMissing => "E-mails / numéros perdus",
        IssueKind.SuspectedMerge => "Fusion suspecte (mélange de 2 personnes)",
        _ => k.ToString(),
    };

    /// <summary>Plain-language diagnosis: where is the problem (reference vs Exchange vs phone).</summary>
    public static List<string> Conclusions(AnalysisResult a)
    {
        var lines = new List<string>();
        string reference = a.Reference;
        var targets = a.Sources.Where(s => s != reference).ToList();
        SourceSummary refSum = a.Summaries[reference];
        lines.Add($"Référence « {reference} » : {refSum.Total} fiches pour {refSum.People} personnes"
                  + (refSum.DuplicateCopies > 0 ? $" (déjà {refSum.DuplicateCopies} doublon(s) avant la migration)." : "."));

        var swappedGroups = new Dictionary<string, HashSet<int>>();
        foreach (string s in targets)
        {
            SourceSummary sum = a.Summaries[s];
            var c = sum.Counts;
            swappedGroups[s] = a.For(s, IssueKind.NameSwapped).Select(i => i.GroupId).ToHashSet();
            lines.Add($"« {s} » : {sum.Total} fiches pour {sum.People} personnes. "
                      + $"{c[IssueKind.NameSwapped]} inversion(s) prénom/nom ({Pct(c[IssueKind.NameSwapped], sum.People)}), "
                      + $"{sum.DuplicateCopies} doublon(s) en trop, "
                      + $"{c[IssueKind.Missing]} manquant(s), {c[IssueKind.Extra]} en plus, "
                      + $"{c[IssueKind.DataMissing]} fiche(s) avec e-mails/numéros perdus, "
                      + $"{c[IssueKind.SuspectedMerge]} fusion(s) suspecte(s).");
            if (c[IssueKind.Missing] > refSum.People / 5 && refSum.People > 0)
                lines.Add($"⚠ Beaucoup de contacts de la référence manquent dans « {s} » : vérifier que le bon dossier / le bon compte a été exporté "
                          + "(sur Samsung : exporter TOUS les comptes : Téléphone, Samsung account, Exchange, Google).");
        }

        var withSwaps = targets.Where(s => swappedGroups[s].Count > 0).ToList();
        if (withSwaps.Count == 0)
            lines.Add("Aucune inversion prénom/nom trouvée par rapport à la référence.");
        else if (withSwaps.Count == 1)
            lines.Add($"→ Les inversions prénom/nom ne concernent que « {withSwaps[0]} » : c'est là qu'il faut corriger.");
        else
        {
            for (int i = 0; i < withSwaps.Count; i++)
                for (int j = i + 1; j < withSwaps.Count; j++)
                {
                    var x = swappedGroups[withSwaps[i]];
                    var y = swappedGroups[withSwaps[j]];
                    int common = x.Intersect(y).Count();
                    int smaller = Math.Min(x.Count, y.Count);
                    if (smaller > 0 && common * 10 >= smaller * 7)
                        lines.Add($"→ Ce sont les mêmes contacts qui sont inversés dans « {withSwaps[i]} » et « {withSwaps[j]} » ({common} en commun) : "
                                  + "les deux sont synchronisés. Corriger Exchange suffit si le téléphone synchronise ses contacts avec le compte Exchange.");
                    else
                        lines.Add($"→ « {withSwaps[i]} » et « {withSwaps[j]} » ont des inversions différentes ({common} en commun sur {smaller}) : "
                                  + "les deux carnets ne sont pas (ou pas entièrement) synchronisés, il faut corriger chacun.");
                }
        }

        foreach (string s in a.Sources)
        {
            SourceSummary sum = a.Summaries[s];
            if (sum.DuplicateCopies > 0 && s != reference)
            {
                int before = refSum.DuplicateCopies;
                lines.Add(sum.DuplicateCopies > before
                    ? $"→ « {s} » contient plus de doublons que la référence ({sum.DuplicateCopies} contre {before}) : la migration / synchronisation a créé des copies."
                    : $"→ « {s} » contient {sum.DuplicateCopies} doublon(s), pas plus que la référence.");
            }
            if (sum.Counts[IssueKind.SuspectedMerge] > 0)
                lines.Add($"→ « {s} » : {sum.Counts[IssueKind.SuspectedMerge]} contact(s) contiennent les coordonnées d'une autre personne : "
                          + "typique d'une fusion automatique (outil « Fusionner les contacts » de Samsung). À contrôler à la main.");
        }
        if (a.IgnoredSharedKeys.Count > 0)
            lines.Add($"Info : {a.IgnoredSharedKeys.Count} numéro(s)/e-mail(s) partagés par 3 personnes ou plus (standard, adresse générique) n'ont pas servi au rapprochement.");
        return lines;
    }

    private static string Pct(int part, int total) => total == 0 ? "0 %" : $"{100.0 * part / total:0.#} %";

    public static void WriteIssuesCsv(string path, AnalysisResult a)
    {
        var rows = new List<string[]>
        {
            new[] { "Groupe", "Source", "Probleme", "Contact", "Dossier", "Reference", "Detail", "Emails", "Telephones", "Id" },
        };
        foreach (Issue i in a.Issues.OrderBy(i => i.GroupId).ThenBy(i => i.Source))
        {
            Contact? c = i.Record ?? i.Reference;
            rows.Add([
                i.GroupId.ToString(), i.Source, Label(i.Kind), i.Record?.DisplayName ?? "", c?.Folder ?? "",
                i.Reference?.DisplayName ?? "", i.Detail + (i.Others.Count > 0 && i.Kind == IssueKind.Duplicate ? " ; copies : " + string.Join(" | ", i.Others.Select(o => o.DisplayName)) : ""),
                string.Join(", ", c?.Emails ?? []), string.Join(", ", c?.Phones ?? []), c?.Id ?? "",
            ]);
        }
        Csv.WriteFile(path, rows);
    }

    /// <summary>One line per person and source, to compare everything side by side in Excel.</summary>
    public static void WriteGroupsCsv(string path, AnalysisResult a)
    {
        var rows = new List<string[]> { new[] { "Groupe", "Source", "Dossier", "Prenom", "DeuxiemePrenom", "Nom", "NomComplet", "ClasserSous", "Societe", "Emails", "Telephones", "Id" } };
        foreach (ContactGroup g in a.Groups)
            foreach (Contact c in g.Members.OrderBy(m => a.Sources.IndexOf(m.Source)))
                rows.Add([g.Id.ToString(), c.Source, c.Folder, c.GivenName, c.MiddleName, c.FamilyName, c.FullName, c.FileAs, c.Company,
                          string.Join(", ", c.Emails), string.Join(", ", c.Phones), c.Id]);
        Csv.WriteFile(path, rows);
    }

    public static string Html(AnalysisResult a, DateTime generated)
    {
        var sb = new StringBuilder();
        string e(string? s) => WebUtility.HtmlEncode(s ?? "");
        sb.Append("<!doctype html><html lang=\"fr\"><head><meta charset=\"utf-8\"><title>Diagnostic des contacts</title><style>")
          .Append("body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222;background:#fff}h1{font-size:22px}h2{font-size:18px;margin-top:32px;border-bottom:1px solid #ccc}")
          .Append("table{border-collapse:collapse;margin:8px 0;font-size:13px}td,th{border:1px solid #ccc;padding:4px 8px;text-align:left;vertical-align:top}th{background:#f2f2f2}")
          .Append("td.n{text-align:right}.warn{color:#a40}.note{color:#555;font-size:13px}li{margin:4px 0}details{margin:8px 0}summary{cursor:pointer;font-weight:600}")
          .Append("</style></head><body>");
        sb.Append($"<h1>Diagnostic des contacts</h1><p class=\"note\">Généré le {generated:dd.MM.yyyy HH:mm}. Référence : <b>{e(a.Reference)}</b>. "
                  + "Ce rapport contient des données personnelles : ne pas le diffuser.</p>");

        sb.Append("<h2>Conclusions</h2><ul>");
        foreach (string line in Conclusions(a)) sb.Append($"<li>{e(line)}</li>");
        sb.Append("</ul>");

        sb.Append("<h2>Résumé par source</h2><table><tr><th>Source</th><th>Fiches</th><th>Personnes</th><th>Doublons en trop</th>");
        IssueKind[] cols = [IssueKind.NameSwapped, IssueKind.NameOrderDiffers, IssueKind.NameDiffers, IssueKind.Missing, IssueKind.Extra, IssueKind.DataMissing, IssueKind.SuspectedMerge];
        foreach (IssueKind k in cols) sb.Append($"<th>{e(Label(k))}</th>");
        sb.Append("</tr>");
        foreach (string s in a.Sources)
        {
            SourceSummary sum = a.Summaries[s];
            sb.Append($"<tr><td>{e(s)}{(s == a.Reference ? " (référence)" : "")}</td><td class=n>{sum.Total}</td><td class=n>{sum.People}</td><td class=n>{sum.DuplicateCopies}</td>");
            foreach (IssueKind k in cols) sb.Append(s == a.Reference ? "<td class=n>–</td>" : $"<td class=n>{sum.Counts[k]}</td>");
            sb.Append("</tr>");
        }
        sb.Append("</table>");

        sb.Append("<h2>Dossiers lus</h2><table><tr><th>Source</th><th>Dossier</th><th>Fiches</th></tr>");
        foreach (string s in a.Sources)
            foreach ((string folder, int count) in a.Summaries[s].Folders)
                sb.Append($"<tr><td>{e(s)}</td><td>{e(folder)}</td><td class=n>{count}</td></tr>");
        sb.Append("</table>");

        IssueKind[] order = [IssueKind.NameSwapped, IssueKind.Duplicate, IssueKind.Missing, IssueKind.SuspectedMerge, IssueKind.DataMissing, IssueKind.NameOrderDiffers, IssueKind.NameDiffers, IssueKind.Extra];
        foreach (IssueKind k in order)
        {
            var list = a.Issues.Where(i => i.Kind == k).OrderBy(i => i.Source).ThenBy(i => (i.Record ?? i.Reference)?.DisplayName).ToList();
            if (list.Count == 0) continue;
            sb.Append($"<h2>{e(Label(k))} ({list.Count})</h2><details{(list.Count <= 100 ? " open" : "")}><summary>Afficher / masquer</summary>");
            sb.Append("<table><tr><th>Groupe</th><th>Source</th><th>Contact</th><th>Dossier</th><th>Référence</th><th>Détail</th></tr>");
            foreach (Issue i in list)
            {
                string detail = i.Detail;
                if (k == IssueKind.Duplicate) detail += " — conservé : " + i.Record!.DisplayName + " ; copies : " + string.Join(" | ", i.Others.Select(o => o.DisplayName + (o.Folder.Length > 0 ? $" [{o.Folder}]" : "")));
                sb.Append($"<tr><td class=n>{i.GroupId}</td><td>{e(i.Source)}</td><td>{e(i.Record?.DisplayName ?? "")}</td><td>{e((i.Record ?? i.Reference)?.Folder)}</td>"
                          + $"<td>{e(i.Reference?.DisplayName ?? "")}</td><td>{e(detail)}</td></tr>");
            }
            sb.Append("</table></details>");
        }
        sb.Append("<p class=\"note\">Le numéro de groupe permet de retrouver toutes les fiches d'une même personne dans <i>groupes.csv</i>.</p>");
        sb.Append("</body></html>");
        return sb.ToString();
    }
}

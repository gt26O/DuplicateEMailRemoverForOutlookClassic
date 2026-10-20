using System.Text;

namespace ContactsDoctor;

internal static class Program
{
    private const string Help = """
        ContactsDoctor — diagnostic et réparation des contacts (téléphone .vcf, Exchange, PST)

        Sources (SPEC) :
          outlook               boîte Outlook par défaut (Exchange / Microsoft 365), dossiers de contacts
          outlook:<nom>         une autre boîte ou un PST déjà ouvert dans Outlook (nom affiché par « stores »)
          C:\...\ancien.pst     fichier PST (ouvert temporairement dans Outlook, en lecture)
          C:\...\contacts.vcf   export du téléphone (Samsung Contacts > Gérer les contacts > Exporter)
          C:\...\contacts.csv   export CSV Outlook / Google
          C:\...\source.json    instantané créé par « export » ou « analyze »

        Commandes :
          stores
              Liste les boîtes et fichiers ouverts dans Outlook avec leur nombre de contacts.

          export <SPEC> <fichier.json|.vcf|.csv>
              Copie les contacts d'une source dans un fichier (photo de l'état actuel).

          analyze --ref <nom>=<SPEC> --source <nom>=<SPEC> [--source ...] [--out <dossier>]
              Compare toutes les sources à la référence (en général le PST d'avant la migration).
              Produit : rapport.html, problemes.csv, groupes.csv, plan-<nom>.csv, <nom>-corrige.vcf (sources .vcf),
              et un instantané JSON de chaque source.

          apply <plan-xxx.csv> [--pst <ancien.pst>] [--target <SPEC>] [--execute]
              Applique dans Outlook/Exchange les lignes du plan marquées « OUI ».
              Sans --execute : simulation, rien n'est modifié. Avec --execute : chaque contact est sauvegardé
              en .msg avant modification, les doublons sont DÉPLACÉS dans « Doublons (ContactsDoctor) », jamais supprimés.
              --pst : ouvre le PST de référence (nécessaire pour les ajouts « AddFromReference »).

          undo <journal.csv> [--pst <ancien.pst>] [--execute]
              Annule ce qu'un « apply --execute » a fait (d'après son journal).

        Exemple :
          ContactsDoctor analyze --ref pst=D:\Client\ancien.pst --source exchange=outlook --source telephone=D:\Client\Contacts.vcf --out D:\Client\diag
        """;

    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (args.Length == 0 || args[0] is "-h" or "--help" or "help" or "/?")
        {
            Console.WriteLine(Help);
            return 0;
        }
        try
        {
            return args[0].ToLowerInvariant() switch
            {
                "stores" => Stores(),
                "export" => Export(args[1..]),
                "analyze" or "analyse" => Analyze(args[1..]),
                "apply" => Apply(args[1..]),
                "undo" => Undo(args[1..]),
                _ => Fail($"Commande inconnue « {args[0]} ». Tapez ContactsDoctor --help"),
            };
        }
        catch (Exception ex)
        {
            return Fail(ex.Message);
        }
    }

    private static int Fail(string message)
    {
        Console.Error.WriteLine("Erreur : " + message);
        return 1;
    }

    private static void Log(string s) => Console.WriteLine(s);

    private static int Stores()
    {
        using var session = new OutlookSession();
        foreach (object store in session.Stores())
        {
            string name = OutlookSession.Str(store, "DisplayName");
            string path = OutlookSession.Str(store, "FilePath");
            int contacts = 0;
            try
            {
                foreach ((object folder, string _) in session.ContactFolders(store))
                    contacts += (int)OutlookSession.Get(OutlookSession.Get(folder, "Items")!, "Count")!;
            }
            catch (Exception) { /* store not available (offline shared mailbox...) */ }
            Log($"{name}  —  {contacts} contact(s){(path.Length > 0 ? "  —  " + path : "")}");
        }
        return 0;
    }

    /// <summary>Loads a source; Outlook is only started when a source needs it.</summary>
    private static List<Contact> Load(string spec, string name, ref OutlookSession? session)
    {
        string ext = Path.GetExtension(spec).ToLowerInvariant();
        List<Contact> contacts;
        switch (ext)
        {
            case ".vcf":
                contacts = VCard.ReadFile(spec, name);
                break;
            case ".csv":
                contacts = ContactCsv.ReadFile(spec, name);
                break;
            case ".json":
                contacts = Snapshot.Load(spec, name);
                break;
            default:
                session ??= new OutlookSession();
                string storeSpec = spec.StartsWith("outlook:", StringComparison.OrdinalIgnoreCase) ? spec[8..] : spec;
                object store = session.OpenStore(storeSpec);
                Log($"Lecture de « {OutlookSession.Str(store, "DisplayName")} » ({name}) :");
                contacts = session.ReadContacts(store, name, Log);
                break;
        }
        Log($"{name} : {contacts.Count} contact(s) lus depuis {spec}");
        return contacts;
    }

    private static int Export(string[] args)
    {
        if (args.Length < 2) return Fail("usage : export <SPEC> <fichier.json|.vcf|.csv>");
        OutlookSession? session = null;
        try
        {
            List<Contact> contacts = Load(args[0], "export", ref session);
            Write(args[1], contacts);
            Log($"Écrit : {Path.GetFullPath(args[1])}");
            return 0;
        }
        finally { session?.Dispose(); }
    }

    private static void Write(string path, List<Contact> contacts)
    {
        switch (Path.GetExtension(path).ToLowerInvariant())
        {
            case ".vcf":
                File.WriteAllText(path, VCard.Write(contacts), new UTF8Encoding(false));
                break;
            case ".csv":
                var rows = new List<string[]> { new[] { "Prénom", "Deuxième prénom", "Nom", "Nom complet", "Classer sous", "Société", "Fonction", "E-mails", "Téléphones", "Dossier" } };
                rows.AddRange(contacts.Select(c => new[] { c.GivenName, c.MiddleName, c.FamilyName, c.FullName, c.FileAs, c.Company, c.JobTitle,
                    string.Join(" ::: ", c.Emails), string.Join(" ::: ", c.Phones), c.Folder }));
                Csv.WriteFile(path, rows);
                break;
            default:
                Snapshot.Save(path, contacts);
                break;
        }
    }

    private static (string Name, string Spec) NameAndSpec(string value)
    {
        int eq = value.IndexOf('=');
        if (eq <= 0) throw new ArgumentException($"« {value} » : attendu nom=source, par ex. exchange=outlook");
        return (value[..eq].Trim(), value[(eq + 1)..].Trim().Trim('"'));
    }

    private static int Analyze(string[] args)
    {
        (string Name, string Spec)? reference = null;
        var sources = new List<(string Name, string Spec)>();
        string outDir = Path.Combine(Environment.CurrentDirectory, $"diagnostic_{DateTime.Now:yyyyMMdd_HHmmss}");
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--ref": reference = NameAndSpec(args[++i]); break;
                case "--source": sources.Add(NameAndSpec(args[++i])); break;
                case "--out": outDir = args[++i]; break;
                default: return Fail($"option inconnue « {args[i]} »");
            }
        }
        if (reference == null) return Fail("--ref est obligatoire (la source qui fait foi, en général le PST d'avant la migration)");
        if (sources.Count == 0) return Fail("au moins une --source est nécessaire");
        var all = new List<(string Name, string Spec)> { reference.Value };
        all.AddRange(sources);
        if (all.Select(s => s.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != all.Count) return Fail("chaque source doit avoir un nom différent");

        Directory.CreateDirectory(outDir);
        Directory.CreateDirectory(Path.Combine(outDir, "instantanes"));
        var contactsBySource = new Dictionary<string, List<Contact>>();
        OutlookSession? session = null;
        try
        {
            foreach ((string name, string spec) in all)
            {
                contactsBySource[name] = Load(spec, name, ref session);
                Snapshot.Save(Path.Combine(outDir, "instantanes", name + ".json"), contactsBySource[name]);
            }
        }
        finally { session?.Dispose(); }

        AnalysisResult result = Analyzer.Analyze(reference.Value.Name, contactsBySource.Values.SelectMany(c => c).ToList());
        File.WriteAllText(Path.Combine(outDir, "rapport.html"), Report.Html(result, DateTime.Now), new UTF8Encoding(false));
        Report.WriteIssuesCsv(Path.Combine(outDir, "problemes.csv"), result);
        Report.WriteGroupsCsv(Path.Combine(outDir, "groupes.csv"), result);

        foreach ((string name, string spec) in sources)
        {
            List<PlanRow> plan = Plan.Build(result, name);
            Plan.Save(Path.Combine(outDir, $"plan-{name}.csv"), plan);
            if (Path.GetExtension(spec).Equals(".vcf", StringComparison.OrdinalIgnoreCase))
            {
                List<Contact> fixedList = Plan.ApplyToList(contactsBySource[name], plan, contactsBySource[reference.Value.Name]);
                File.WriteAllText(Path.Combine(outDir, $"{name}-corrige.vcf"), VCard.Write(fixedList), new UTF8Encoding(false));
            }
        }

        Log("");
        foreach (string line in Report.Conclusions(result)) Log("• " + line);
        Log("");
        Log($"Résultats dans : {Path.GetFullPath(outDir)}");
        Log("Ouvrez rapport.html, puis relisez plan-<source>.csv dans Excel (colonne « Appliquer » OUI/NON) avant « apply ».");
        Log("ATTENTION : ces fichiers contiennent des données personnelles du client.");
        return 0;
    }

    private static int Apply(string[] args)
    {
        if (args.Length < 1) return Fail("usage : apply <plan.csv> [--pst ancien.pst] [--target SPEC] [--execute]");
        string planPath = args[0];
        bool execute = false;
        string target = "";
        var psts = new List<string>();
        for (int i = 1; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--execute": execute = true; break;
                case "--pst": psts.Add(args[++i]); break;
                case "--target": target = args[++i]; break;
                default: return Fail($"option inconnue « {args[i]} »");
            }
        }
        List<PlanRow> plan = Plan.Load(planPath);
        if (execute)
        {
            Console.Write($"{plan.Count(r => r.Apply)} modification(s) vont être faites dans Outlook. Tapez OUI pour continuer : ");
            if (!string.Equals(Console.ReadLine()?.Trim(), "OUI", StringComparison.OrdinalIgnoreCase)) return Fail("annulé par l'utilisateur");
        }
        using var session = new OutlookSession();
        foreach (string pst in psts) session.OpenStore(pst);
        var fixer = new OutlookFixer(session, Path.GetDirectoryName(Path.GetFullPath(planPath))!, execute, Log);
        fixer.Apply(plan, target.StartsWith("outlook:", StringComparison.OrdinalIgnoreCase) ? target[8..] : target);
        return fixer.Failed == 0 ? 0 : 2;
    }

    private static int Undo(string[] args)
    {
        if (args.Length < 1) return Fail("usage : undo <journal.csv> [--pst ancien.pst] [--execute]");
        bool execute = args.Contains("--execute", StringComparer.OrdinalIgnoreCase);
        using var session = new OutlookSession();
        for (int i = 1; i < args.Length; i++)
            if (args[i].Equals("--pst", StringComparison.OrdinalIgnoreCase)) session.OpenStore(args[++i]);
        var fixer = new OutlookFixer(session, Path.GetDirectoryName(Path.GetFullPath(args[0]))!, execute, Log);
        fixer.Undo(args[0]);
        return fixer.Failed == 0 ? 0 : 2;
    }
}

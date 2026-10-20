using static ContactsDoctor.OutlookSession;

namespace ContactsDoctor;

/// <summary>
/// Applies a reviewed plan to Outlook/Exchange, and undoes it. Every change is preceded by a .msg backup
/// of the contact and written to a journal; nothing is ever deleted (duplicates are moved to a sub-folder).
/// Without <c>execute</c> it only prints what would be done.
/// </summary>
public sealed class OutlookFixer(OutlookSession session, string workDir, bool execute, Action<string> log)
{
    private const int OlMsgUnicode = 9;

    private static readonly string[] JournalHeader =
    [
        "Date", "Action", "Contact", "Id_avant", "StoreId", "Id_apres", "DossierOrigine_Id",
        "Prenom_avant", "DeuxiemePrenom_avant", "Nom_avant", "ClasserSous_avant",
        "Prenom_apres", "DeuxiemePrenom_apres", "Nom_apres", "ClasserSous_apres", "Sauvegarde", "Resultat",
        "Conserve_Id", "Champs_ajoutes",
    ];

    private readonly List<string[]> journal = [JournalHeader];
    private string? backupDir;

    public int Done { get; private set; }
    public int Skipped { get; private set; }
    public int Failed { get; private set; }

    public string? JournalPath { get; private set; }

    public void Apply(IReadOnlyList<PlanRow> plan, string targetStoreSpec)
    {
        object? targetFolder = null;
        foreach (PlanRow row in plan.Where(r => r.Apply))
        {
            try
            {
                switch (row.Action)
                {
                    case PlanAction.SwapName:
                    case PlanAction.RenameFromReference:
                        Rename(row);
                        break;
                    case PlanAction.MoveDuplicate:
                        MoveDuplicate(row);
                        break;
                    case PlanAction.AddFromReference:
                        targetFolder ??= session.DefaultContactsFolder(session.OpenStore(targetStoreSpec));
                        AddFromReference(row, targetFolder);
                        break;
                }
            }
            catch (Exception ex)
            {
                Failed++;
                log($"  ERREUR {row.Action} « {row.Contact} » : {ex.Message}");
                Journal(row.Action.ToString(), row.Contact, row.Id, row.StoreId, "", "", null, null, "", "ERREUR " + ex.Message);
            }
        }
        int notSelected = plan.Count(r => !r.Apply);
        log($"Plan : {plan.Count} ligne(s), {notSelected} non sélectionnée(s) (Appliquer = NON).");
        Finish("journal");
    }

    private void Rename(PlanRow row)
    {
        object item = session.GetItem(row.Id, row.StoreId);
        string first = Str(item, "FirstName"), middle = Str(item, "MiddleName"), last = Str(item, "LastName"), fileAs = Str(item, "FileAs");
        if (Normalizer.Text(first) != Normalizer.Text(row.GivenBefore) || Normalizer.Text(last) != Normalizer.Text(row.FamilyBefore))
        {
            Skipped++;
            log($"  ignoré « {row.Contact} » : le contact a changé depuis l'analyse (actuellement {first} / {last}).");
            return;
        }
        string newFileAs = Plan.RecomputeFileAs(fileAs, first, last, row.GivenAfter, row.FamilyAfter);
        log($"  {(execute ? "" : "[simulation] ")}{row.Action} « {first} {last} » → prénom « {row.GivenAfter} », nom « {row.FamilyAfter} »");
        if (!execute) { Done++; return; }

        string backup = Backup(item, row.Contact);
        Set(item, "FirstName", row.GivenAfter);
        Set(item, "MiddleName", row.MiddleAfter);
        Set(item, "LastName", row.FamilyAfter);
        if (newFileAs != fileAs) Set(item, "FileAs", newFileAs);
        Call(item, "Save");
        Done++;
        Journal(row.Action.ToString(), row.Contact, row.Id, row.StoreId, row.Id, "",
            [first, middle, last, fileAs], [row.GivenAfter, row.MiddleAfter, row.FamilyAfter, newFileAs], backup, "OK");
        Release(item);
    }

    private void MoveDuplicate(PlanRow row)
    {
        object item = session.GetItem(row.Id, row.StoreId);
        object parent = Get(item, "Parent")!;
        object? keep = row.KeepId.Length > 0 ? session.GetItem(row.KeepId, row.StoreId) : null;
        List<string> noRoom = [];
        List<(string Property, string Value)> merge = keep == null ? [] : MergeableFields(item, keep, out noRoom);
        if (noRoom.Count > 0)
        {
            Skipped++;
            log($"  ignoré doublon « {row.Contact} » : plus de case libre dans le contact conservé pour {string.Join(", ", noRoom)} — à fusionner à la main.");
            return;
        }
        log($"  {(execute ? "" : "[simulation] ")}doublon « {row.Contact} » → dossier « {DuplicatesFolderName} »"
            + (merge.Count > 0 ? $" ; ajouté au contact conservé : {string.Join(", ", merge.Select(m => m.Value))}" : ""));
        if (!execute) { Done++; return; }

        string backup = Backup(item, row.Contact);
        if (keep != null && merge.Count > 0)
        {
            backup += "|" + Backup(keep, row.Contact + " (conservé)");
            foreach ((string property, string value) in merge) Set(keep, property, value);
            Call(keep, "Save");
        }
        object dest = GetOrCreateSubFolder(parent, DuplicatesFolderName);
        object moved = Call(item, "Move", dest)!;
        Done++;
        Journal(row.Action.ToString(), row.Contact, row.Id, row.StoreId, Str(moved, "EntryID"), Str(parent, "EntryID"), null, null, backup, "OK",
            keep == null ? "" : row.KeepId, string.Join("|", merge.Select(m => $"{m.Property}={m.Value}")));
        Release(moved);
        Release(item);
        Release(keep);
    }

    private static readonly string[] SparePhoneProperties =
        ["OtherTelephoneNumber", "Business2TelephoneNumber", "Home2TelephoneNumber", "CarTelephoneNumber", "CallbackTelephoneNumber", "RadioTelephoneNumber"];

    /// <summary>
    /// Emails/phones of the duplicate that the kept contact does not have, and the empty field of the kept
    /// contact each one goes to (same kind of phone when free, otherwise a spare phone field).
    /// </summary>
    private static List<(string Property, string Value)> MergeableFields(object duplicate, object keep, out List<string> noRoom)
    {
        var result = new List<(string, string)>();
        noRoom = [];
        var used = new HashSet<string>();

        var keepEmails = Enumerable.Range(1, 3).Select(n => Normalizer.Email(Str(keep, $"Email{n}Address"))).Where(e => e.Length > 0).ToHashSet();
        for (int n = 1; n <= 3; n++)
        {
            string email = Str(duplicate, $"Email{n}Address");
            if (Normalizer.Email(email) is not { Length: > 0 } key || !keepEmails.Add(key)) continue;
            string? free = Enumerable.Range(1, 3).Select(i => $"Email{i}Address").FirstOrDefault(p => Str(keep, p).Length == 0 && !used.Contains(p));
            if (free == null) { noRoom.Add(email); continue; }
            used.Add(free);
            result.Add((free, email));
        }

        var keepPhones = PhoneProperties.Select(p => Normalizer.Phone(Str(keep, p))).Where(p => p.Length > 0).ToHashSet();
        foreach (string property in PhoneProperties)
        {
            string phone = Str(duplicate, property);
            if (Normalizer.Phone(phone) is not { Length: > 0 } key || !keepPhones.Add(key)) continue;
            string? free = new[] { property }.Concat(SparePhoneProperties).FirstOrDefault(p => Str(keep, p).Length == 0 && !used.Contains(p));
            if (free == null) { noRoom.Add(phone); continue; }
            used.Add(free);
            result.Add((free, phone));
        }
        return result;
    }

    private void AddFromReference(PlanRow row, object targetFolder)
    {
        object refItem = session.GetItem(row.RefId, row.RefStoreId);
        log($"  {(execute ? "" : "[simulation] ")}ajout de « {row.Contact} » depuis la référence");
        if (!execute) { Done++; return; }

        object copy = Call(refItem, "Copy")!;           // copy is created next to the original (in the PST)...
        object moved = Call(copy, "Move", targetFolder)!; // ...and moved to the mailbox, so the PST ends unchanged.
        Done++;
        Journal(row.Action.ToString(), row.Contact, row.RefId, Str(Get(targetFolder, "Store")!, "StoreID"), Str(moved, "EntryID"), "", null, null, "", "OK");
        Release(moved);
        Release(refItem);
    }

    /// <summary>Reverts what a journal recorded, newest first.</summary>
    public void Undo(string journalPath)
    {
        List<string[]> rows = Csv.Parse(Csv.ReadAllTextSmart(journalPath));
        if (rows.Count < 2) { log("Journal vide."); return; }
        var col = rows[0].Select((h, i) => (h, i)).ToDictionary(x => x.h, x => x.i);
        string get(string[] r, string name) => col.TryGetValue(name, out int i) && i < r.Length ? r[i] : "";

        foreach (string[] r in rows.Skip(1).Reverse())
        {
            if (get(r, "Resultat") != "OK") continue;
            string action = get(r, "Action"), contact = get(r, "Contact"), id = get(r, "Id_apres"), storeId = get(r, "StoreId");
            try
            {
                object item = session.GetItem(id, storeId);
                log($"  {(execute ? "" : "[simulation] ")}annulation {action} « {contact} »");
                if (!execute) { Done++; continue; }
                switch (Enum.Parse<PlanAction>(action))
                {
                    case PlanAction.SwapName:
                    case PlanAction.RenameFromReference:
                        if (Normalizer.Text(Str(item, "FirstName")) != Normalizer.Text(get(r, "Prenom_apres"))
                            || Normalizer.Text(Str(item, "LastName")) != Normalizer.Text(get(r, "Nom_apres")))
                        {
                            Skipped++;
                            log($"    ignoré : le contact a été modifié depuis.");
                            continue;
                        }
                        Set(item, "FirstName", get(r, "Prenom_avant"));
                        Set(item, "MiddleName", get(r, "DeuxiemePrenom_avant"));
                        Set(item, "LastName", get(r, "Nom_avant"));
                        Set(item, "FileAs", get(r, "ClasserSous_avant"));
                        Call(item, "Save");
                        Journal(action + "_annule", contact, id, storeId, id, "", null, null, "", "OK");
                        break;
                    case PlanAction.MoveDuplicate:
                        object origin = session.GetFolder(get(r, "DossierOrigine_Id"), storeId);
                        object back = Call(item, "Move", origin)!;
                        if (get(r, "Conserve_Id") is { Length: > 0 } keepId && get(r, "Champs_ajoutes") is { Length: > 0 } merged)
                        {
                            object keep = session.GetItem(keepId, storeId);
                            foreach (string pair in merged.Split('|'))
                            {
                                int eq = pair.IndexOf('=');
                                if (eq > 0 && Str(keep, pair[..eq]) == pair[(eq + 1)..]) Set(keep, pair[..eq], "");
                            }
                            Call(keep, "Save");
                            Release(keep);
                        }
                        Journal(action + "_annule", contact, id, storeId, Str(back, "EntryID"), "", null, null, "", "OK");
                        break;
                    case PlanAction.AddFromReference:
                        Call(item, "Delete"); // goes to "Éléments supprimés", still recoverable
                        Journal(action + "_annule", contact, id, storeId, "", "", null, null, "", "OK (dans Éléments supprimés)");
                        break;
                }
                Done++;
                Release(item);
            }
            catch (Exception ex)
            {
                Failed++;
                log($"  ERREUR annulation {action} « {contact} » : {ex.Message}");
            }
        }
        Finish("journal_annulation");
    }

    private string Backup(object item, string contact)
    {
        backupDir ??= Directory.CreateDirectory(Path.Combine(workDir, $"sauvegarde_{DateTime.Now:yyyyMMdd_HHmmss}")).FullName;
        string safe = string.Concat(contact.Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch));
        if (safe.Length > 60) safe = safe[..60];
        string path = Path.Combine(backupDir, $"{journal.Count:D4}_{safe}.msg");
        Call(item, "SaveAs", path, OlMsgUnicode);
        return path;
    }

    private void Journal(string action, string contact, string idBefore, string storeId, string idAfter, string originFolderId,
        string[]? before, string[]? after, string backup, string result, string keepId = "", string mergedFields = "")
    {
        before ??= ["", "", "", ""];
        after ??= ["", "", "", ""];
        journal.Add([DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), action, contact, idBefore, storeId, idAfter, originFolderId,
            .. before, .. after, backup, result, keepId, mergedFields]);
    }

    private void Finish(string name)
    {
        if (!execute)
        {
            log($"SIMULATION : {Done} action(s) seraient faites, {Skipped} ignorée(s). Rien n'a été modifié. Ajoutez --execute pour appliquer.");
            return;
        }
        backupDir ??= Directory.CreateDirectory(Path.Combine(workDir, $"sauvegarde_{DateTime.Now:yyyyMMdd_HHmmss}")).FullName;
        JournalPath = Path.Combine(backupDir, name + ".csv");
        Csv.WriteFile(JournalPath, journal);
        log($"Terminé : {Done} fait(s), {Skipped} ignoré(s), {Failed} erreur(s). Journal : {JournalPath}");
    }
}

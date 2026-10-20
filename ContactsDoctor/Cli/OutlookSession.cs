using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ContactsDoctor;

/// <summary>
/// Late-bound access to Outlook Classic (COM). Late binding means no Office interop assembly is needed
/// and any Outlook version (2013 → Microsoft 365) works. Exchange mailboxes and PST files are both "stores".
/// </summary>
public sealed partial class OutlookSession : IDisposable
{
    private const int OlFolderContacts = 10;
    private const int OlFolderDeletedItems = 3;
    private const int OlContactItemType = 2; // Folder.DefaultItemType for contact folders
    private const int OlContactClass = 40;   // ContactItem.Class

    /// <summary>Contact folders that are not the user's address book (caches, suggestions, Skype...).</summary>
    private static readonly HashSet<string> SkippedFolders = new(StringComparer.OrdinalIgnoreCase)
    {
        "Recipient Cache", "PersonMetadata", "GAL Contacts", "Organizational Contacts", "Companies", "ExternalContacts",
        "Lync Contacts", "MeContact", "Suggested Contacts", "Contacts suggérés", "Quick Contacts", "Contacts rapides",
        "Skype for Business Contacts", "Contacts Skype Entreprise", "Doublons (ContactsDoctor)",
    };

    public static readonly string[] PhoneProperties =
    [
        "MobileTelephoneNumber", "BusinessTelephoneNumber", "Business2TelephoneNumber", "HomeTelephoneNumber", "Home2TelephoneNumber",
        "PrimaryTelephoneNumber", "OtherTelephoneNumber", "CompanyMainTelephoneNumber", "CarTelephoneNumber", "AssistantTelephoneNumber",
        "CallbackTelephoneNumber", "RadioTelephoneNumber", "ISDNNumber", "PagerNumber", "TelexNumber", "TTYTDDTelephoneNumber",
        "BusinessFaxNumber", "HomeFaxNumber", "OtherFaxNumber",
    ];

    public const string DuplicatesFolderName = "Doublons (ContactsDoctor)";

    private readonly object app;
    public object Namespace { get; }
    private readonly List<object> storesWeOpened = [];

    public OutlookSession()
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("L'accès à Outlook n'est possible que sous Windows avec Outlook Classic installé.");
        Type type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook Classic n'est pas installé (le « nouvel Outlook » n'est pas pris en charge).");
        app = Activator.CreateInstance(type)!;
        Namespace = Call(app, "GetNamespace", "MAPI")!;
        Call(Namespace, "Logon", Type.Missing, Type.Missing, false, false);
    }

    // ---- small reflection helpers for IDispatch ----
    public static object? Get(object o, string name) => o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, null);
    public static void Set(object o, string name, object? value) => o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, [value]);
    public static object? Call(object o, string name, params object?[] args) => o.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, o, args);
    public static string Str(object o, string name)
    {
        try { return Get(o, name)?.ToString()?.Trim() ?? ""; }
        catch (Exception) { return ""; }
    }
    public static void Release(object? o)
    {
        if (o != null && OperatingSystem.IsWindows() && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
    }

    public IEnumerable<object> Stores()
    {
        object stores = Get(Namespace, "Stores")!;
        int count = (int)Get(stores, "Count")!;
        for (int i = 1; i <= count; i++) yield return Call(stores, "Item", i)!;
    }

    /// <summary>"" or "outlook" = default mailbox (Exchange); a path to a .pst; otherwise a store display name.</summary>
    public object OpenStore(string spec)
    {
        if (spec.Length == 0 || spec.Equals("outlook", StringComparison.OrdinalIgnoreCase))
            return Get(Namespace, "DefaultStore")!;

        if (spec.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
        {
            string full = Path.GetFullPath(spec);
            if (!File.Exists(full)) throw new FileNotFoundException("Fichier PST introuvable", full);
            if ((File.GetAttributes(full) & FileAttributes.ReadOnly) != 0)
                throw new InvalidOperationException($"Le fichier {full} est en lecture seule : Outlook ne peut pas l'ouvrir. Retirez l'attribut « Lecture seule » (travaillez sur une copie).");
            object? existing = FindStoreByPath(full);
            if (existing != null) return existing;
            Call(Namespace, "AddStore", full);
            object store = FindStoreByPath(full) ?? throw new InvalidOperationException($"Outlook n'a pas pu ouvrir {full}");
            storesWeOpened.Add(store);
            return store;
        }

        return Stores().FirstOrDefault(s => Str(s, "DisplayName").Equals(spec, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Aucune boîte / aucun fichier nommé « {spec} » dans Outlook. Utilisez la commande « stores ».");
    }

    private object? FindStoreByPath(string fullPath) =>
        Stores().FirstOrDefault(s => Str(s, "FilePath").Equals(fullPath, StringComparison.OrdinalIgnoreCase));

    /// <summary>All contact folders of a store (the default "Contacts" and any other contact folder), except caches.</summary>
    public List<(object Folder, string Path)> ContactFolders(object store)
    {
        var result = new List<(object, string)>();
        string deletedId = "";
        try { deletedId = Str(Call(store, "GetDefaultFolder", OlFolderDeletedItems)!, "EntryID"); }
        catch (Exception) { /* PST without default folders */ }

        void Walk(object folder, string path)
        {
            object folders = Get(folder, "Folders")!;
            int count = (int)Get(folders, "Count")!;
            for (int i = 1; i <= count; i++)
            {
                object sub = Call(folders, "Item", i)!;
                string name = Str(sub, "Name");
                string subPath = path.Length == 0 ? name : path + "\\" + name;
                if (Str(sub, "EntryID") == deletedId || SkippedFolders.Contains(name) || name.StartsWith('{')) continue;
                int itemType = -1;
                try { itemType = (int)Get(sub, "DefaultItemType")!; } catch (Exception) { }
                if (itemType == OlContactItemType) result.Add((sub, subPath));
                Walk(sub, subPath);
            }
        }
        Walk(Call(store, "GetRootFolder")!, "");
        return result;
    }

    public object DefaultContactsFolder(object store) => Call(store, "GetDefaultFolder", OlFolderContacts)!;

    public List<Contact> ReadContacts(object store, string source, Action<string>? log = null)
    {
        var contacts = new List<Contact>();
        string storeId = Str(store, "StoreID");
        foreach ((object folder, string path) in ContactFolders(store))
        {
            object items = Get(folder, "Items")!;
            int count = (int)Get(items, "Count")!;
            log?.Invoke($"  {path} : {count} élément(s)");
            for (int i = 1; i <= count; i++)
            {
                object item = Call(items, "Item", i)!;
                try
                {
                    if (Get(item, "Class") is int cls && cls == OlContactClass)
                    {
                        Contact c = FromItem(item, source);
                        c.Folder = path;
                        c.StoreId = storeId;
                        contacts.Add(c);
                    }
                }
                finally { Release(item); }
            }
            Release(items);
        }
        return contacts;
    }

    public static Contact FromItem(object item, string source)
    {
        var c = new Contact
        {
            Source = source,
            Id = Str(item, "EntryID"),
            GivenName = Str(item, "FirstName"),
            MiddleName = Str(item, "MiddleName"),
            FamilyName = Str(item, "LastName"),
            FullName = Str(item, "FullName"),
            FileAs = Str(item, "FileAs"),
            Company = Str(item, "CompanyName"),
            JobTitle = Str(item, "JobTitle"),
        };
        for (int n = 1; n <= 3; n++)
        {
            string address = Str(item, $"Email{n}Address");
            if (Str(item, $"Email{n}AddressType").Equals("EX", StringComparison.OrdinalIgnoreCase) || !address.Contains('@'))
            {
                // Exchange internal address (/o=ExchangeLabs/...): the SMTP address is usually in the display name.
                Match m = SmtpInText().Match(Str(item, $"Email{n}DisplayName"));
                address = m.Success ? m.Value : "";
            }
            if (address.Length > 0) c.Emails.Add(address);
        }
        foreach (string p in PhoneProperties)
            if (Str(item, p) is { Length: > 0 } phone) c.Phones.Add(phone);
        try { if (Get(item, "LastModificationTime") is DateTime dt) c.LastModified = dt; } catch (Exception) { }
        return c;
    }

    [GeneratedRegex(@"[\w.+'-]+@[\w-]+(\.[\w-]+)+")]
    private static partial Regex SmtpInText();

    public object GetItem(string entryId, string storeId) =>
        storeId.Length > 0 ? Call(Namespace, "GetItemFromID", entryId, storeId)! : Call(Namespace, "GetItemFromID", entryId)!;

    public object GetFolder(string entryId, string storeId) =>
        storeId.Length > 0 ? Call(Namespace, "GetFolderFromID", entryId, storeId)! : Call(Namespace, "GetFolderFromID", entryId)!;

    /// <summary>Sub-folder of <paramref name="parent"/> created on first use (never deleted by the tool).</summary>
    public static object GetOrCreateSubFolder(object parent, string name)
    {
        object folders = Get(parent, "Folders")!;
        int count = (int)Get(folders, "Count")!;
        for (int i = 1; i <= count; i++)
        {
            object f = Call(folders, "Item", i)!;
            if (Str(f, "Name").Equals(name, StringComparison.OrdinalIgnoreCase)) return f;
        }
        return Call(folders, "Add", name, OlFolderContacts)!;
    }

    public void Dispose()
    {
        foreach (object store in storesWeOpened)
        {
            try { Call(Namespace, "RemoveStore", Call(store, "GetRootFolder")); }
            catch (Exception) { /* Outlook may already have closed it */ }
        }
        storesWeOpened.Clear();
    }
}

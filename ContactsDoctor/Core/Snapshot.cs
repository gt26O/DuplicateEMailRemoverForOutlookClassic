using System.Text.Encodings.Web;
using System.Text.Json;

namespace ContactsDoctor;

/// <summary>Frozen copy of a source's contacts (JSON) so an analysis can be redone later, or on another PC.</summary>
public static class Snapshot
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static void Save(string path, IEnumerable<Contact> contacts) =>
        File.WriteAllText(path, JsonSerializer.Serialize(contacts.ToList(), Options));

    public static List<Contact> Load(string path, string source)
    {
        List<Contact> contacts = JsonSerializer.Deserialize<List<Contact>>(File.ReadAllText(path), Options) ?? [];
        foreach (Contact c in contacts) c.Source = source;
        return contacts;
    }
}

using System.Text;

namespace ContactsDoctor;

/// <summary>
/// Reads and writes vCard files (.vcf). Handles what the Samsung "Contacts" app exports:
/// vCard 2.1 with QUOTED-PRINTABLE / CHARSET=UTF-8, folded BASE64 photos, item1.EMAIL groups,
/// as well as vCard 3.0/4.0 (Google, iCloud, Outlook).
/// </summary>
public static class VCard
{
    public static List<Contact> ReadFile(string path, string source) => Read(File.ReadAllText(path, Encoding.UTF8), source);

    public static List<Contact> Read(string text, string source)
    {
        var contacts = new List<Contact>();
        List<string>? current = null;
        foreach (string line in Unfold(text))
        {
            if (line.Trim().Equals("BEGIN:VCARD", StringComparison.OrdinalIgnoreCase))
            {
                current = [];
                continue;
            }
            if (line.Trim().Equals("END:VCARD", StringComparison.OrdinalIgnoreCase))
            {
                if (current != null)
                    contacts.Add(FromLines(current, source, (contacts.Count + 1).ToString("D5")));
                current = null;
                continue;
            }
            current?.Add(line);
        }
        return contacts;
    }

    /// <summary>Joins folded lines (leading space/tab) and quoted-printable soft line breaks (trailing '=').</summary>
    public static IEnumerable<string> Unfold(string text)
    {
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        StringBuilder? pending = null;
        bool pendingIsQpContinuation = false;
        foreach (string raw in lines)
        {
            if (pending != null && pendingIsQpContinuation)
            {
                pending.Length--; // drop the soft-break '='
                pending.Append(raw);
            }
            else if (pending != null && raw.Length > 0 && (raw[0] == ' ' || raw[0] == '\t'))
            {
                pending.Append(raw, 1, raw.Length - 1);
            }
            else
            {
                if (pending != null) yield return pending.ToString();
                pending = new StringBuilder(raw);
            }
            string soFar = pending.ToString();
            pendingIsQpContinuation = soFar.EndsWith('=') && IsQuotedPrintable(soFar);
        }
        if (pending != null && pending.Length > 0) yield return pending.ToString();
    }

    private static bool IsQuotedPrintable(string line)
    {
        int colon = line.IndexOf(':');
        if (colon < 0) return false;
        string head = line[..colon].ToUpperInvariant();
        return head.Contains("QUOTED-PRINTABLE") || head.Contains("ENCODING=QP");
    }

    private sealed record Property(string Name, string[] Params, string Value);

    private static Property? Parse(string line)
    {
        int colon = -1;
        bool inQuotes = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] == '"') inQuotes = !inQuotes;
            else if (line[i] == ':' && !inQuotes) { colon = i; break; }
        }
        if (colon <= 0) return null;
        string[] head = line[..colon].Split(';');
        string name = head[0];
        int dot = name.LastIndexOf('.');
        if (dot >= 0) name = name[(dot + 1)..]; // "item1.EMAIL" -> "EMAIL"
        string[] parameters = head.Skip(1).ToArray();
        string value = line[(colon + 1)..];

        string upperParams = string.Join(";", parameters).ToUpperInvariant();
        if (upperParams.Contains("QUOTED-PRINTABLE") || upperParams.Contains("ENCODING=QP"))
        {
            string charset = parameters.Select(p => p.Split('=')).Where(kv => kv.Length == 2 && kv[0].Equals("CHARSET", StringComparison.OrdinalIgnoreCase))
                .Select(kv => kv[1].Trim('"')).FirstOrDefault() ?? "UTF-8";
            value = DecodeQuotedPrintable(value, charset);
        }
        return new Property(name.ToUpperInvariant(), parameters, value);
    }

    public static string DecodeQuotedPrintable(string value, string charset)
    {
        var bytes = new List<byte>(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            if (value[i] == '=' && i + 2 < value.Length && Uri.IsHexDigit(value[i + 1]) && Uri.IsHexDigit(value[i + 2]))
            {
                bytes.Add(Convert.ToByte(value.Substring(i + 1, 2), 16));
                i += 2;
            }
            else
            {
                bytes.AddRange(Encoding.UTF8.GetBytes(value[i].ToString()));
            }
        }
        Encoding encoding;
        try { encoding = Encoding.GetEncoding(charset); }
        catch (ArgumentException) { encoding = Encoding.UTF8; }
        return encoding.GetString(bytes.ToArray());
    }

    /// <summary>Splits on ';' (or another separator) that is not escaped with a backslash.</summary>
    public static string[] SplitStructured(string value, char separator = ';')
    {
        var parts = new List<string>();
        var sb = new StringBuilder();
        for (int i = 0; i < value.Length; i++)
        {
            if (value[i] == '\\' && i + 1 < value.Length)
            {
                sb.Append(value[i]).Append(value[i + 1]);
                i++;
            }
            else if (value[i] == separator)
            {
                parts.Add(sb.ToString());
                sb.Clear();
            }
            else sb.Append(value[i]);
        }
        parts.Add(sb.ToString());
        return parts.ToArray();
    }

    public static string Unescape(string value)
    {
        var sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            if (value[i] == '\\' && i + 1 < value.Length)
            {
                char next = value[++i];
                sb.Append(next is 'n' or 'N' ? '\n' : next);
            }
            else sb.Append(value[i]);
        }
        return sb.ToString().Trim();
    }

    public static string Escape(string value) =>
        value.Replace("\\", "\\\\").Replace(";", "\\;").Replace(",", "\\,").Replace("\r\n", "\\n").Replace("\n", "\\n");

    private static Contact FromLines(List<string> lines, string source, string id)
    {
        var c = new Contact { Source = source, Id = id, RawVCard = lines };
        foreach (string line in lines)
        {
            Property? p = Parse(line);
            if (p == null) continue;
            switch (p.Name)
            {
                case "N":
                    string[] n = SplitStructured(p.Value);
                    c.FamilyName = Unescape(n.ElementAtOrDefault(0) ?? "");
                    c.GivenName = Unescape(n.ElementAtOrDefault(1) ?? "");
                    c.MiddleName = Unescape(n.ElementAtOrDefault(2) ?? "");
                    break;
                case "FN":
                    c.FullName = Unescape(p.Value);
                    break;
                case "ORG":
                    c.Company = Unescape(SplitStructured(p.Value)[0]);
                    break;
                case "TITLE":
                    c.JobTitle = Unescape(p.Value);
                    break;
                case "EMAIL":
                    if (Unescape(p.Value) is { Length: > 0 } email) c.Emails.Add(email);
                    break;
                case "TEL":
                    string tel = Unescape(p.Value);
                    if (tel.StartsWith("tel:", StringComparison.OrdinalIgnoreCase)) tel = tel[4..];
                    if (tel.Length > 0) c.Phones.Add(tel);
                    break;
                case "REV":
                    if (DateTime.TryParse(p.Value, out DateTime rev)) c.LastModified = rev;
                    break;
            }
        }
        return c;
    }

    /// <summary>Writes contacts to vCard. Contacts read from a .vcf keep their version and all their original
    /// lines (photo, address, birthday, phone types...); only N and FN are rewritten, and emails/phones
    /// that are not already in the card (e.g. merged from a duplicate) are appended.</summary>
    public static string Write(IEnumerable<Contact> contacts)
    {
        var sb = new StringBuilder();
        foreach (Contact c in contacts)
        {
            string version = "3.0";
            var emailsInCard = new HashSet<string>();
            var phonesInCard = new HashSet<string>();
            var keptLines = new List<string>();
            if (c.RawVCard != null)
            {
                foreach (string line in c.RawVCard)
                {
                    Property? p = Parse(line);
                    if (p == null) continue;
                    if (p.Name == "VERSION") { version = p.Value.Trim(); continue; }
                    if (p.Name is "N" or "FN") continue;
                    if (p.Name == "EMAIL") emailsInCard.Add(Normalizer.Email(Unescape(p.Value)));
                    if (p.Name == "TEL") phonesInCard.Add(Normalizer.Phone(Unescape(p.Value)));
                    keptLines.Add(line);
                }
            }
            else
            {
                if (c.Company.Length > 0) keptLines.Add($"ORG:{Escape(c.Company)}");
                if (c.JobTitle.Length > 0) keptLines.Add($"TITLE:{Escape(c.JobTitle)}");
            }

            bool v21 = version == "2.1";
            string charset = v21 ? ";CHARSET=UTF-8" : "";
            sb.Append("BEGIN:VCARD\r\n");
            sb.Append($"VERSION:{version}\r\n");
            Func<string, string> esc = v21 ? (v => v.Replace(";", "\\;").Replace("\r\n", " ").Replace("\n", " ")) : Escape;
            sb.Append($"N{charset}:{esc(c.FamilyName)};{esc(c.GivenName)};{esc(c.MiddleName)};;\r\n");
            sb.Append($"FN{charset}:{esc(c.DisplayName)}\r\n");
            foreach (string line in keptLines) sb.Append(line).Append("\r\n");
            foreach (string e in c.Emails)
            {
                if (Normalizer.Email(e) is { Length: > 0 } key && emailsInCard.Add(key))
                    sb.Append(v21 ? $"EMAIL;INTERNET:{e}\r\n" : $"EMAIL;TYPE=INTERNET:{Escape(e)}\r\n");
            }
            foreach (string t in c.Phones)
            {
                string key = Normalizer.Phone(t);
                if (key.Length > 0 && phonesInCard.Add(key))
                    sb.Append($"TEL:{(v21 ? t : Escape(t))}\r\n");
            }
            sb.Append("END:VCARD\r\n");
        }
        return sb.ToString();
    }
}

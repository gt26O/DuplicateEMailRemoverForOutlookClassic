using System.Globalization;
using System.Text;

namespace ContactsDoctor;

/// <summary>Normalisation used to compare contacts coming from different systems.</summary>
public static class Normalizer
{
    /// <summary>Lower case, no accents, only letters/digits separated by single spaces.</summary>
    public static string Text(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        string decomposed = s.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(decomposed.Length);
        bool lastWasSpace = true;
        foreach (char c in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) continue;
            if (char.IsLetterOrDigit(c))
            {
                sb.Append(char.ToLowerInvariant(c));
                lastWasSpace = false;
            }
            else if (!lastWasSpace)
            {
                sb.Append(' ');
                lastWasSpace = true;
            }
        }
        return sb.ToString().Trim();
    }

    /// <summary>Trimmed lower-case email, or "" when it does not look like an SMTP address.</summary>
    public static string Email(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        string e = s.Trim().Trim('<', '>', '"', '\'').ToLowerInvariant();
        if (e.StartsWith("mailto:")) e = e[7..];
        int at = e.IndexOf('@');
        return at > 0 && at < e.Length - 1 && !e.Contains(' ') ? e : "";
    }

    /// <summary>
    /// Phone key: the last 9 digits (enough to ignore +41 / 0041 / 0 prefixes for CH and FR numbers).
    /// Shorter numbers are kept whole; numbers with fewer than 6 digits are ignored.
    /// </summary>
    public static string Phone(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var digits = new StringBuilder();
        foreach (char c in s)
        {
            if (c >= '0' && c <= '9') digits.Append(c);
            else if (char.IsLetter(c) && digits.Length > 0) break; // "079 123 45 67 ext 12" -> stop at the extension
        }
        string d = digits.ToString();
        if (d.Length < 6) return "";
        return d.Length > 9 ? d[^9..] : d;
    }

    /// <summary>Name tokens of a contact, from the structured fields, else the full name / file-as.</summary>
    public static string[] NameTokens(Contact c)
    {
        string source = string.Join(" ", c.GivenName, c.MiddleName, c.FamilyName);
        if (Text(source).Length == 0) source = c.FullName;
        if (Text(source).Length == 0) source = c.FileAs;
        return Text(source).Split(' ', StringSplitOptions.RemoveEmptyEntries);
    }

    /// <summary>Order-insensitive name key: "Dupont Jean" and "Jean Dupont" give the same key.</summary>
    public static string NameKey(Contact c) => string.Join(" ", NameTokens(c).OrderBy(t => t, StringComparer.Ordinal));

    /// <summary>Order-sensitive name: used to see that two contacts have their names in a different order.</summary>
    public static string OrderedName(Contact c) => string.Join(" ", NameTokens(c));

    public static IEnumerable<string> EmailKeys(Contact c) => c.Emails.Select(Email).Where(e => e.Length > 0).Distinct();

    public static IEnumerable<string> PhoneKeys(Contact c) => c.Phones.Select(Phone).Where(p => p.Length > 0).Distinct();
}

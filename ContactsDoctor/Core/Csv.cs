using System.Text;

namespace ContactsDoctor;

/// <summary>Minimal RFC 4180 CSV reader/writer (quotes, embedded separators and line breaks).</summary>
public static class Csv
{
    /// <summary>
    /// Reads a text file written as UTF-8 (with or without BOM) or, when it is not valid UTF-8, as Windows-1252:
    /// that is what a French Excel produces with "Enregistrer sous > CSV (séparateur : point-virgule)".
    /// </summary>
    public static string ReadAllTextSmart(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        try
        {
            return new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes).TrimStart('\uFEFF');
        }
        catch (DecoderFallbackException)
        {
            return Latin1Windows.GetString(bytes);
        }
    }

    private static Encoding Latin1Windows
    {
        get
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                return Encoding.GetEncoding(1252);
            }
            catch (Exception) { return Encoding.Latin1; }
        }
    }

    public static char DetectSeparator(string headerLine) =>
        headerLine.Count(ch => ch == ';') > headerLine.Count(ch => ch == ',') ? ';' : ',';

    public static List<string[]> Parse(string text, char? separator = null)
    {
        if (text.Length > 0 && text[0] == '﻿') text = text[1..];
        int firstBreak = text.IndexOfAny(['\r', '\n']);
        char sep = separator ?? DetectSeparator(firstBreak < 0 ? text : text[..firstBreak]);

        var rows = new List<string[]>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool inQuotes = false;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (inQuotes)
            {
                if (c == '"' && i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                else if (c == '"') inQuotes = false;
                else field.Append(c);
            }
            else if (c == '"') inQuotes = true;
            else if (c == sep) { row.Add(field.ToString()); field.Clear(); }
            else if (c == '\r' || c == '\n')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                row.Add(field.ToString()); field.Clear();
                if (row.Count > 1 || row[0].Length > 0) rows.Add(row.ToArray());
                row = [];
            }
            else field.Append(c);
        }
        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            rows.Add(row.ToArray());
        }
        return rows;
    }

    public static string Field(string? value, char sep = ';')
    {
        value ??= "";
        return value.IndexOfAny([sep, '"', '\r', '\n']) >= 0 ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
    }

    public static string Line(IEnumerable<string?> values, char sep = ';') => string.Join(sep, values.Select(v => Field(v, sep)));

    /// <summary>Writes UTF-8 with BOM and ';' so it opens correctly in a French Excel.</summary>
    public static void WriteFile(string path, IEnumerable<string[]> rows)
    {
        var sb = new StringBuilder();
        foreach (string[] r in rows) sb.Append(Line(r)).Append("\r\n");
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
    }
}

/// <summary>
/// Reads contact CSV exports: Outlook ("Importer/Exporter" > CSV, English or French headers),
/// Google Contacts and similar. Columns are recognised by name.
/// </summary>
public static class ContactCsv
{
    private enum Column { Ignore, Given, Middle, Family, FullName, FileAs, Company, JobTitle, Email, Phone }

    private static Column Classify(string header)
    {
        string h = Normalizer.Text(header);
        if (h.Length == 0) return Column.Ignore;
        bool has(string s) => h.Contains(s);
        if (has("type") || has("display") || has("affich") || has("label") || has("libelle")) return Column.Ignore;
        if (has("middle") || has("deuxieme prenom") || has("second prenom") || has("additional name")) return Column.Middle;
        if (has("first name") || has("given name") || h == "prenom") return Column.Given;
        if (has("last name") || has("family name") || has("surname") || h == "nom" || h == "nom de famille") return Column.Family;
        if (h is "name" or "nom complet" or "full name") return Column.FullName;
        if (has("file as") || has("classer sous")) return Column.FileAs;
        if (h is "company" or "societe" or "organization 1 name" or "organization name" or "entreprise") return Column.Company;
        if (h is "job title" or "fonction" or "organization 1 title" or "poste") return Column.JobTitle;
        if ((has("mail") || has("messagerie")) && !has("mailing")) return Column.Email;
        if (has("phone") || has("telephone") || has("tel ") || h.StartsWith("tel") || has("mobile") || has("portable") || has("fax")
            || has("pager") || has("radio") || has("telex") || has("rnis") || has("isdn"))
            return Column.Phone;
        return Column.Ignore;
    }

    public static List<Contact> ReadFile(string path, string source) => Read(Csv.ReadAllTextSmart(path), source);

    public static List<Contact> Read(string text, string source)
    {
        List<string[]> rows = Csv.Parse(text);
        var contacts = new List<Contact>();
        if (rows.Count < 2) return contacts;
        Column[] columns = rows[0].Select(Classify).ToArray();
        for (int r = 1; r < rows.Count; r++)
        {
            var c = new Contact { Source = source, Id = r.ToString("D5") };
            string[] row = rows[r];
            for (int i = 0; i < Math.Min(row.Length, columns.Length); i++)
            {
                string v = row[i].Trim();
                if (v.Length == 0) continue;
                switch (columns[i])
                {
                    case Column.Given: c.GivenName = v; break;
                    case Column.Middle: c.MiddleName = v; break;
                    case Column.Family: c.FamilyName = v; break;
                    case Column.FullName: c.FullName = v; break;
                    case Column.FileAs: c.FileAs = v; break;
                    case Column.Company: c.Company = v; break;
                    case Column.JobTitle: c.JobTitle = v; break;
                    // Google puts several values in one cell separated by " ::: "
                    case Column.Email: c.Emails.AddRange(v.Split(":::").Select(x => x.Trim()).Where(x => Normalizer.Email(x).Length > 0)); break;
                    case Column.Phone: c.Phones.AddRange(v.Split(":::").Select(x => x.Trim()).Where(x => x.Length > 0)); break;
                }
            }
            if (Normalizer.NameTokens(c).Length > 0 || c.Emails.Count > 0 || c.Phones.Count > 0 || c.Company.Length > 0)
                contacts.Add(c);
        }
        return contacts;
    }
}

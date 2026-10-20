using ContactsDoctor;
using Xunit;

public class NormalizerAndCsvTests
{
    [Theory]
    [InlineData("+41 79 123 45 67", "791234567")]
    [InlineData("0041791234567", "791234567")]
    [InlineData("079/123.45.67", "791234567")]
    [InlineData("+33 6 12 34 56 78", "612345678")]
    [InlineData("1234", "")]
    [InlineData("021 345 67 89 ext 12", "213456789")]
    public void PhoneKey(string input, string expected) => Assert.Equal(expected, Normalizer.Phone(input));

    [Fact]
    public void NameKeyIgnoresOrderCaseAndAccents()
    {
        var a = new Contact { GivenName = "Jérôme", FamilyName = "Müller" };
        var b = new Contact { FullName = "MULLER jerome" };
        Assert.Equal(Normalizer.NameKey(a), Normalizer.NameKey(b));
        Assert.NotEqual(Normalizer.OrderedName(a), Normalizer.OrderedName(b));
    }

    [Fact]
    public void ReadsFrenchOutlookCsv()
    {
        string csv = "﻿\"Prénom\",\"Deuxième prénom\",\"Nom\",\"Société\",\"Téléphone (bureau)\",\"Tél. mobile\",\"Adresse de messagerie\",\"Type de messagerie\",\"Nom complet de l'adresse de messagerie\"\r\n"
                   + "\"Jean\",\"\",\"Dupont\",\"ACME\",\"021 111 22 33\",\"079 444 55 66\",\"jean@acme.ch\",\"SMTP\",\"Jean Dupont (jean@acme.ch)\"\r\n"
                   + "\"Anne\",\"\",\"Martin\",\"\",\"\",\"\",\"\",\"\",\"\"\r\n";
        List<Contact> list = ContactCsv.Read(csv, "csv");
        Assert.Equal(2, list.Count);
        Assert.Equal("Jean", list[0].GivenName);
        Assert.Equal("Dupont", list[0].FamilyName);
        Assert.Equal("ACME", list[0].Company);
        Assert.Equal(["jean@acme.ch"], list[0].Emails);
        Assert.Equal(2, list[0].Phones.Count);
    }

    [Fact]
    public void ReadsGoogleCsvWithSemicolonAndMultiValues()
    {
        string csv = "Given Name;Family Name;E-mail 1 - Type;E-mail 1 - Value;Phone 1 - Type;Phone 1 - Value\n"
                   + "Paul;Durand;* Home;paul@x.ch ::: p.durand@y.ch;Mobile;\"079 000 11 22\"\n";
        Contact c = Assert.Single(ContactCsv.Read(csv, "g"));
        Assert.Equal(["paul@x.ch", "p.durand@y.ch"], c.Emails);
        Assert.Equal(["079 000 11 22"], c.Phones);
    }
}

public class EncodingTests
{
    [Fact]
    public void ReadsExcelAnsiAndUtf8Csv()
    {
        string text = "Prénom;Nom\r\nJérôme;Müller\r\n";
        string dir = Directory.CreateTempSubdirectory().FullName;
        try
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            File.WriteAllText(Path.Combine(dir, "ansi.csv"), text, System.Text.Encoding.GetEncoding(1252));
            File.WriteAllText(Path.Combine(dir, "utf8.csv"), text, new System.Text.UTF8Encoding(true));
            Assert.Equal(text, Csv.ReadAllTextSmart(Path.Combine(dir, "ansi.csv")));
            Assert.Equal(text, Csv.ReadAllTextSmart(Path.Combine(dir, "utf8.csv")));
            Assert.Equal("Jérôme", ContactCsv.ReadFile(Path.Combine(dir, "ansi.csv"), "x")[0].GivenName);
        }
        finally { Directory.Delete(dir, true); }
    }
}

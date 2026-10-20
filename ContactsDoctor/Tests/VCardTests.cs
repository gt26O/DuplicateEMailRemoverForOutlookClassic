using ContactsDoctor;
using Xunit;

public class VCardTests
{
    // Shape of a Samsung Contacts export (vCard 2.1, quoted-printable, folded photo, item groups).
    private const string Samsung = "BEGIN:VCARD\r\n" +
        "VERSION:2.1\r\n" +
        "N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:M=C3=BCller;J=C3=A9r=C3=B4me;;;\r\n" +
        "FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:J=C3=A9r=C3=B4me M=C3=BCller\r\n" +
        "TEL;CELL:+41 79 123 45 67\r\n" +
        "TEL;HOME:021 345 67 89\r\n" +
        "EMAIL;HOME:jerome@example.ch\r\n" +
        "item1.EMAIL;TYPE=INTERNET:j.muller@example.com\r\n" +
        "PHOTO;ENCODING=BASE64;JPEG:/9j/4AAQSkZJRgABAQAAAQABAAD\r\n" +
        " AAAAAAAAAAAAAAAA\r\n" +
        "\r\n" +
        "NOTE;ENCODING=QUOTED-PRINTABLE;CHARSET=UTF-8:Ligne =C3=A9crite sur =\r\n" +
        "deux lignes\r\n" +
        "END:VCARD\r\n" +
        "BEGIN:VCARD\r\n" +
        "VERSION:3.0\r\n" +
        "N:Dupont\\, fils;Jean;;;\r\n" +
        "FN:Jean Dupont\r\n" +
        "TEL;TYPE=CELL:0791112233\r\n" +
        "END:VCARD\r\n";

    [Fact]
    public void ReadsSamsungExport()
    {
        List<Contact> list = VCard.Read(Samsung, "tel");
        Assert.Equal(2, list.Count);
        Contact c = list[0];
        Assert.Equal("Jérôme", c.GivenName);
        Assert.Equal("Müller", c.FamilyName);
        Assert.Equal("Jérôme Müller", c.FullName);
        Assert.Equal(["+41 79 123 45 67", "021 345 67 89"], c.Phones);
        Assert.Equal(["jerome@example.ch", "j.muller@example.com"], c.Emails);
        Assert.Contains(c.RawVCard!, l => l.StartsWith("NOTE") && l.Contains("=C3=A9crite sur deux lignes"));
        Assert.Contains(c.RawVCard!, l => l.StartsWith("PHOTO") && l.EndsWith("AAAAAAAAAAAAAAAA"));
        Assert.Equal("Dupont, fils", list[1].FamilyName);
        Assert.Equal("Jean", list[1].GivenName);
    }

    [Fact]
    public void WriteKeepsOriginalLinesAndRoundTrips()
    {
        List<Contact> list = VCard.Read(Samsung, "tel");
        (list[0].GivenName, list[0].FamilyName) = (list[0].FamilyName, list[0].GivenName);
        list[0].Phones.Add("+41 79 999 88 77");
        list[0].Phones.Add("079 123 45 67"); // same number as the first one, must not be duplicated
        string written = VCard.Write(list);

        Assert.Contains("PHOTO;ENCODING=BASE64;JPEG:", written);
        Assert.Contains("TEL;CELL:+41 79 123 45 67", written);
        Assert.Contains("VERSION:2.1", written);

        List<Contact> again = VCard.Read(written, "tel");
        Assert.Equal(2, again.Count);
        Assert.Equal("Müller", again[0].GivenName);
        Assert.Equal("Jérôme", again[0].FamilyName);
        Assert.Equal(3, again[0].Phones.Count);
        Assert.Equal("Dupont, fils", again[1].FamilyName);
    }

    [Fact]
    public void WritesContactsWithoutRawLines()
    {
        var c = new Contact { GivenName = "Anne", FamilyName = "Martin; SA", Company = "ACME", Emails = ["anne@acme.ch"], Phones = ["044 111 22 33"] };
        List<Contact> back = VCard.Read(VCard.Write([c]), "x");
        Assert.Equal("Martin; SA", back[0].FamilyName);
        Assert.Equal("ACME", back[0].Company);
        Assert.Equal(["anne@acme.ch"], back[0].Emails);
    }
}

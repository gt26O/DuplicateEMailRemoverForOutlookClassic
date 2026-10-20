using ContactsDoctor;
using Xunit;

public class AnalyzerTests
{
    private static Contact C(string source, string id, string given, string family, string[]? emails = null, string[]? phones = null) => new()
    {
        Source = source, Id = id, GivenName = given, FamilyName = family, Emails = [.. emails ?? []], Phones = [.. phones ?? []],
    };

    /// <summary>A small version of the client case: PST before migration, Exchange and phone after.</summary>
    private static List<Contact> Scenario() =>
    [
        // reference (PST before the migration)
        C("pst", "p1", "Jean", "Dupont", ["jean@dupont.ch"], ["079 111 11 11"]),
        C("pst", "p2", "Anne", "Martin", ["anne@martin.ch"], ["079 222 22 22"]),
        C("pst", "p3", "Paul", "Durand", ["paul@durand.ch"]),
        C("pst", "p4", "Marc", "Rochat", phones: ["079 444 44 44", "021 444 44 44"]),
        C("pst", "p5", "Sophie", "Rochat", phones: ["079 555 55 55", "021 444 44 44"]),
        C("pst", "p6", "Luc", "Blanc", ["luc@acme.ch"], ["021 000 00 00"]),
        C("pst", "p7", "Eva", "Noir", ["eva@acme.ch"], ["021 000 00 00"]),
        C("pst", "p8", "Ivan", "Gris", ["ivan@acme.ch"], ["021 000 00 00"]),

        // Exchange after the migration
        C("exchange", "e1", "Dupont", "Jean", ["jean@dupont.ch"], ["+41791111111"]), // swapped
        C("exchange", "e2", "Anne", "Martin", ["anne@martin.ch"], ["079 222 22 22"]),
        C("exchange", "e2b", "Anne", "Martin", ["anne@martin.ch"], ["078 999 99 99"]), // duplicate with an extra number
        C("exchange", "e3", "Lucie", "Favre", ["lucie@favre.ch"]), // created after the migration
        C("exchange", "e4", "Marc", "Rochat", phones: ["079 444 44 44", "021 444 44 44"]),
        C("exchange", "e5", "Sophie", "Rochat", phones: ["079 555 55 55", "021 444 44 44"]),
        C("exchange", "e6", "Luc", "Blanc", ["luc@acme.ch"], ["021 000 00 00"]),
        C("exchange", "e7", "Eva", "Noir", ["eva@acme.ch"], ["021 000 00 00"]),
        C("exchange", "e8", "Ivan", "Gris", ["ivan@acme.ch"], ["021 000 00 00"]),
        // Paul Durand is missing

        // Phone (Samsung export) after the "merge contacts" tool
        C("tel", "t1", "Dupont", "Jean", ["jean@dupont.ch"], ["079 111 11 11"]), // same swap as Exchange
        C("tel", "t2", "Anne", "Martin", ["anne@martin.ch"], ["079 222 22 22"]),
        C("tel", "t4", "Marc", "Rochat", phones: ["079 444 44 44", "079 555 55 55"]), // Sophie merged into Marc
        C("tel", "t6", "Luc", "Blanc", ["luc@acme.ch"]),
    ];

    [Fact]
    public void DetectsEachKindOfProblem()
    {
        AnalysisResult a = Analyzer.Analyze("pst", Scenario());

        Issue swap = Assert.Single(a.For("exchange", IssueKind.NameSwapped));
        Assert.Equal("e1", swap.Record!.Id);
        Assert.Equal("p1", swap.Reference!.Id);
        Assert.Single(a.For("tel", IssueKind.NameSwapped));

        Issue dup = Assert.Single(a.For("exchange", IssueKind.Duplicate));
        Assert.Single(dup.Others);
        Assert.Empty(a.For("pst", IssueKind.Duplicate));

        Assert.Equal("p3", Assert.Single(a.For("exchange", IssueKind.Missing)).Reference!.Id);
        Assert.Equal("e3", Assert.Single(a.For("exchange", IssueKind.Extra)).Record!.Id);

        Issue merge = Assert.Single(a.For("tel", IssueKind.SuspectedMerge));
        Assert.Equal("t4", merge.Record!.Id);
        Assert.Equal("p5", Assert.Single(merge.Others).Id);
        Assert.Contains(a.For("tel", IssueKind.Missing), i => i.Reference!.Id == "p5");

        // The switchboard number is shared by 3 people: ignored, so Luc/Eva/Ivan stay separate.
        Assert.Contains("p:210000000", a.IgnoredSharedKeys.Select(k => "p:" + k));
        Assert.Contains(a.For("tel", IssueKind.DataMissing), i => i.Record!.Id == "t6");
        Assert.Empty(a.For("exchange", IssueKind.SuspectedMerge));
        Assert.DoesNotContain(a.Issues, i => i.Kind == IssueKind.Duplicate && i.Record!.FamilyName is "Blanc" or "Noir" or "Gris");
    }

    [Fact]
    public void ConclusionSaysExchangeAndPhoneShareTheSameSwaps()
    {
        AnalysisResult a = Analyzer.Analyze("pst", Scenario());
        List<string> lines = Report.Conclusions(a);
        Assert.Contains(lines, l => l.Contains("mêmes contacts qui sont inversés"));
        Assert.Contains(lines, l => l.Contains("Fusionner les contacts"));
    }

    [Fact]
    public void PlanSwapsAddsAndMergesDuplicates()
    {
        AnalysisResult a = Analyzer.Analyze("pst", Scenario());
        List<PlanRow> plan = Plan.Build(a, "exchange");

        PlanRow swap = Assert.Single(plan, r => r.Action == PlanAction.SwapName);
        Assert.True(swap.Apply);
        Assert.Equal(("Jean", "Dupont"), (swap.GivenAfter, swap.FamilyAfter));

        PlanRow dup = Assert.Single(plan, r => r.Action == PlanAction.MoveDuplicate);
        Assert.True(dup.Apply);
        Assert.Equal("e2", dup.KeepId);
        Assert.Contains("078 999 99 99", dup.Reason); // the copy's extra number is merged into the kept contact

        PlanRow add = Assert.Single(plan, r => r.Action == PlanAction.AddFromReference);
        Assert.Equal("p3", add.RefId);
    }

    [Fact]
    public void PlanSurvivesCsvRoundTripAndFixesAList()
    {
        List<Contact> all = Scenario();
        AnalysisResult a = Analyzer.Analyze("pst", all);
        List<PlanRow> plan = Plan.Build(a, "exchange");

        string path = Path.GetTempFileName();
        try
        {
            Plan.Save(path, plan);
            List<PlanRow> loaded = Plan.Load(path);
            Assert.Equal(plan.Count, loaded.Count);
            Assert.Equal(plan.Select(p => (p.Apply, p.Action, p.Id, p.GivenAfter)), loaded.Select(p => (p.Apply, p.Action, p.Id, p.GivenAfter)));

            List<Contact> exchange = all.Where(c => c.Source == "exchange").ToList();
            List<Contact> fixedList = Plan.ApplyToList(exchange, loaded, all.Where(c => c.Source == "pst").ToList());
            Assert.Contains(fixedList, c => c.Id == "e1" && c.GivenName == "Jean" && c.FamilyName == "Dupont");
            Assert.DoesNotContain(fixedList, c => c.Id == "e2b");
            Assert.Contains("078 999 99 99", fixedList.Single(c => c.Id == "e2").Phones); // merged, not lost
            Assert.Contains(fixedList, c => c.FamilyName == "Durand");

            AnalysisResult after = Analyzer.Analyze("pst", [.. all.Where(c => c.Source == "pst"), .. fixedList]);
            Assert.Empty(after.For("exchange", IssueKind.NameSwapped));
            Assert.Empty(after.For("exchange", IssueKind.Duplicate));
            Assert.Empty(after.For("exchange", IssueKind.Missing));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void HtmlReportIsEscaped()
    {
        List<Contact> all = [C("pst", "1", "<b>", "x"), C("tel", "2", "x", "<b>")];
        string html = Report.Html(Analyzer.Analyze("pst", all), new DateTime(2026, 10, 20));
        Assert.DoesNotContain("<b>", html.Replace("<b>pst</b>", ""));
        Assert.Contains("&lt;b&gt;", html);
    }

    [Theory]
    [InlineData("Jean, Dupont", "Dupont, Jean")] // "Nom, Prénom" built from the swapped fields
    [InlineData("Dupont Jean", "Jean Dupont")]   // "Prénom Nom" built from the swapped fields
    [InlineData("Jean Dupont", "Jean Dupont")]   // already reads right: kept
    [InlineData("ACME SA", "ACME SA")]           // company: kept
    [InlineData("", "")]
    public void FileAsFollowsTheSwap(string fileAs, string expected) =>
        Assert.Equal(expected, Plan.RecomputeFileAs(fileAs, oldFirst: "Dupont", oldLast: "Jean", newFirst: "Jean", newLast: "Dupont"));
}

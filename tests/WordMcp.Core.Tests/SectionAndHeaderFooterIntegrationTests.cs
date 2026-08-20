using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.HeaderFooter;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Section;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the section and header/footer commands. They use their own document
/// because section breaks and page setup changes are document-wide and would break the assertions
/// of the shared fixture.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class SectionAndHeaderFooterIntegrationTests : IDisposable
{
    private static readonly SectionCommands Sections = new();
    private static readonly HeaderFooterCommands HeadersFooters = new();
    private static readonly ParagraphCommands Paragraphs = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public SectionAndHeaderFooterIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-sechf-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "sections.docx")).SessionId;
    }

    [Fact]
    public void Section_ListReportsSingleSectionForNewDocument()
    {
        var list = Sections.List(Batch);

        Assert.Equal(1, list.TotalCount);
        Assert.Single(list.Sections);
        Assert.Equal(1, list.Sections[0].Index);
        Assert.False(list.Sections[0].DifferentFirstPage);
        Assert.True(list.Sections[0].PageSetup.PageWidth > 0);
        Assert.Equal("portrait", list.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void Section_AddCreatesSecondSection()
    {
        Paragraphs.Add(Batch, "First section body");

        var added = Sections.Add(Batch, "next-page");

        Assert.Equal(2, added.TotalCount);
        Assert.Equal("next-page", added.Section.StartType);

        var list = Sections.List(Batch);
        Assert.Equal(2, list.TotalCount);
        Assert.Equal([1, 2], list.Sections.Select(s => s.Index));
    }

    [Fact]
    public void Section_AddContinuousUsesContinuousStart()
    {
        Paragraphs.Add(Batch, "Body");

        var added = Sections.Add(Batch, "continuous");

        Assert.Equal(2, added.TotalCount);
        Assert.Equal("continuous", added.Section.StartType);
    }

    [Fact]
    public void Section_PageSetupAppliesToOneSectionOnly()
    {
        Paragraphs.Add(Batch, "Body");
        Sections.Add(Batch, "next-page");

        var before = Sections.List(Batch).Sections[0].PageSetup;

        var updated = Sections.PageSetup(
            Batch, sectionIndex: 2, topMargin: 100, bottomMargin: 90, orientation: "landscape");

        Assert.Equal(2, updated.Section.Index);
        Assert.Equal(100, updated.Section.PageSetup.TopMargin, 1);
        Assert.Equal(90, updated.Section.PageSetup.BottomMargin, 1);
        Assert.Equal("landscape", updated.Section.PageSetup.Orientation);

        // Landscape has to swap the page dimensions.
        Assert.True(updated.Section.PageSetup.PageWidth > updated.Section.PageSetup.PageHeight);

        var after = Sections.List(Batch).Sections[0].PageSetup;
        Assert.Equal(before.TopMargin, after.TopMargin, 1);
        Assert.Equal("portrait", after.Orientation);
    }

    [Fact]
    public void Section_PageSetupWithoutIndexAppliesToWholeDocument()
    {
        Paragraphs.Add(Batch, "Body");
        Sections.Add(Batch, "next-page");

        Sections.PageSetup(Batch, leftMargin: 120, rightMargin: 120, paperSize: "a4");

        var list = Sections.List(Batch);
        Assert.Equal(2, list.TotalCount);
        Assert.All(list.Sections, section =>
        {
            Assert.Equal(120, section.PageSetup.LeftMargin, 1);
            Assert.Equal(120, section.PageSetup.RightMargin, 1);
        });
    }

    [Fact]
    public void Section_PageSetupRejectsMissingSection()
        => Assert.Throws<ArgumentOutOfRangeException>(
            () => Sections.PageSetup(Batch, sectionIndex: 7, topMargin: 72));

    [Fact]
    public void HeaderFooter_SetGetAndClearHeader()
    {
        var set = HeadersFooters.Set(Batch, "Quarterly report", alignment: "center");

        Assert.Equal(1, set.UpdatedCount);
        Assert.Equal("Quarterly report", set.HeadersFooters[0].Text);
        Assert.Equal("header", set.HeadersFooters[0].Kind);
        Assert.Equal("primary", set.HeadersFooters[0].Type);
        Assert.True(set.HeadersFooters[0].IsActive);

        var read = HeadersFooters.Get(Batch);
        Assert.Equal("Quarterly report", read.HeadersFooters[0].Text);

        var cleared = HeadersFooters.Clear(Batch);
        Assert.Equal(string.Empty, cleared.HeadersFooters[0].Text);
    }

    [Fact]
    public void HeaderFooter_FooterIsIndependentOfHeader()
    {
        HeadersFooters.Set(Batch, "Header text");
        HeadersFooters.Set(Batch, "Page footer", kind: "footer");

        Assert.Equal("Header text", HeadersFooters.Get(Batch).HeadersFooters[0].Text);

        var footer = HeadersFooters.Get(Batch, kind: "footer");
        Assert.Equal("Page footer", footer.HeadersFooters[0].Text);
        Assert.Equal("footer", footer.HeadersFooters[0].Kind);
    }

    [Fact]
    public void HeaderFooter_SectionSpecificWriteBreaksTheLink()
    {
        Paragraphs.Add(Batch, "Body");
        Sections.Add(Batch, "next-page");

        HeadersFooters.Set(Batch, "Shared header");

        var both = HeadersFooters.Get(Batch);
        Assert.Equal(2, both.TotalCount);
        Assert.All(both.HeadersFooters, entry => Assert.Equal("Shared header", entry.Text));

        var second = HeadersFooters.Set(Batch, "Appendix header", sectionIndex: 2);

        Assert.Equal(1, second.UpdatedCount);
        Assert.Equal(2, second.HeadersFooters[0].SectionIndex);
        Assert.False(second.HeadersFooters[0].LinkedToPrevious);

        var after = HeadersFooters.Get(Batch);
        Assert.Equal("Shared header", after.HeadersFooters[0].Text);
        Assert.Equal("Appendix header", after.HeadersFooters[1].Text);
    }

    [Fact]
    public void HeaderFooter_FirstPageTypeIsActivatedOnWrite()
    {
        Assert.False(Sections.List(Batch).Sections[0].DifferentFirstPage);

        var set = HeadersFooters.Set(Batch, "Cover page", type: "first-page");

        Assert.True(set.HeadersFooters[0].IsActive);
        Assert.Equal("first-page", set.HeadersFooters[0].Type);

        // Writing the first-page header has to flip the section switch, otherwise Word stores the
        // text but never renders it.
        Assert.True(Sections.List(Batch).Sections[0].DifferentFirstPage);

        var primary = HeadersFooters.Get(Batch);
        Assert.Equal(string.Empty, primary.HeadersFooters[0].Text);
    }

    [Fact]
    public void HeaderFooter_EvenPagesTypeIsActivatedOnWrite()
    {
        HeadersFooters.Set(Batch, "Even footer", kind: "footer", type: "even-pages");

        var list = Sections.List(Batch);
        Assert.True(list.Sections[0].DifferentOddEvenPages);

        var read = HeadersFooters.Get(Batch, kind: "footer", type: "even-pages");
        Assert.Equal("Even footer", read.HeadersFooters[0].Text);
        Assert.True(read.HeadersFooters[0].IsActive);
    }

    [Fact]
    public void HeaderFooter_RejectsMissingSection()
        => Assert.Throws<ArgumentOutOfRangeException>(() => HeadersFooters.Get(Batch, sectionIndex: 9));

    public void Dispose()
    {
        _sessions.Dispose();

        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
            // A locked temp file must not fail the test run.
        }
    }
}

using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Footnote;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Text;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the footnote commands against a real Word instance.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class FootnoteIntegrationTests : IDisposable
{
    private static readonly FootnoteCommands Footnotes = new();
    private static readonly ParagraphCommands Paragraphs = new();
    private static readonly TextCommands Texts = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public FootnoteIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-footnote-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "footnotes.docx")).SessionId;
    }

    [Fact]
    public void Footnote_AddListSetTextAndDeleteRoundTrip()
    {
        int index = Paragraphs.Add(Batch, "Revenue grew by fifteen percent.").Paragraph!.Index;

        var added = Footnotes.Add(Batch, index, "Annual report 2025, page 14.");

        Assert.True(added.Success);
        Assert.Equal(1, added.Note!.Index);
        Assert.Equal("footnote", added.Note.Kind);
        Assert.Equal(index, added.Note.ParagraphIndex);
        Assert.Contains("Annual report 2025", added.Note.Text, StringComparison.Ordinal);
        Assert.Equal(1, added.TotalCount);

        var list = Footnotes.List(Batch);
        Assert.Equal(1, list.TotalCount);
        Assert.Contains(list.Notes, n => n.Text.Contains("Annual report 2025", StringComparison.Ordinal));

        var changed = Footnotes.SetText(Batch, 1, "Quarterly report Q4, page 3.");
        Assert.Contains("Quarterly report Q4", changed.Note!.Text, StringComparison.Ordinal);

        var deleted = Footnotes.Delete(Batch, 1);
        Assert.Equal(0, deleted.TotalCount);
        Assert.Empty(Footnotes.List(Batch).Notes);
    }

    [Fact]
    public void Footnote_TextGetDoesNotContainTheNoteText()
    {
        int index = Paragraphs.Add(Batch, "The body sentence stays visible.").Paragraph!.Index;
        Footnotes.Add(Batch, index, "Only reachable through the footnote tool.");

        string body = Texts.Get(Batch).Text;

        Assert.Contains("The body sentence stays visible.", body, StringComparison.Ordinal);
        Assert.DoesNotContain("Only reachable through the footnote tool.", body, StringComparison.Ordinal);
    }

    [Fact]
    public void Footnote_EndnotesAreASeparateCollection()
    {
        int index = Paragraphs.Add(Batch, "A claim that needs two sources.").Paragraph!.Index;

        Footnotes.Add(Batch, index, "Footnote source.");
        Footnotes.Add(Batch, index, "Endnote source.", "endnote");

        var footnotes = Footnotes.List(Batch);
        var endnotes = Footnotes.List(Batch, "endnote");

        Assert.Equal(1, footnotes.TotalCount);
        Assert.Equal(1, endnotes.TotalCount);
        Assert.Equal("endnote", endnotes.Notes[0].Kind);
        Assert.Contains("Endnote source.", endnotes.Notes[0].Text, StringComparison.Ordinal);
        Assert.DoesNotContain(footnotes.Notes, n => n.Text.Contains("Endnote source.", StringComparison.Ordinal));
    }

    [Fact]
    public void Footnote_AnchorTextPlacesTheMarkAfterThePhrase()
    {
        int index = Paragraphs.Add(Batch, "Growth was strong although costs rose.").Paragraph!.Index;

        Footnotes.Add(Batch, index, "Measured year over year.", "footnote", "strong");

        var note = Footnotes.List(Batch).Notes[0];
        Assert.Equal(index, note.ParagraphIndex);
    }

    [Fact]
    public void Footnote_AnchorTextThatIsAbsentIsRejected()
    {
        int index = Paragraphs.Add(Batch, "A short paragraph.").Paragraph!.Index;

        Assert.Throws<ArgumentException>(() => Footnotes.Add(Batch, index, "Note.", "footnote", "missing phrase"));
    }

    [Fact]
    public void Footnote_UnknownIndexIsRejected()
    {
        Paragraphs.Add(Batch, "One paragraph.");

        Assert.Throws<ArgumentOutOfRangeException>(() => Footnotes.SetText(Batch, 3, "Note."));
        Assert.Throws<ArgumentOutOfRangeException>(() => Footnotes.Delete(Batch, 1));
    }

    [Fact]
    public void Footnote_ParagraphBeyondTheDocumentIsRejected()
    {
        Paragraphs.Add(Batch, "One paragraph.");

        Assert.Throws<ArgumentOutOfRangeException>(() => Footnotes.Add(Batch, 999, "Note."));
    }

    [Fact]
    public void Footnote_DeletingRenumbersTheRest()
    {
        int index = Paragraphs.Add(Batch, "A paragraph carrying three notes.").Paragraph!.Index;

        Footnotes.Add(Batch, index, "First note.");
        Footnotes.Add(Batch, index, "Second note.");
        Footnotes.Add(Batch, index, "Third note.");

        Assert.Equal(3, Footnotes.List(Batch).TotalCount);

        Footnotes.Delete(Batch, 1);

        var left = Footnotes.List(Batch);
        Assert.Equal(2, left.TotalCount);
        Assert.Equal([1, 2], left.Notes.Select(n => n.Index).ToArray());
        Assert.DoesNotContain(left.Notes, n => n.Text.Contains("First note.", StringComparison.Ordinal));
    }

    [Fact]
    public void Footnote_ListShortensLongNoteText()
    {
        int index = Paragraphs.Add(Batch, "A paragraph with a long note.").Paragraph!.Index;
        Footnotes.Add(Batch, index, new string('x', 300));

        Assert.Equal(20, Footnotes.List(Batch, "footnote", 20).Notes[0].Text.Length);
    }

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

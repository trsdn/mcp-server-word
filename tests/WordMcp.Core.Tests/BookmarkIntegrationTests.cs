using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Bookmark;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Text;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the bookmark commands against a real Word instance.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class BookmarkIntegrationTests : IDisposable
{
    private static readonly BookmarkCommands Bookmarks = new();
    private static readonly ParagraphCommands Paragraphs = new();
    private static readonly TextCommands Texts = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public BookmarkIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-bookmark-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "bookmarks.docx")).SessionId;
    }

    [Fact]
    public void Bookmark_AddListGetTextAndDeleteRoundTrip()
    {
        int index = Paragraphs.Add(Batch, "The introduction explains the scope.").Paragraph!.Index;

        var added = Bookmarks.Add(Batch, "Intro", index);

        Assert.Equal("Intro", added.Bookmark!.Name);
        Assert.Equal(index, added.Bookmark.ParagraphIndex);
        Assert.False(added.Bookmark.Empty);
        Assert.Equal(1, added.TotalCount);

        var list = Bookmarks.List(Batch);
        Assert.Contains(list.Bookmarks, b => b.Name == "Intro");

        var text = Bookmarks.GetText(Batch, "Intro");
        Assert.Equal("The introduction explains the scope.", text.Text);
        Assert.Equal(text.Text.Length, text.Length);

        var deleted = Bookmarks.Delete(Batch, "Intro");
        Assert.Equal(0, deleted.TotalCount);
        Assert.Empty(Bookmarks.List(Batch).Bookmarks);
    }

    [Fact]
    public void Bookmark_AnchorTextNarrowsTheBookmarkToAPhrase()
    {
        int index = Paragraphs.Add(Batch, "Revenue grew by fifteen percent last year.").Paragraph!.Index;

        Bookmarks.Add(Batch, "Growth", index, null, "fifteen percent");

        Assert.Equal("fifteen percent", Bookmarks.GetText(Batch, "Growth").Text);
    }

    [Fact]
    public void Bookmark_SpansSeveralParagraphs()
    {
        int first = Paragraphs.Add(Batch, "First paragraph.").Paragraph!.Index;
        Paragraphs.Add(Batch, "Second paragraph.");
        int last = Paragraphs.Add(Batch, "Third paragraph.").Paragraph!.Index;

        Bookmarks.Add(Batch, "Block", first, last);

        string text = Bookmarks.GetText(Batch, "Block").Text;
        Assert.Contains("First paragraph.", text, StringComparison.Ordinal);
        Assert.Contains("Third paragraph.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Bookmark_SurvivesAnEditInAnEarlierParagraph()
    {
        Paragraphs.Add(Batch, "Alpha paragraph.");
        int index = Paragraphs.Add(Batch, "The bookmarked passage.").Paragraph!.Index;
        Bookmarks.Add(Batch, "Passage", index);

        Texts.Replace(Batch, "Alpha", "Alpha and a good deal more text");

        Assert.Equal("The bookmarked passage.", Bookmarks.GetText(Batch, "Passage").Text);
    }

    [Fact]
    public void Bookmark_AddRejectsADuplicateName()
    {
        int index = Paragraphs.Add(Batch, "Some paragraph.").Paragraph!.Index;
        Bookmarks.Add(Batch, "Twice", index);

        Assert.Throws<ArgumentException>(() => Bookmarks.Add(Batch, "Twice", index));
    }

    [Fact]
    public void Bookmark_AddRejectsAMissingParagraph()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Bookmarks.Add(Batch, "Nowhere", 500));

    [Fact]
    public void Bookmark_AddRejectsAnchorTextThatIsNotInTheParagraph()
    {
        int index = Paragraphs.Add(Batch, "A short sentence.").Paragraph!.Index;

        Assert.Throws<ArgumentException>(() => Bookmarks.Add(Batch, "Missing", index, null, "not present"));
    }

    [Fact]
    public void Bookmark_GetTextRejectsAnUnknownName()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Bookmarks.GetText(Batch, "Unknown"));

    [Fact]
    public void Bookmark_DeleteRejectsAnUnknownName()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Bookmarks.Delete(Batch, "Unknown"));

    [Fact]
    public void Bookmark_ListShortensThePreview()
    {
        int index = Paragraphs.Add(Batch, new string('x', 300)).Paragraph!.Index;
        Bookmarks.Add(Batch, "Long", index);

        var info = Assert.Single(Bookmarks.List(Batch, maxTextLength: 20).Bookmarks);

        Assert.Equal(20, info.Text.Length);
        Assert.Equal(300, Bookmarks.GetText(Batch, "Long").Length);
    }

    [Fact]
    public void Bookmark_DeleteKeepsTheText()
    {
        int index = Paragraphs.Add(Batch, "Text that must stay.").Paragraph!.Index;
        Bookmarks.Add(Batch, "Keeper", index);

        Bookmarks.Delete(Batch, "Keeper");

        Assert.Contains("Text that must stay.", Texts.Get(Batch).Text, StringComparison.Ordinal);
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

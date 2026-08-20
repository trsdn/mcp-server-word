using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.List;
using WordMcp.Core.Commands.Paragraph;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the list commands. They use their own document because list numbering is
/// document-wide state that would leak between assertions of the shared fixture.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class ListIntegrationTests : IDisposable
{
    private static readonly ListCommands Lists = new();
    private static readonly ParagraphCommands Paragraphs = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public ListIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-list-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "lists.docx")).SessionId;
    }

    /// <summary>
    /// Adds paragraphs and returns the 1-based index of the first one.
    /// </summary>
    private int AddParagraphs(params string[] texts)
    {
        int first = 0;

        foreach (string text in texts)
        {
            var result = Paragraphs.Add(Batch, text);
            first = first == 0 ? result.Paragraph!.Index : first;
        }

        return first;
    }

    [Fact]
    public void List_GetReportsNothingForDocumentWithoutLists()
    {
        AddParagraphs("Plain text");

        var lists = Lists.Get(Batch);

        Assert.Empty(lists.Paragraphs);
        Assert.True(lists.TotalCount > 0);
    }

    [Fact]
    public void List_GetIncludesUnlistedParagraphsOnRequest()
    {
        AddParagraphs("Plain text");

        var all = Lists.Get(Batch, listedOnly: false);

        Assert.NotEmpty(all.Paragraphs);
        Assert.Contains(all.Paragraphs, p => p.ListType == "none" && p.Level == 0);
    }

    [Fact]
    public void List_ApplyBulletsFormatsTheRange()
    {
        int first = AddParagraphs("Apples", "Pears", "Plums");

        var applied = Lists.Apply(Batch, first, first + 2);

        Assert.Equal(3, applied.UpdatedCount);
        Assert.All(applied.Paragraphs, p => Assert.Equal("bullet", p.ListType));
        Assert.All(applied.Paragraphs, p => Assert.Equal(1, p.Level));
        Assert.All(applied.Paragraphs, p => Assert.NotEmpty(p.ListLabel));
    }

    [Fact]
    public void List_ApplyNumbersProducesAscendingLabels()
    {
        int first = AddParagraphs("First", "Second", "Third");

        var applied = Lists.Apply(Batch, first, first + 2, "number");

        Assert.All(applied.Paragraphs, p => Assert.Equal("number", p.ListType));
        Assert.StartsWith("1", applied.Paragraphs[0].ListLabel, StringComparison.Ordinal);
        Assert.StartsWith("2", applied.Paragraphs[1].ListLabel, StringComparison.Ordinal);
        Assert.StartsWith("3", applied.Paragraphs[2].ListLabel, StringComparison.Ordinal);
    }

    [Fact]
    public void List_ApplyWithoutEndIndexTouchesOneParagraph()
    {
        int first = AddParagraphs("Only this one", "Not this one");

        var applied = Lists.Apply(Batch, first);

        Assert.Equal(1, applied.UpdatedCount);
        Assert.Equal("bullet", applied.Paragraphs[0].ListType);
    }

    [Fact]
    public void List_SetLevelIndentsParagraphsOfAnOutlineList()
    {
        int first = AddParagraphs("Chapter", "Detail", "Detail two", "Chapter two");
        Lists.Apply(Batch, first, first + 3, "outline-number");

        var nested = Lists.SetLevel(Batch, first + 1, 2, first + 2);

        Assert.Equal(2, nested.UpdatedCount);
        Assert.All(nested.Paragraphs, p => Assert.Equal(2, p.Level));

        var all = Lists.Get(Batch);
        Assert.Equal(1, all.Paragraphs.Single(p => p.Index == first).Level);
        Assert.Equal(1, all.Paragraphs.Single(p => p.Index == first + 3).Level);
    }

    [Fact]
    public void List_SetLevelRejectsParagraphOutsideAList()
    {
        int first = AddParagraphs("Not a list item");

        Assert.Throws<ArgumentException>(() => Lists.SetLevel(Batch, first, 2));
    }

    [Fact]
    public void List_RestartBeginsNumberingAtOneAgain()
    {
        int first = AddParagraphs("One", "Two", "Three", "Four");
        Lists.Apply(Batch, first, first + 3, "number");

        var restarted = Lists.Restart(Batch, first + 2);

        Assert.StartsWith("1", restarted.Paragraphs[0].ListLabel, StringComparison.Ordinal);

        var all = Lists.Get(Batch);
        Assert.StartsWith("1", all.Paragraphs.Single(p => p.Index == first).ListLabel, StringComparison.Ordinal);
        Assert.StartsWith("2", all.Paragraphs.Single(p => p.Index == first + 1).ListLabel, StringComparison.Ordinal);
        Assert.StartsWith("1", all.Paragraphs.Single(p => p.Index == first + 2).ListLabel, StringComparison.Ordinal);
        Assert.StartsWith("2", all.Paragraphs.Single(p => p.Index == first + 3).ListLabel, StringComparison.Ordinal);
    }

    [Fact]
    public void List_RestartRejectsParagraphOutsideAList()
    {
        int first = AddParagraphs("Not a list item");

        Assert.Throws<ArgumentException>(() => Lists.Restart(Batch, first));
    }

    [Fact]
    public void List_RemoveStripsTheFormatting()
    {
        int first = AddParagraphs("Alpha", "Beta");
        Lists.Apply(Batch, first, first + 1, "number");

        var removed = Lists.Remove(Batch, first, first + 1);

        Assert.Equal(2, removed.UpdatedCount);
        Assert.All(removed.Paragraphs, p => Assert.Equal("none", p.ListType));
        Assert.All(removed.Paragraphs, p => Assert.Empty(p.ListLabel));

        Assert.DoesNotContain(Lists.Get(Batch).Paragraphs, p => p.Index == first);
    }

    [Fact]
    public void List_RejectsRangeBeyondTheDocument()
    {
        int first = AddParagraphs("Only paragraph");

        Assert.Throws<ArgumentOutOfRangeException>(() => Lists.Apply(Batch, first, first + 50));
    }

    [Fact]
    public void List_RejectsEndIndexBeforeStartIndex()
    {
        int first = AddParagraphs("A", "B");

        Assert.Throws<ArgumentOutOfRangeException>(() => Lists.Apply(Batch, first + 1, first));
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

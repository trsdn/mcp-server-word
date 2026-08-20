using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Comment;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Revision;
using WordMcp.Core.Commands.Text;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the comment and revision commands. They use their own document because
/// change tracking is document-wide state.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class CommentAndRevisionIntegrationTests : IDisposable
{
    private static readonly CommentCommands Comments = new();
    private static readonly RevisionCommands Revisions = new();
    private static readonly ParagraphCommands Paragraphs = new();
    private static readonly TextCommands Texts = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public CommentAndRevisionIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-review-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "review.docx")).SessionId;
    }

    [Fact]
    public void Comment_AddListAndDeleteRoundTrip()
    {
        int index = Paragraphs.Add(Batch, "The quick brown fox jumps.").Paragraph!.Index;

        var added = Comments.Add(Batch, index, "Please shorten this sentence.");

        Assert.Equal(1, added.TotalCount);
        Assert.Equal("Please shorten this sentence.", added.Comment!.Text);
        Assert.Contains("quick brown fox", added.Comment.ScopeText, StringComparison.Ordinal);

        var list = Comments.List(Batch);
        Assert.Equal(1, list.TotalCount);
        Assert.Equal(index, list.Comments[0].ParagraphIndex);
        Assert.False(string.IsNullOrWhiteSpace(list.Comments[0].Author));

        var deleted = Comments.Delete(Batch, 1);
        Assert.Equal(0, deleted.TotalCount);
        Assert.Empty(Comments.List(Batch).Comments);
    }

    [Fact]
    public void Comment_AnchorTextNarrowsTheScope()
    {
        int index = Paragraphs.Add(Batch, "Revenue grew by fifteen percent last year.").Paragraph!.Index;

        var added = Comments.Add(Batch, index, "Source?", "fifteen percent");

        Assert.Equal("fifteen percent", added.Comment!.ScopeText);
    }

    [Fact]
    public void Comment_AddRejectsAnchorTextThatIsNotInTheParagraph()
    {
        int index = Paragraphs.Add(Batch, "A short sentence.").Paragraph!.Index;

        Assert.Throws<ArgumentException>(() => Comments.Add(Batch, index, "Note", "not present"));
    }

    [Fact]
    public void Comment_AddRejectsMissingParagraph()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Comments.Add(Batch, 500, "Note"));

    [Fact]
    public void Comment_ResolveEitherMarksTheCommentOrReportsItAsUnsupported()
    {
        int index = Paragraphs.Add(Batch, "Draft wording.").Paragraph!.Index;
        Comments.Add(Batch, index, "Needs review.");

        // Modern comments treat everything added through the API as an unposted draft and refuse to
        // resolve it, so both outcomes are correct as long as the failure is a clear message.
        try
        {
            var resolved = Comments.Resolve(Batch, 1);

            Assert.True(resolved.Comment!.Resolved);
            Assert.Empty(Comments.List(Batch, unresolvedOnly: true).Comments);

            var reopened = Comments.Resolve(Batch, 1, resolved: false);
            Assert.False(reopened.Comment!.Resolved);
            Assert.Single(Comments.List(Batch, unresolvedOnly: true).Comments);
        }
        catch (NotSupportedException ex)
        {
            Assert.Contains("Delete the comment instead", ex.Message, StringComparison.Ordinal);
            Assert.Single(Comments.List(Batch, unresolvedOnly: true).Comments);
        }
    }

    [Fact]
    public void Comment_RejectsMissingComment()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Comments.Delete(Batch, 7));

    [Fact]
    public void Revision_TrackingIsOffForANewDocument()
    {
        var list = Revisions.List(Batch);

        Assert.False(list.TrackingEnabled);
        Assert.Equal(0, list.TotalCount);
    }

    [Fact]
    public void Revision_EditsAreRecordedWhileTrackingIsOn()
    {
        Paragraphs.Add(Batch, "Original wording stays.");

        var enabled = Revisions.SetTracking(Batch, true);
        Assert.True(enabled.TrackingEnabled);

        Texts.Replace(Batch, "Original", "Revised");

        var list = Revisions.List(Batch);
        Assert.True(list.TrackingEnabled);
        Assert.True(list.TotalCount > 0);
        Assert.Contains(list.Revisions, r => r.Type is "insert" or "delete" or "replace");
        Assert.All(list.Revisions, r => Assert.False(string.IsNullOrWhiteSpace(r.Author)));
    }

    [Fact]
    public void Revision_AcceptAllClearsTheChanges()
    {
        Paragraphs.Add(Batch, "Original wording stays.");
        Revisions.SetTracking(Batch, true);
        Texts.Replace(Batch, "Original", "Revised");

        var accepted = Revisions.Accept(Batch);

        Assert.True(accepted.HandledCount > 0);
        Assert.Equal(0, accepted.TotalCount);
        Assert.Contains("Revised", Texts.Get(Batch).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Revision_RejectAllRestoresTheOriginalText()
    {
        Paragraphs.Add(Batch, "Original wording stays.");
        Revisions.SetTracking(Batch, true);
        Texts.Replace(Batch, "Original", "Revised");

        var rejected = Revisions.Reject(Batch);

        Assert.Equal(0, rejected.TotalCount);

        string text = Texts.Get(Batch).Text;
        Assert.Contains("Original", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Revised", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Revision_AcceptOneLeavesTheOthers()
    {
        Paragraphs.Add(Batch, "Alpha stays here.");
        Paragraphs.Add(Batch, "Beta stays here.");
        Revisions.SetTracking(Batch, true);
        Texts.Replace(Batch, "Alpha", "Gamma");
        Texts.Replace(Batch, "Beta", "Delta");

        int before = Revisions.List(Batch).TotalCount;
        Assert.True(before > 1);

        var accepted = Revisions.Accept(Batch, 1);

        Assert.Equal(1, accepted.HandledCount);
        Assert.Equal(before - 1, accepted.TotalCount);
    }

    [Fact]
    public void Revision_RejectsMissingRevision()
    {
        Revisions.SetTracking(Batch, true);

        Assert.Throws<ArgumentOutOfRangeException>(() => Revisions.Accept(Batch, 9));
    }

    [Fact]
    public void Revision_TrackingCanBeTurnedOffAgain()
    {
        Revisions.SetTracking(Batch, true);

        var disabled = Revisions.SetTracking(Batch, false);

        Assert.False(disabled.TrackingEnabled);
        Assert.False(Revisions.List(Batch).TrackingEnabled);
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

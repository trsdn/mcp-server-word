using WordMcp.Core.Commands.Comment;
using WordMcp.Core.Commands.Revision;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the comment and revision commands. Everything asserted here happens
/// before the batch is touched, so these tests need no Word installation and run in CI.
/// </summary>
public class CommentAndRevisionValidationTests
{
    private static readonly CommentCommands Comments = new();
    private static readonly RevisionCommands Revisions = new();
    private static readonly ThrowingBatch Batch = new();

    [Fact]
    public void Comment_AddRejectsParagraphIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Comments.Add(Batch, 0, "Note"));

    [Fact]
    public void Comment_AddRejectsEmptyText()
        => Assert.Throws<ArgumentException>(() => Comments.Add(Batch, 1, "   "));

    [Fact]
    public void Comment_AddReachesBatchForValidArguments()
        => Assert.Throws<NotSupportedException>(() => Comments.Add(Batch, 1, "Note", "anchor"));

    [Fact]
    public void Comment_ResolveRejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Comments.Resolve(Batch, 0));

    [Fact]
    public void Comment_DeleteRejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Comments.Delete(Batch, -3));

    [Fact]
    public void Comment_AllCommandsRejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Comments.List(null!));
        Assert.Throws<ArgumentNullException>(() => Comments.Add(null!, 1, "Note"));
        Assert.Throws<ArgumentNullException>(() => Comments.Resolve(null!, 1));
        Assert.Throws<ArgumentNullException>(() => Comments.Delete(null!, 1));
    }

    [Fact]
    public void Revision_AcceptRejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Revisions.Accept(Batch, 0));

    [Fact]
    public void Revision_RejectRejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Revisions.Reject(Batch, 0));

    [Fact]
    public void Revision_AcceptWithoutIndexReachesBatch()
        => Assert.Throws<NotSupportedException>(() => Revisions.Accept(Batch));

    [Fact]
    public void Revision_RejectWithoutIndexReachesBatch()
        => Assert.Throws<NotSupportedException>(() => Revisions.Reject(Batch));

    [Fact]
    public void Revision_SetTrackingReachesBatch()
        => Assert.Throws<NotSupportedException>(() => Revisions.SetTracking(Batch, true));

    [Fact]
    public void Revision_AllCommandsRejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Revisions.List(null!));
        Assert.Throws<ArgumentNullException>(() => Revisions.Accept(null!));
        Assert.Throws<ArgumentNullException>(() => Revisions.Reject(null!));
        Assert.Throws<ArgumentNullException>(() => Revisions.SetTracking(null!, true));
    }
}

using WordMcp.Core.Commands.Footnote;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the footnote commands. Everything asserted here happens before the batch
/// is touched, so these tests need no Word installation and run in CI.
/// </summary>
public class FootnoteValidationTests
{
    private static readonly FootnoteCommands Footnotes = new();
    private static readonly ThrowingBatch Batch = new();

    [Theory]
    [InlineData("footnote")]
    [InlineData("Footnote")]
    [InlineData(" FOOTNOTES ")]
    [InlineData("endnote")]
    [InlineData("Endnotes")]
    public void List_AcceptsBothNoteKinds(string kind)
        => Assert.Throws<NotSupportedException>(() => Footnotes.List(Batch, kind));

    [Theory]
    [InlineData("note")]
    [InlineData("comment")]
    [InlineData("foot-note")]
    public void List_RejectsUnknownNoteKinds(string kind)
        => Assert.Throws<ArgumentException>(() => Footnotes.List(Batch, kind));

    [Fact]
    public void List_RejectsEmptyNoteKind()
        => Assert.Throws<ArgumentException>(() => Footnotes.List(Batch, "   "));

    [Fact]
    public void List_RejectsMaxTextLengthBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Footnotes.List(Batch, "footnote", 0));

    [Fact]
    public void Add_RejectsParagraphIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Footnotes.Add(Batch, 0, "Source."));

    [Fact]
    public void Add_RejectsEmptyText()
        => Assert.Throws<ArgumentException>(() => Footnotes.Add(Batch, 1, "   "));

    [Fact]
    public void Add_AcceptsAValidCall()
        => Assert.Throws<NotSupportedException>(() => Footnotes.Add(Batch, 1, "Source.", "endnote", "phrase"));

    [Fact]
    public void SetText_RejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Footnotes.SetText(Batch, 0, "Source."));

    [Fact]
    public void SetText_RejectsNullText()
        => Assert.Throws<ArgumentNullException>(() => Footnotes.SetText(Batch, 1, null!));

    [Fact]
    public void SetText_AcceptsEmptyTextToClearANote()
        => Assert.Throws<NotSupportedException>(() => Footnotes.SetText(Batch, 1, string.Empty));

    [Fact]
    public void Delete_RejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Footnotes.Delete(Batch, 0));

    [Fact]
    public void Delete_RejectsUnknownNoteKind()
        => Assert.Throws<ArgumentException>(() => Footnotes.Delete(Batch, 1, "sidenote"));

    [Fact]
    public void AllCommandsRejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Footnotes.List(null!));
        Assert.Throws<ArgumentNullException>(() => Footnotes.Add(null!, 1, "Source."));
        Assert.Throws<ArgumentNullException>(() => Footnotes.SetText(null!, 1, "Source."));
        Assert.Throws<ArgumentNullException>(() => Footnotes.Delete(null!, 1));
    }
}

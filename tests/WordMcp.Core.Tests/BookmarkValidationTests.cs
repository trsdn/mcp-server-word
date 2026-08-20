using WordMcp.Core.Commands.Bookmark;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the bookmark commands. Everything asserted here happens before the batch
/// is touched, so these tests need no Word installation and run in CI.
/// </summary>
public class BookmarkValidationTests
{
    private static readonly BookmarkCommands Bookmarks = new();
    private static readonly ThrowingBatch Batch = new();

    [Theory]
    [InlineData("My Bookmark")]
    [InlineData("1Intro")]
    [InlineData("intro-section")]
    [InlineData("intro.section")]
    [InlineData("_intro")]
    [InlineData("Überschrift")]
    public void Add_RejectsNamesWordWouldReject(string name)
        => Assert.Throws<ArgumentException>(() => Bookmarks.Add(Batch, name, 1));

    [Theory]
    [InlineData("Intro")]
    [InlineData("section_2")]
    [InlineData("a")]
    public void Add_AcceptsValidNames(string name)
        => Assert.Throws<NotSupportedException>(() => Bookmarks.Add(Batch, name, 1));

    [Fact]
    public void Add_RejectsNamesLongerThanFortyCharacters()
        => Assert.Throws<ArgumentException>(() => Bookmarks.Add(Batch, new string('a', 41), 1));

    [Fact]
    public void Add_AcceptsANameOfExactlyFortyCharacters()
        => Assert.Throws<NotSupportedException>(() => Bookmarks.Add(Batch, new string('a', 40), 1));

    [Fact]
    public void Add_RejectsParagraphIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Bookmarks.Add(Batch, "Intro", 0));

    [Fact]
    public void Add_RejectsEndBeforeStart()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Bookmarks.Add(Batch, "Intro", 4, 2));

    [Fact]
    public void Add_RejectsAnchorTextCombinedWithAParagraphRange()
        => Assert.Throws<ArgumentException>(() => Bookmarks.Add(Batch, "Intro", 1, 3, "phrase"));

    [Fact]
    public void Add_AcceptsAnchorTextForASingleParagraph()
        => Assert.Throws<NotSupportedException>(() => Bookmarks.Add(Batch, "Intro", 1, null, "phrase"));

    [Fact]
    public void List_RejectsMaxTextLengthBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Bookmarks.List(Batch, 0));

    [Fact]
    public void GetText_RejectsInvalidNames()
        => Assert.Throws<ArgumentException>(() => Bookmarks.GetText(Batch, "no good"));

    [Fact]
    public void Delete_RejectsEmptyNames()
        => Assert.Throws<ArgumentException>(() => Bookmarks.Delete(Batch, "   "));

    [Fact]
    public void AllCommandsRejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Bookmarks.List(null!));
        Assert.Throws<ArgumentNullException>(() => Bookmarks.Add(null!, "Intro", 1));
        Assert.Throws<ArgumentNullException>(() => Bookmarks.GetText(null!, "Intro"));
        Assert.Throws<ArgumentNullException>(() => Bookmarks.Delete(null!, "Intro"));
    }
}

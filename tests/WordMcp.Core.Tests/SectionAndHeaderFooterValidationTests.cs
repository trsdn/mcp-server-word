using WordMcp.Core.Commands.HeaderFooter;
using WordMcp.Core.Commands.Section;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the section and header/footer commands. All of it runs before the batch
/// is touched, so these tests need no Word installation and run in CI.
/// </summary>
public class SectionAndHeaderFooterValidationTests
{
    private static readonly SectionCommands Sections = new();
    private static readonly HeaderFooterCommands HeadersFooters = new();
    private static readonly ThrowingBatch Batch = new();

    [Fact]
    public void Add_RejectsUnknownStartType()
        => Assert.Throws<ArgumentException>(() => Sections.Add(Batch, "sideways"));

    [Fact]
    public void Add_RejectsNewColumnAsSectionBreak()
        => Assert.Throws<ArgumentException>(() => Sections.Add(Batch, "new-column"));

    [Theory]
    [InlineData("next-page")]
    [InlineData("continuous")]
    [InlineData("even-page")]
    [InlineData("odd-page")]
    public void Add_AcceptsKnownStartTypes(string startType)
        => Assert.Throws<NotSupportedException>(() => Sections.Add(Batch, startType));

    [Fact]
    public void Add_RejectsParagraphIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(
            () => Sections.Add(Batch, "next-page", paragraphIndex: 0));

    [Fact]
    public void PageSetup_RequiresAtLeastOneChange()
        => Assert.Throws<ArgumentException>(() => Sections.PageSetup(Batch));

    [Fact]
    public void PageSetup_RejectsNegativeMargins()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Sections.PageSetup(Batch, topMargin: -1));

    [Fact]
    public void PageSetup_RejectsSectionIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(
            () => Sections.PageSetup(Batch, sectionIndex: 0, topMargin: 72));

    [Fact]
    public void PageSetup_RejectsUnknownOrientation()
        => Assert.Throws<ArgumentException>(() => Sections.PageSetup(Batch, orientation: "sideways"));

    [Fact]
    public void PageSetup_RejectsUnknownPaperSize()
        => Assert.Throws<ArgumentException>(() => Sections.PageSetup(Batch, paperSize: "a2"));

    [Fact]
    public void PageSetup_ValidatesBeforeWriting()
    {
        // A valid request must reach the batch rather than being rejected.
        Assert.Throws<NotSupportedException>(
            () => Sections.PageSetup(Batch, topMargin: 72, orientation: "landscape", paperSize: "a4"));
    }

    [Theory]
    [InlineData("sideways")]
    [InlineData(" ")]
    public void HeaderFooter_RejectsUnknownKind(string kind)
        => Assert.Throws<ArgumentException>(() => HeadersFooters.Get(Batch, kind));

    [Fact]
    public void HeaderFooter_RejectsUnknownType()
        => Assert.Throws<ArgumentException>(() => HeadersFooters.Get(Batch, "header", type: "last-page"));

    [Fact]
    public void HeaderFooter_RejectsSectionIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(
            () => HeadersFooters.Get(Batch, "header", sectionIndex: 0));

    [Fact]
    public void HeaderFooter_RejectsNullTextOnSet()
        => Assert.Throws<ArgumentNullException>(() => HeadersFooters.Set(Batch, null!));

    [Fact]
    public void HeaderFooter_RejectsUnknownAlignment()
        => Assert.Throws<ArgumentException>(
            () => HeadersFooters.Set(Batch, "text", alignment: "diagonal"));

    [Theory]
    [InlineData("header", "primary")]
    [InlineData("footer", "first-page")]
    [InlineData("header", "even-pages")]
    public void HeaderFooter_AcceptsKnownCombinations(string kind, string type)
        => Assert.Throws<NotSupportedException>(() => HeadersFooters.Set(Batch, "text", kind, type: type));

    [Fact]
    public void Commands_RejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Sections.List(null!));
        Assert.Throws<ArgumentNullException>(() => HeadersFooters.Get(null!));
        Assert.Throws<ArgumentNullException>(() => HeadersFooters.Set(null!, "text"));
        Assert.Throws<ArgumentNullException>(() => HeadersFooters.Clear(null!));
    }
}

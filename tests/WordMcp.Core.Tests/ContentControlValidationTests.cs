using WordMcp.Core.Commands.ContentControl;
using WordMcp.Core.Utilities;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the content control commands. Everything asserted here happens before the
/// batch is touched, so these tests need no Word installation and run in CI.
/// </summary>
public class ContentControlValidationTests
{
    private static readonly ContentControlCommands Controls = new();
    private static readonly ThrowingBatch Batch = new();

    [Fact]
    public void List_WorksWithoutASelector()
        => Assert.Throws<NotSupportedException>(() => Controls.List(Batch));

    [Fact]
    public void List_RejectsMaxTextLengthBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Controls.List(Batch, null, 0));

    [Fact]
    public void Get_RequiresATagOrAnId()
        => Assert.Throws<ArgumentException>(() => Controls.Get(Batch));

    [Fact]
    public void Get_RejectsABlankTag()
        => Assert.Throws<ArgumentException>(() => Controls.Get(Batch, "   "));

    [Theory]
    [InlineData("CustomerName", null)]
    [InlineData(null, "123456789")]
    public void Get_AcceptsEitherSelector(string? tag, string? id)
        => Assert.Throws<NotSupportedException>(() => Controls.Get(Batch, tag, id));

    [Fact]
    public void SetText_RequiresASelector()
        => Assert.Throws<ArgumentException>(() => Controls.SetText(Batch, "Contoso"));

    [Fact]
    public void SetText_RequiresTextOrACheckState()
        => Assert.Throws<ArgumentException>(() => Controls.SetText(Batch, null, "CustomerName"));

    [Fact]
    public void SetText_AcceptsACheckStateWithoutText()
        => Assert.Throws<NotSupportedException>(() => Controls.SetText(Batch, null, "Agreed", null, true));

    [Fact]
    public void SetText_AcceptsEmptyTextToClearAControl()
        => Assert.Throws<NotSupportedException>(() => Controls.SetText(Batch, string.Empty, "CustomerName"));

    [Fact]
    public void Add_RejectsParagraphIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Controls.Add(Batch, 0));

    [Fact]
    public void Add_RejectsAnUnknownType()
        => Assert.Throws<ArgumentException>(() => Controls.Add(Batch, 1, "spinner"));

    [Theory]
    [InlineData("rich-text")]
    [InlineData("text")]
    [InlineData("Plain Text")]
    [InlineData("checkbox")]
    [InlineData("dropdown_list")]
    [InlineData("date")]
    public void Add_AcceptsTheDocumentedTypes(string type)
        => Assert.Throws<NotSupportedException>(() => Controls.Add(Batch, 1, type, "Notes", "Notes"));

    [Fact]
    public void Delete_RequiresASelector()
        => Assert.Throws<ArgumentException>(() => Controls.Delete(Batch));

    [Fact]
    public void Delete_AcceptsATag()
        => Assert.Throws<NotSupportedException>(() => Controls.Delete(Batch, "Draft", null, true));

    [Fact]
    public void TypeNamesRoundTrip()
    {
        foreach (string name in new[]
                 {
                     "rich-text", "text", "picture", "combo-box", "dropdown-list",
                     "building-block-gallery", "date", "group", "checkbox", "repeating-section"
                 })
        {
            Assert.Equal(name, WordConversions.FromWdContentControlType(WordConversions.ToWdContentControlType(name)));
        }
    }

    [Fact]
    public void AllCommandsRejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Controls.List(null!));
        Assert.Throws<ArgumentNullException>(() => Controls.Get(null!, "Tag"));
        Assert.Throws<ArgumentNullException>(() => Controls.SetText(null!, "Text", "Tag"));
        Assert.Throws<ArgumentNullException>(() => Controls.Add(null!, 1));
        Assert.Throws<ArgumentNullException>(() => Controls.Delete(null!, "Tag"));
    }
}

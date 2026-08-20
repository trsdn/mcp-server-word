using WordMcp.Core.Commands.Style;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the style commands. Everything asserted here happens before the batch is
/// touched, so these tests need no Word installation and run in CI.
/// </summary>
public class StyleValidationTests
{
    private static readonly StyleCommands Styles = new();
    private static readonly ThrowingBatch Batch = new();

    [Fact]
    public void List_RejectsUnknownStyleType()
        => Assert.Throws<ArgumentException>(() => Styles.List(Batch, styleType: "sideways"));

    [Theory]
    [InlineData("paragraph")]
    [InlineData("character")]
    [InlineData("table")]
    [InlineData("list")]
    public void List_AcceptsKnownStyleTypes(string styleType)
        => Assert.Throws<NotSupportedException>(() => Styles.List(Batch, styleType: styleType));

    [Fact]
    public void Create_RejectsEmptyName()
        => Assert.Throws<ArgumentException>(() => Styles.Create(Batch, "  "));

    [Fact]
    public void Create_RejectsUnknownStyleType()
        => Assert.Throws<ArgumentException>(() => Styles.Create(Batch, "Callout", "sideways"));

    [Fact]
    public void Create_DefaultsToParagraphStyle()
        => Assert.Throws<NotSupportedException>(() => Styles.Create(Batch, "Callout"));

    [Fact]
    public void Modify_RejectsEmptyName()
        => Assert.Throws<ArgumentException>(() => Styles.Modify(Batch, "", bold: true));

    [Fact]
    public void Modify_RejectsNonPositiveFontSize()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Styles.Modify(Batch, "Callout", fontSize: 0));

    [Fact]
    public void Modify_RejectsNonPositiveLineSpacing()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Styles.Modify(Batch, "Callout", lineSpacing: -1));

    [Fact]
    public void Modify_RejectsUnknownColor()
        => Assert.Throws<ArgumentException>(() => Styles.Modify(Batch, "Callout", color: "not-a-color"));

    [Fact]
    public void Modify_RejectsUnknownAlignment()
        => Assert.Throws<ArgumentException>(() => Styles.Modify(Batch, "Callout", alignment: "sideways"));

    [Fact]
    public void Modify_AcceptsValidFormatting()
        => Assert.Throws<NotSupportedException>(
            () => Styles.Modify(Batch, "Callout", color: "#C00000", alignment: "center", fontSize: 11));

    [Fact]
    public void Delete_RejectsEmptyName()
        => Assert.Throws<ArgumentException>(() => Styles.Delete(Batch, "   "));

    [Fact]
    public void Delete_ReachesBatchForValidName()
        => Assert.Throws<NotSupportedException>(() => Styles.Delete(Batch, "Callout"));

    [Fact]
    public void AllCommands_RejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Styles.List(null!));
        Assert.Throws<ArgumentNullException>(() => Styles.Create(null!, "Callout"));
        Assert.Throws<ArgumentNullException>(() => Styles.Modify(null!, "Callout"));
        Assert.Throws<ArgumentNullException>(() => Styles.Delete(null!, "Callout"));
    }
}

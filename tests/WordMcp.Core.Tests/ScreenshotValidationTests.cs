using WordMcp.Core.Commands.Screenshot;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the screenshot command. Everything asserted here happens before the
/// batch is touched, so these tests need no Word installation and run in CI.
/// </summary>
public class ScreenshotValidationTests
{
    private static readonly ScreenshotCommands Screenshots = new();
    private static readonly ThrowingBatch Batch = new();

    [Fact]
    public void Page_RejectsPageBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Screenshots.Page(Batch, 0));

    [Theory]
    [InlineData(35)]
    [InlineData(0)]
    [InlineData(-10)]
    public void Page_RejectsDpiBelowTheMinimum(int dpi)
        => Assert.Throws<ArgumentOutOfRangeException>(() => Screenshots.Page(Batch, 1, null, dpi));

    [Fact]
    public void Page_RejectsDpiAboveTheMaximum()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Screenshots.Page(Batch, 1, null, 601));

    [Theory]
    [InlineData(36)]
    [InlineData(150)]
    [InlineData(600)]
    public void Page_AcceptsDpiInsideTheRange(int dpi)
        => Assert.Throws<NotSupportedException>(() => Screenshots.Page(Batch, 1, null, dpi));

    [Fact]
    public void Page_RejectsNullBatch()
        => Assert.Throws<ArgumentNullException>(() => Screenshots.Page(null!));

    [Fact]
    public void Page_DoesNotCreateTheOutputDirectoryBeforeReachingWord()
    {
        string directory = Path.Combine(Path.GetTempPath(), "wordmcp-shot-" + Guid.NewGuid().ToString("N"));

        // The batch throws, so nothing is rendered - but the directory is prepared beforehand,
        // which is what makes a nested output path work.
        Assert.Throws<NotSupportedException>(
            () => Screenshots.Page(Batch, 1, Path.Combine(directory, "page.png")));

        Assert.True(Directory.Exists(directory));
        Directory.Delete(directory, recursive: true);
    }
}

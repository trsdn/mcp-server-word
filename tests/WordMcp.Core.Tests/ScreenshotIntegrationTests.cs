using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Paragraph;
using WordMcp.Core.Commands.Screenshot;
using WordMcp.Core.Commands.Text;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the screenshot command against a real Word instance.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class ScreenshotIntegrationTests : IDisposable
{
    private static readonly ScreenshotCommands Screenshots = new();
    private static readonly ParagraphCommands Paragraphs = new();
    private static readonly TextCommands Texts = new();

    private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public ScreenshotIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-screenshot-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "screenshot.docx")).SessionId;
    }

    [Fact]
    public void Screenshot_RendersTheFirstPageAsPng()
    {
        Paragraphs.Add(Batch, "A visible line of text.");
        string target = Path.Combine(_directory, "page1.png");

        var result = Screenshots.Page(Batch, 1, target);

        Assert.Equal(1, result.Page);
        Assert.True(result.PageCount >= 1);
        Assert.Equal(target, result.OutputPath);
        Assert.Equal(150, result.Dpi);
        Assert.True(result.Width > 500, $"Width was {result.Width}.");
        Assert.True(result.Height > result.Width, "A portrait page should be taller than it is wide.");
        Assert.True(result.FileSizeBytes > 0);
        Assert.Null(result.ImageBase64);

        Assert.True(File.Exists(target));
        Assert.Equal(PngSignature, File.ReadAllBytes(target).Take(PngSignature.Length).ToArray());
    }

    [Fact]
    public void Screenshot_WritesToATempFileWhenNoPathIsGiven()
    {
        Paragraphs.Add(Batch, "Anywhere will do.");

        var result = Screenshots.Page(Batch);

        try
        {
            Assert.True(File.Exists(result.OutputPath));
            Assert.EndsWith(".png", result.OutputPath, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(result.OutputPath);
        }
    }

    [Fact]
    public void Screenshot_HigherDpiProducesALargerImage()
    {
        Paragraphs.Add(Batch, "Resolution test.");

        var low = Screenshots.Page(Batch, 1, Path.Combine(_directory, "low.png"), dpi: 72);
        var high = Screenshots.Page(Batch, 1, Path.Combine(_directory, "high.png"), dpi: 300);

        Assert.True(high.Width > low.Width, $"{high.Width} should exceed {low.Width}.");
        Assert.True(high.Height > low.Height);
    }

    [Fact]
    public void Screenshot_ReturnsTheImageInlineWhenAsked()
    {
        Paragraphs.Add(Batch, "Inline test.");

        var result = Screenshots.Page(
            Batch, 1, Path.Combine(_directory, "inline.png"), dpi: 72, includeImage: true);

        Assert.False(string.IsNullOrEmpty(result.ImageBase64));

        byte[] decoded = Convert.FromBase64String(result.ImageBase64!);
        Assert.Equal(PngSignature, decoded.Take(PngSignature.Length).ToArray());
        Assert.Equal(result.FileSizeBytes, decoded.Length);
    }

    [Fact]
    public void Screenshot_RendersALaterPage()
    {
        Paragraphs.Add(Batch, "Page one content.");

        // Chr(12) is Word's page break character, which is the shortest way to force a second page.
        Texts.Append(Batch, "\f");
        Paragraphs.Add(Batch, "Page two content.");

        var result = Screenshots.Page(Batch, 2, Path.Combine(_directory, "page2.png"), dpi: 72);

        Assert.Equal(2, result.Page);
        Assert.True(result.PageCount >= 2, $"Page count was {result.PageCount}.");
        Assert.True(File.Exists(result.OutputPath));
    }

    [Fact]
    public void Screenshot_RejectsAPageThatDoesNotExist()
    {
        Paragraphs.Add(Batch, "Only one page here.");

        Assert.Throws<ArgumentOutOfRangeException>(() => Screenshots.Page(Batch, 99));
    }

    [Fact]
    public void Screenshot_CreatesAMissingOutputDirectory()
    {
        Paragraphs.Add(Batch, "Nested output.");
        string target = Path.Combine(_directory, "nested", "deeper", "page.png");

        var result = Screenshots.Page(Batch, 1, target, dpi: 72);

        Assert.True(File.Exists(result.OutputPath));
    }

    [Fact]
    public void Screenshot_LeavesNoPdfBehind()
    {
        Paragraphs.Add(Batch, "Cleanup test.");
        var before = Directory.GetFiles(Path.GetTempPath(), "wordmcp-page-*.pdf").Length;

        Screenshots.Page(Batch, 1, Path.Combine(_directory, "cleanup.png"), dpi: 72);

        Assert.Equal(before, Directory.GetFiles(Path.GetTempPath(), "wordmcp-page-*.pdf").Length);
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

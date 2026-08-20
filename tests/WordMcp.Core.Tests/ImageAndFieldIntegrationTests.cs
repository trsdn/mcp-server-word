using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Field;
using WordMcp.Core.Commands.Image;
using WordMcp.Core.Commands.Paragraph;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Integration tests for the image and field commands. They run against their own document
/// because both add structure (pictures, a table of contents, page numbers) that would interfere
/// with the assertions of the shared fixture.
/// </summary>
[Trait("Category", "RequiresWord")]
public sealed class ImageAndFieldIntegrationTests : IDisposable
{
    private static readonly ImageCommands Images = new();
    private static readonly FieldCommands Fields = new();
    private static readonly ParagraphCommands Paragraphs = new();

    private readonly SessionManager _sessions = new();
    private readonly string _directory;
    private readonly string _sessionId;

    private IWordBatch Batch => _sessions.GetBatch(_sessionId);

    public ImageAndFieldIntegrationTests()
    {
        _directory = Directory
            .CreateDirectory(Path.Combine(Path.GetTempPath(), "wordmcp-imgfld-" + Guid.NewGuid().ToString("N")))
            .FullName;

        _sessionId = _sessions.Create(Path.Combine(_directory, "images.docx")).SessionId;
    }

    [Fact]
    public void Image_InsertListResizeReplaceAndDelete()
    {
        string red = CreateImage("red.bmp", 0x00, 0x00, 0xFF, 40, 20);
        string blue = CreateImage("blue.bmp", 0xFF, 0x00, 0x00, 40, 20);

        var inserted = Images.Insert(Batch, red, width: 120, caption: "Figure 1", altText: "A red rectangle");

        Assert.True(inserted.Success);
        Assert.NotNull(inserted.Image);
        Assert.Equal(1, inserted.TotalCount);
        Assert.Equal("A red rectangle", inserted.Image!.AltText);
        Assert.Equal(120, inserted.Image.Width, 1);

        // A locked aspect ratio has to scale the height along with the width.
        Assert.Equal(60, inserted.Image.Height, 1);

        var list = Images.List(Batch);
        Assert.Equal(1, list.TotalCount);
        Assert.False(list.Images[0].IsLinked);

        var resized = Images.Resize(Batch, 1, scalePercent: 50);
        Assert.Equal(60, resized.Image!.Width, 1);
        Assert.Equal(30, resized.Image.Height, 1);

        var replaced = Images.Replace(Batch, 1, blue);
        Assert.Equal(60, replaced.Image!.Width, 1);
        Assert.Equal(30, replaced.Image.Height, 1);
        Assert.Equal("A red rectangle", replaced.Image.AltText);

        var relabelled = Images.SetAltText(Batch, 1, "A blue rectangle");
        Assert.Equal("A blue rectangle", relabelled.Image!.AltText);

        var deleted = Images.Delete(Batch, 1);
        Assert.Equal(0, deleted.TotalCount);
    }

    [Fact]
    public void Image_ResizeWithoutAspectRatioUsesBothDimensions()
    {
        string image = CreateImage("free.bmp", 0x00, 0xFF, 0x00, 40, 20);
        Images.Insert(Batch, image);

        var resized = Images.Resize(Batch, 1, width: 200, height: 50, lockAspectRatio: false);

        Assert.Equal(200, resized.Image!.Width, 1);
        Assert.Equal(50, resized.Image.Height, 1);
    }

    [Fact]
    public void Image_RejectsMissingFileAndUnsupportedFormat()
    {
        Assert.Throws<FileNotFoundException>(
            () => Images.Insert(Batch, Path.Combine(_directory, "nope.png")));

        string text = Path.Combine(_directory, "not-an-image.txt");
        File.WriteAllText(text, "hello");
        Assert.Throws<ArgumentException>(() => Images.Insert(Batch, text));
    }

    [Fact]
    public void Image_RejectsOutOfRangeIndex()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => Images.Delete(Batch, 99));
    }

    [Fact]
    public void Field_TableOfContentsIsEmptyWithoutHeadings()
    {
        Paragraphs.Add(Batch, "Just body text, no heading style.");

        var toc = Fields.InsertTableOfContents(Batch);

        Assert.True(toc.Success);
        Assert.Equal(0, toc.EntryCount);
        Assert.Contains("heading style", toc.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Field_TableOfContentsPicksUpHeadings()
    {
        Paragraphs.Add(Batch, "First chapter", style: "Heading 1");
        Paragraphs.Add(Batch, "Some body text.");
        Paragraphs.Add(Batch, "A section", style: "Heading 2");
        Paragraphs.Add(Batch, "More body text.");

        var toc = Fields.InsertTableOfContents(Batch, lowerHeadingLevel: 2);

        Assert.True(toc.Success);
        Assert.Equal(2, toc.EntryCount);

        var updated = Fields.UpdateTableOfContents(Batch);
        Assert.Equal(1, updated.UpdatedCount);
        Assert.Equal(2, updated.EntryCount);

        var fields = Fields.List(Batch);
        Assert.Contains(fields.Fields, f => f.Code.Contains("TOC", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Field_RejectsInvertedHeadingLevels()
    {
        Assert.Throws<ArgumentException>(
            () => Fields.InsertTableOfContents(Batch, upperHeadingLevel: 3, lowerHeadingLevel: 1));
    }

    [Fact]
    public void Field_InsertsPageNumberAndUpdatesAll()
    {
        var inserted = Fields.InsertPageNumber(Batch, includeTotalPages: true);

        Assert.True(inserted.Success);
        Assert.True(inserted.UpdatedCount >= 1);

        // Header and footer fields are invisible to Document.Fields, so update-all has to walk
        // the sections to reach them.
        var updated = Fields.UpdateAll(Batch);
        Assert.True(updated.UpdatedCount >= 2);
    }

    [Fact]
    public void Field_RejectsUnknownPosition()
    {
        Assert.Throws<ArgumentException>(() => Fields.InsertPageNumber(Batch, position: "sidebar"));
    }

    /// <summary>
    /// Writes an uncompressed 24-bit BMP so the tests do not need an image fixture on disk.
    /// </summary>
    private string CreateImage(string name, byte b, byte g, byte r, int width, int height)
    {
        string path = Path.Combine(_directory, name);

        int rowSize = ((width * 3) + 3) / 4 * 4;
        int pixelDataSize = rowSize * height;
        const int headerSize = 54;

        var buffer = new byte[headerSize + pixelDataSize];

        buffer[0] = (byte)'B';
        buffer[1] = (byte)'M';
        WriteInt32(buffer, 2, buffer.Length);
        WriteInt32(buffer, 10, headerSize);
        WriteInt32(buffer, 14, 40);
        WriteInt32(buffer, 18, width);
        WriteInt32(buffer, 22, height);
        buffer[26] = 1;
        buffer[28] = 24;
        WriteInt32(buffer, 34, pixelDataSize);
        WriteInt32(buffer, 38, 2835);
        WriteInt32(buffer, 42, 2835);

        for (int y = 0; y < height; y++)
        {
            int offset = headerSize + (y * rowSize);
            for (int x = 0; x < width; x++)
            {
                buffer[offset + (x * 3)] = b;
                buffer[offset + (x * 3) + 1] = g;
                buffer[offset + (x * 3) + 2] = r;
            }
        }

        File.WriteAllBytes(path, buffer);
        return path;
    }

    private static void WriteInt32(byte[] buffer, int offset, int value)
        => BitConverter.GetBytes(value).CopyTo(buffer, offset);

    public void Dispose()
    {
        _sessions.Dispose();

        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}

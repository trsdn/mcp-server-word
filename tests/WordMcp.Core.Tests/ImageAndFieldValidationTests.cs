using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using WordMcp.ComInterop.Session;
using WordMcp.Core.Commands.Field;
using WordMcp.Core.Commands.Image;
using Xunit;

namespace WordMcp.Core.Tests;

/// <summary>
/// Argument validation of the image and field commands. All of it runs before the batch is
/// touched, so these tests need no Word installation and run in CI.
/// </summary>
public class ImageAndFieldValidationTests
{
    private static readonly ImageCommands Images = new();
    private static readonly FieldCommands Fields = new();
    private static readonly ThrowingBatch Batch = new();

    [Fact]
    public void Insert_RejectsMissingFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "wordmcp-does-not-exist-" + Guid.NewGuid().ToString("N") + ".png");

        Assert.Throws<FileNotFoundException>(() => Images.Insert(Batch, path));
    }

    [Fact]
    public void Insert_RejectsUnsupportedExtension()
    {
        string path = Path.Combine(Path.GetTempPath(), "wordmcp-" + Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, "not an image");

        try
        {
            var ex = Assert.Throws<ArgumentException>(() => Images.Insert(Batch, path));
            Assert.Contains(".txt", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Insert_RejectsEmptyPath()
        => Assert.Throws<ArgumentException>(() => Images.Insert(Batch, "  "));

    [Fact]
    public void Insert_RejectsParagraphIndexBelowOne()
    {
        string path = CreateSupportedFile();

        try
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => Images.Insert(Batch, path, paragraphIndex: 0));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Resize_RequiresAtLeastOneDimension()
        => Assert.Throws<ArgumentException>(() => Images.Resize(Batch, 1));

    [Fact]
    public void Resize_RejectsNonPositiveScale()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Images.Resize(Batch, 1, scalePercent: 0));

    [Fact]
    public void Resize_RejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Images.Resize(Batch, 0, width: 100));

    [Fact]
    public void Delete_RejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Images.Delete(Batch, 0));

    [Fact]
    public void SetAltText_RejectsIndexBelowOne()
        => Assert.Throws<ArgumentOutOfRangeException>(() => Images.SetAltText(Batch, 0, "text"));

    [Theory]
    [InlineData(0, 3)]
    [InlineData(1, 10)]
    public void InsertToc_RejectsHeadingLevelsOutsideOneToNine(int upper, int lower)
        => Assert.Throws<ArgumentOutOfRangeException>(
            () => Fields.InsertTableOfContents(Batch, upperHeadingLevel: upper, lowerHeadingLevel: lower));

    [Fact]
    public void InsertToc_RejectsInvertedHeadingLevels()
        => Assert.Throws<ArgumentException>(
            () => Fields.InsertTableOfContents(Batch, upperHeadingLevel: 3, lowerHeadingLevel: 2));

    [Theory]
    [InlineData("sidebar")]
    [InlineData("")]
    public void InsertPageNumber_RejectsUnknownPosition(string position)
        => Assert.Throws<ArgumentException>(() => Fields.InsertPageNumber(Batch, position));

    [Theory]
    [InlineData("footer")]
    [InlineData("bottom")]
    [InlineData("header")]
    [InlineData("top")]
    public void InsertPageNumber_AcceptsKnownPositions(string position)
    {
        // Reaching the batch means validation passed — the fake batch then makes it fail loudly.
        Assert.Throws<NotSupportedException>(() => Fields.InsertPageNumber(Batch, position));
    }

    [Fact]
    public void Commands_RejectNullBatch()
    {
        Assert.Throws<ArgumentNullException>(() => Images.List(null!));
        Assert.Throws<ArgumentNullException>(() => Fields.List(null!));
    }

    private static string CreateSupportedFile()
    {
        string path = Path.Combine(Path.GetTempPath(), "wordmcp-" + Guid.NewGuid().ToString("N") + ".png");
        File.WriteAllBytes(path, [0x89, 0x50, 0x4E, 0x47]);
        return path;
    }

    /// <summary>
    /// Stands in for a Word batch. Any call means validation let the request through, which is
    /// exactly what these tests assert.
    /// </summary>
    private sealed class ThrowingBatch : IWordBatch
    {
        public string DocumentPath => "fake.docx";

        public ILogger Logger => NullLogger.Instance;

        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(30);

        public int? WordProcessId => null;

        public T Execute<T>(Func<WordContext, CancellationToken, T> operation, CancellationToken cancellationToken = default)
            => throw new NotSupportedException("Reached the batch.");

        public void Execute(Action<WordContext, CancellationToken> operation, CancellationToken cancellationToken = default)
            => throw new NotSupportedException("Reached the batch.");

        public void Save(CancellationToken cancellationToken = default)
            => throw new NotSupportedException("Reached the batch.");

        public bool IsWordProcessAlive() => true;

        public void Dispose()
        {
        }
    }
}

using System.IO.Compression;
using WordMcp.ComInterop;
using WordMcp.Core.Utilities;
using Xunit;

namespace WordMcp.Core.Tests;

public class WordStylesTests
{
    [Theory]
    [InlineData("Heading 1", -2)]
    [InlineData("heading 3", -4)]
    [InlineData("  Title  ", -63)]
    [InlineData("Table Grid", -155)]
    [InlineData("Normal", -1)]
    public void Resolve_MapsBuiltInNamesToLanguageIndependentIds(string name, int expected)
    {
        Assert.Equal(expected, WordStyles.Resolve(name));
    }

    [Theory]
    [InlineData("My Custom Style")]
    [InlineData("Überschrift 1")]
    public void Resolve_PassesThroughUnknownNames(string name)
    {
        Assert.Equal(name, WordStyles.Resolve(name));
    }

    [Fact]
    public void Resolve_RejectsNull()
    {
        Assert.Throws<ArgumentNullException>(() => WordStyles.Resolve(null!));
    }
}

public class EmptyDocumentFactoryTests : IDisposable
{
    private readonly string _directory =
        Path.Combine(Path.GetTempPath(), "wordmcp-factory-" + Guid.NewGuid().ToString("N"));

    public EmptyDocumentFactoryTests() => Directory.CreateDirectory(_directory);

    [Theory]
    [InlineData(false, "wordprocessingml.document.main+xml")]
    [InlineData(true, "macroEnabled.main+xml")]
    public void Create_WritesValidPackage(bool macroEnabled, string expectedContentType)
    {
        string path = Path.Combine(_directory, macroEnabled ? "test.docm" : "test.docx");

        EmptyDocumentFactory.Create(path, macroEnabled);

        using var archive = ZipFile.OpenRead(path);
        Assert.NotNull(archive.GetEntry("_rels/.rels"));
        Assert.NotNull(archive.GetEntry("word/document.xml"));

        ZipArchiveEntry types = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("[Content_Types].xml"));
        using var reader = new StreamReader(types.Open());
        Assert.Contains(expectedContentType, reader.ReadToEnd(), StringComparison.Ordinal);
    }

    [Fact]
    public void Create_OverwritesExistingFile()
    {
        string path = Path.Combine(_directory, "existing.docx");
        File.WriteAllText(path, "this is not a document");

        EmptyDocumentFactory.Create(path, isMacroEnabled: false);

        using var archive = ZipFile.OpenRead(path);
        Assert.NotNull(archive.GetEntry("word/document.xml"));
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
        try
        {
            Directory.Delete(_directory, true);
        }
        catch (IOException)
        {
        }
    }
}
